#!/usr/bin/env node
// Export Realtor.ca listings to CSV, Excel (.xlsx) or JSONL.
//
// Calls the yugenox/realtor-ca-property-scraper Actor on Apify for one or more Canadian cities,
// then writes one row per active listing: MLS number, price, beds, baths, size, address,
// coordinates, listing date, main photo, listing URL and the listing brokerage. The Actor returns
// the newest listings first.
//
// --append keeps one growing file and adds only MLS numbers it doesn't have yet, so a daily run
// tells you which listings are new since the last one.
//
// Run with --help for every option. Requires Node 20.6+; APIFY_TOKEN in the environment (or in
// .env, loaded with `node --env-file=.env`) for new runs and --dataset.

import { appendFileSync, existsSync, readFileSync, statSync, writeFileSync } from 'node:fs';

const ACTOR = 'yugenox/realtor-ca-property-scraper';

// The cities the Actor searches (its input schema, checked 2026-10-06). Anything else is
// skipped by the Actor, so the script stops before starting a run for it. --any-city skips
// this check, for when the Actor adds cities.
const CITIES = ['Toronto', 'Mississauga', 'Brampton', 'Vaughan', 'Markham', 'Oakville', 'Hamilton',
    'Ottawa', 'London', 'Vancouver', 'Surrey', 'Burnaby', 'Victoria', 'Calgary', 'Edmonton',
    'Montreal', 'Winnipeg', 'Halifax'];
const MAX_PER_CITY = 500; // the most listings one Realtor.ca search area returns

const HELP = `Export Realtor.ca listings to CSV, Excel (.xlsx) or JSONL.

Usage:
  node --env-file=.env listings.mjs "Toronto, ON" ["Ottawa, ON" ...] [options]
  node --env-file=.env listings.mjs --file cities.txt --out listings.xlsx
  node --env-file=.env listings.mjs "Calgary, AB" --append --out calgary.csv --new-out new-today.csv
  node --env-file=.env listings.mjs --dataset <datasetId> --out listings.xlsx
  node listings.mjs --from rows.json --out listings.xlsx

What to search
  <city>...             "Toronto, ON" or "Toronto". Supported cities:
                        ${CITIES.slice(0, 9).join(', ')},
                        ${CITIES.slice(9).join(', ')}.
  --file <path>         One city per line. Blank lines and lines starting with # are skipped.
  --type <sale|rent>    For sale (default) or for rent.
  --property <residential|commercial>
                        Residential (default) or commercial.
  --limit <n>           Listings per city (default 100, max ${MAX_PER_CITY}).
  --any-city            Don't check cities against the list above.

Output
  --out <path>          .csv (default listings.csv), .xlsx or .jsonl
  --append              Add only listings whose mls_number is not in --out yet (.csv or
                        .jsonl). Run it daily to collect new listings into one file.
  --new-out <path>      With --append: also write just this run's new listings here
                        (.csv, .xlsx or .jsonl), replacing the file.
  --no-agents           Leave the listing agents' names and phone numbers out (the
                        brokerage stays).
  --max-charge <usd>    Hard cost cap for the run, enforced by Apify (default 2).

Re-export without a new run
  --dataset <id>        Export the dataset of a run you already made (Apify Console >
                        Storage > Datasets). No new run starts.
  --from <file>         Convert a JSON or JSONL file of the Actor's rows (Console export)
                        or of this script's rows. Needs no token.

  -h, --help            Show this help.

Price: $0.001 per listing plus a start fee per run (Apify Store, checked 2026-10-06; the
Store page shows the current price). Not affiliated with CREA or REALTOR.ca.`;

// Column order for CSV, XLSX and JSONL.
const COLUMNS = ['mls_number', 'price', 'price_value', 'bedrooms', 'bathrooms', 'size_interior',
    'property_type', 'ownership_type', 'stories', 'parking', 'address', 'city_province', 'postal_code',
    'province', 'latitude', 'longitude', 'listed_date', 'days_on_realtor', 'photo_url', 'listing_url',
    'brokerages', 'listing_agents', 'agent_phones', 'searched', 'first_seen'];
const NUMBER_COLUMNS = new Set(['price_value', 'latitude', 'longitude', 'days_on_realtor']);

function parseArgs(argv) {
    const o = { cities: [], out: 'listings.csv', limit: 100, maxCharge: 2, append: false };
    const num = (flag, v) => {
        const n = Number(v);
        if (v === undefined || !Number.isFinite(n) || n < 0) throw new Error(`${flag} needs a number`);
        return n;
    };
    const val = (flag, v) => {
        if (v === undefined || v.startsWith('--')) throw new Error(`${flag} needs a value`);
        return v;
    };
    for (let i = 0; i < argv.length; i++) {
        const a = argv[i];
        if (a === '-h' || a === '--help') o.help = true;
        else if (a === '--file') o.cities.push(...readList(val(a, argv[++i])));
        else if (a === '--type') o.type = val(a, argv[++i]);
        else if (a === '--property') o.property = val(a, argv[++i]);
        else if (a === '--limit') o.limit = num(a, argv[++i]);
        else if (a === '--any-city') o.anyCity = true;
        else if (a === '--out') o.out = val(a, argv[++i]);
        else if (a === '--append') o.append = true;
        else if (a === '--new-out') o.newOut = val(a, argv[++i]);
        else if (a === '--no-agents') o.noAgents = true;
        else if (a === '--max-charge') o.maxCharge = num(a, argv[++i]);
        else if (a === '--dataset') o.dataset = val(a, argv[++i]);
        else if (a === '--from') o.from = val(a, argv[++i]);
        else if (a.startsWith('-')) throw new Error(`Unknown option ${a} (see --help)`);
        else o.cities.push(a);
    }
    if (o.help) return o;
    if (o.type && !['sale', 'rent'].includes(o.type)) throw new Error('--type is sale or rent');
    if (o.property && !['residential', 'commercial'].includes(o.property)) {
        throw new Error('--property is residential or commercial');
    }
    if (!Number.isInteger(o.limit) || o.limit < 1) throw new Error('--limit needs a whole number of 1 or more');
    if (o.limit > MAX_PER_CITY) {
        console.error(`--limit ${o.limit} -> ${MAX_PER_CITY}: one Realtor.ca search area returns at most ${MAX_PER_CITY} listings.`);
        o.limit = MAX_PER_CITY;
    }
    for (const f of [o.out, o.newOut].filter(Boolean)) {
        if (!/\.(csv|xlsx|jsonl)$/i.test(f)) throw new Error(`${f}: output files must end in .csv, .xlsx or .jsonl`);
    }
    if (o.append && /\.xlsx$/i.test(o.out)) throw new Error('--append works with .csv or .jsonl (Excel opens the CSV)');
    if (o.newOut && !o.append) throw new Error('--new-out needs --append (new = not in --out yet)');
    const sources = [o.cities.length > 0, Boolean(o.dataset), Boolean(o.from)].filter(Boolean).length;
    if (sources === 0) throw new Error('Give one or more cities, --file cities.txt, --dataset <id> or --from <file> (see --help)');
    if (sources > 1) throw new Error('Use one of: cities/--file, --dataset, --from');
    if (o.cities.length && !o.anyCity) {
        const known = new Set(CITIES.map((c) => c.toLowerCase()));
        const unknown = o.cities.filter((c) => !known.has(c.split(',')[0].trim().toLowerCase()));
        if (unknown.length) {
            throw new Error(`Not a supported city: ${unknown.join('; ')}\nSupported: ${CITIES.join(', ')} `
                + '(no postal codes, provinces or neighbourhoods; write Montreal without the accent).');
        }
    }
    return o;
}

function readList(path) {
    return readFileSync(path, 'utf8').split('\n').map((s) => s.trim()).filter((s) => s && !s.startsWith('#'));
}

const today = () => new Date().toISOString().slice(0, 10);
const join = (list) => {
    const parts = [...new Set((list ?? []).filter((v) => v != null && v !== ''))];
    return parts.length ? parts.join(' | ') : null;
};

// One Actor row (or one row this script wrote earlier) -> one flat output row.
function toRow(r, { noAgents }) {
    const row = 'mls_number' in r
        ? Object.fromEntries(COLUMNS.map((c) => [c, r[c] ?? null]))
        : {
            mls_number: r.mlsNumber ?? null,
            price: r.price ?? null,
            price_value: r.priceValue ?? null,
            bedrooms: r.bedrooms ?? null,
            bathrooms: r.bathrooms ?? null,
            size_interior: r.sizeInterior ?? null,
            property_type: r.propertyType ?? null,
            ownership_type: r.ownershipType ?? null,
            stories: r.stories ?? null,
            parking: r.parking ?? null,
            address: r.address ?? null,
            city_province: r.cityProvince ?? null,
            postal_code: r.postalCode ?? null,
            province: r.province ?? null,
            latitude: r.latitude ?? null,
            longitude: r.longitude ?? null,
            listed_date: listedDate(r.listedDateUTC),
            days_on_realtor: daysSince(listedDate(r.listedDateUTC)),
            photo_url: r.photoUrl ?? null,
            listing_url: r.listingUrl ?? null,
            brokerages: join((r.agents ?? []).map((a) => a.brokerage)),
            listing_agents: join((r.agents ?? []).map((a) => a.name)),
            agent_phones: join((r.agents ?? []).map((a) => a.phone)),
            searched: r.location ?? null,
            first_seen: today(),
        };
    if (noAgents) { row.listing_agents = null; row.agent_phones = null; }
    if (!row.first_seen) row.first_seen = today();
    return row;
}

// Realtor.ca sends the listing date as .NET ticks (100 ns since 0001-01-01 UTC) or as an ISO
// string. Either way the export gets an ISO date-time in UTC.
function listedDate(v) {
    if (v == null || v === '') return null;
    if (/^\d{15,}$/.test(String(v))) {
        const ms = Number(BigInt(String(v)) / 10000n - 62135596800000n);
        return new Date(ms).toISOString().replace(/\.\d{3}Z$/, 'Z');
    }
    const d = new Date(v);
    return Number.isNaN(d.getTime()) ? String(v) : d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

// Whole days between listed_date and now. Computed here because the Actor's daysOnRealtor holds
// the number from Realtor.ca's "8 min ago" / "3 hours ago" label for listings under a day old.
function daysSince(iso) {
    if (!iso) return null;
    const ms = Date.now() - Date.parse(iso);
    return Number.isFinite(ms) ? Math.max(0, Math.floor(ms / 86400000)) : null;
}

const isListing = (r) => 'mls_number' in r || 'mlsNumber' in r;

// CSV cell. Text that starts with = + - @, a tab or a carriage return gets a leading ' so Excel and
// Google Sheets show it as text instead of running it as a formula.
function csvCell(v) {
    if (v == null) return '';
    let s = String(v);
    if (typeof v === 'string' && /^[=+\-@\t\r]/.test(s)) s = `'${s}`;
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

// Minimal RFC 4180 reader, enough to find the mls_number column of a CSV this script wrote.
function parseCsv(text) {
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (q) {
            if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; }
            else if (ch === '"') q = false;
            else cell += ch;
        } else if (ch === '"') q = true;
        else if (ch === ',') { row.push(cell); cell = ''; }
        else if (ch === '\n' || ch === '\r') {
            if (ch === '\r' && text[i + 1] === '\n') i++;
            row.push(cell); rows.push(row); row = []; cell = '';
        } else cell += ch;
    }
    if (cell || row.length) { row.push(cell); rows.push(row); }
    return rows;
}

function existingIds(path) {
    if (!existsSync(path) || statSync(path).size === 0) return new Set();
    const text = readFileSync(path, 'utf8').replace(/^﻿/, '');
    if (/\.jsonl$/i.test(path)) {
        return new Set(text.split('\n').filter(Boolean).map((l) => JSON.parse(l).mls_number).filter(Boolean));
    }
    const [header, ...rows] = parseCsv(text);
    if (header.join(',') !== COLUMNS.join(',')) {
        throw new Error(`${path} has different columns than this script writes; use a new --out file`);
    }
    const idx = COLUMNS.indexOf('mls_number');
    return new Set(rows.map((r) => r[idx]).filter(Boolean));
}

async function writeXlsx(path, rows) {
    // Loaded only for .xlsx so CSV/JSONL users don't need it.
    const { default: writeExcelFile } = await import('write-excel-file/node').catch(() => {
        throw new Error('Writing .xlsx needs write-excel-file: run npm install');
    });
    // Strip characters that are not allowed in XLSX (XML) cells.
    const clean = (s) => (s == null ? null : String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ''));
    const cell = (col, v) => {
        if (v == null || v === '') return null;
        if (NUMBER_COLUMNS.has(col)) return { value: Number(v), type: Number };
        if (col === 'listed_date') return { value: new Date(v), type: Date, format: 'yyyy-mm-dd hh:mm' };
        if (col === 'first_seen') return { value: new Date(`${v}T00:00:00Z`), type: Date, format: 'yyyy-mm-dd' };
        return { value: clean(v), type: String };
    };
    const data = [
        COLUMNS.map((value) => ({ value, fontWeight: 'bold' })),
        ...rows.map((r) => COLUMNS.map((c) => cell(c, r[c]))),
    ];
    const widths = { address: 30, city_province: 40, listing_url: 50, photo_url: 50, brokerages: 40,
        listing_agents: 30, agent_phones: 26, listed_date: 17 };
    const columns = COLUMNS.map((c) => ({ width: widths[c] ?? 14 }));
    await writeExcelFile(data, { columns, stickyRowsCount: 1, sheet: 'Listings' }).toFile(path);
}

async function writeRows(path, rows, append) {
    if (/\.xlsx$/i.test(path)) return writeXlsx(path, rows);
    if (/\.jsonl$/i.test(path)) {
        const body = rows.map((r) => JSON.stringify(r)).join('\n') + (rows.length ? '\n' : '');
        if (append) appendFileSync(path, body);
        else writeFileSync(path, body);
        return;
    }
    const lines = rows.map((r) => COLUMNS.map((c) => csvCell(r[c])).join(','));
    const fresh = !append || !existsSync(path) || statSync(path).size === 0;
    // UTF-8 with a BOM, so Excel shows accents (Montréal, Québec) correctly when it opens the file.
    if (fresh) writeFileSync(path, `﻿${[COLUMNS.join(','), ...lines].join('\r\n')}\r\n`);
    else if (lines.length) appendFileSync(path, `${lines.join('\r\n')}\r\n`);
}

function readFromFile(path) {
    const text = readFileSync(path, 'utf8').replace(/^﻿/, '').trim();
    if (!text) return [];
    const data = text.startsWith('[') ? JSON.parse(text) : text.split('\n').filter(Boolean).map((l) => JSON.parse(l));
    return Array.isArray(data) ? data : [data];
}

async function main() {
    const o = parseArgs(process.argv.slice(2));
    if (o.help) { console.log(HELP); return; }

    let items;
    let failedRun = false;
    if (o.from) {
        items = readFromFile(o.from);
    } else {
        if (!process.env.APIFY_TOKEN) throw new Error('Set APIFY_TOKEN (see .env.example), or use --from <file>');
        const { ApifyClient } = await import('apify-client'); // loaded here so --help and --from work before npm install
        const client = new ApifyClient({ token: process.env.APIFY_TOKEN });
        let datasetId = o.dataset;
        if (!datasetId) {
            const input = {
                locations: o.cities,
                limitPerLocation: o.limit,
                transactionType: o.type ?? 'sale',
                propertyType: o.property ?? 'residential',
            };
            console.error(`Running ${ACTOR}: ${o.cities.length} city(ies), up to ${o.limit} ${input.propertyType} `
                + `listings for ${input.transactionType} each, cost cap $${o.maxCharge} ...`);
            // maxTotalChargeUsd caps what this run can cost you (pay-per-event Actors only).
            const run = await client.actor(ACTOR).call(input, { maxTotalChargeUsd: o.maxCharge });
            console.error(`Run ${run.id} finished: ${run.status}${run.statusMessage ? ` (${run.statusMessage})` : ''}`);
            failedRun = run.status !== 'SUCCEEDED';
            datasetId = run.defaultDatasetId;
        }
        items = [];
        for await (const it of client.dataset(datasetId).listItems()) items.push(it);
    }
    if (failedRun && items.length === 0) {
        console.error(`No listings, so ${o.out} was left as it was.`);
        process.exit(1);
    }

    let rows = items.filter(isListing).map((r) => toRow(r, o));
    let skipped = 0;
    if (o.append) {
        const seen = existingIds(o.out);
        const all = rows;
        rows = all.filter((r) => !r.mls_number || !seen.has(String(r.mls_number)));
        skipped = all.length - rows.length;
        // The Actor returns the newest listings first. If a city's run came back full and none of
        // its listings were in the file yet, it may not reach back to the previous run: some
        // listings in between were never seen.
        if (seen.size && !o.from && !o.dataset) {
            for (const city of new Set(all.map((r) => r.searched))) {
                const got = all.filter((r) => r.searched === city);
                if (got.length >= o.limit && got.every((r) => !seen.has(String(r.mls_number)))) {
                    console.error(`${city}: all ${got.length} listings are new, so this run may not reach back to the `
                        + 'last one. Run more often or raise --limit to close the gap.');
                }
            }
        }
    }
    await writeRows(o.out, rows, o.append);
    if (o.newOut) await writeRows(o.newOut, rows, false);

    if (o.append) {
        console.error(`${rows.length} new listings appended to ${o.out}${skipped ? ` (${skipped} already there)` : ''}`
            + `${o.newOut ? `; the new ones are also in ${o.newOut}` : ''}`);
    } else {
        console.error(`Wrote ${rows.length} listings to ${o.out}`);
    }
}

main().catch((err) => {
    console.error(err.message);
    process.exit(1);
});
