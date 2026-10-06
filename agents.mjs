#!/usr/bin/env node
// Export the Realtor.ca agent directory to CSV, Excel (.xlsx) or JSONL.
//
// Calls the yugenox/realtor-ca-agent-scraper Actor on Apify for one or more Canadian cities,
// neighbourhoods or provinces, then writes one row per agent: name, title, brokerage, office
// address, phone (when published), profile link and photo. Realtor.ca doesn't publish agent
// emails, so there are none.
//
// Run with --help for every option. Requires Node 20.6+; APIFY_TOKEN in the environment (or in
// .env, loaded with `node --env-file=.env`) for new runs and --dataset.

import { appendFileSync, existsSync, readFileSync, statSync, writeFileSync } from 'node:fs';

const ACTOR = 'yugenox/realtor-ca-agent-scraper';

// Realtor.ca's agent search never resolves a postal code (full or first 3 characters), so the
// script stops before starting a run for one.
const POSTAL_RE = /^[A-Za-z]\d[A-Za-z][ -]?(\d[A-Za-z]\d)?$/;

const HELP = `Export the Realtor.ca agent directory to CSV, Excel (.xlsx) or JSONL.

Usage:
  node --env-file=.env agents.mjs "Toronto, ON" ["North York" ...] [options]
  node --env-file=.env agents.mjs --file locations.txt --limit 200 --out agents.xlsx
  node --env-file=.env agents.mjs --dataset <datasetId> --out agents.xlsx
  node agents.mjs --from rows.json --out agents.xlsx

What to search
  <location>...         A city ("Toronto, ON", "Mississauga"), a district or neighbourhood
                        ("North York", "Yorkville") or a province ("Ontario", "ON").
                        Spelling must match Realtor.ca. Postal codes don't work.
                        Provinces work but are slower than naming cities.
  --file <path>         One location per line. Blank lines and lines starting with # are skipped.
  --limit <n>           Agents per location (default 50).

Output
  --out <path>          .csv (default agents.csv), .xlsx or .jsonl
  --append              Add only agents whose individual_id is not in --out yet (.csv or .jsonl).
  --max-charge <usd>    Hard cost cap for the run, enforced by Apify (default 2).

Re-export without a new run
  --dataset <id>        Export the dataset of a run you already made (Apify Console >
                        Storage > Datasets). No new run starts.
  --from <file>         Convert a JSON or JSONL file of the Actor's rows (Console export)
                        or of this script's rows. Needs no token.

  -h, --help            Show this help.

Price: $0.001 per agent plus a start fee per run (Apify Store, checked 2026-10-06; the Store page
shows the current price). Not affiliated with CREA or REALTOR.ca.`;

// Column order for CSV, XLSX and JSONL.
const COLUMNS = ['individual_id', 'name', 'position', 'brokerage', 'office_address', 'phone',
    'profile_url', 'photo_url', 'searched', 'first_seen'];

function parseArgs(argv) {
    const o = { locations: [], out: 'agents.csv', limit: 50, maxCharge: 2, append: false };
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
        else if (a === '--file') o.locations.push(...readList(val(a, argv[++i])));
        else if (a === '--limit') o.limit = num(a, argv[++i]);
        else if (a === '--out') o.out = val(a, argv[++i]);
        else if (a === '--append') o.append = true;
        else if (a === '--max-charge') o.maxCharge = num(a, argv[++i]);
        else if (a === '--dataset') o.dataset = val(a, argv[++i]);
        else if (a === '--from') o.from = val(a, argv[++i]);
        else if (a.startsWith('-')) throw new Error(`Unknown option ${a} (see --help)`);
        else o.locations.push(a);
    }
    if (o.help) return o;
    if (!Number.isInteger(o.limit) || o.limit < 1) throw new Error('--limit needs a whole number of 1 or more');
    if (!/\.(csv|xlsx|jsonl)$/i.test(o.out)) throw new Error('--out must end in .csv, .xlsx or .jsonl');
    if (o.append && /\.xlsx$/i.test(o.out)) throw new Error('--append works with .csv or .jsonl (Excel opens the CSV)');
    const sources = [o.locations.length > 0, Boolean(o.dataset), Boolean(o.from)].filter(Boolean).length;
    if (sources === 0) throw new Error('Give one or more locations, --file locations.txt, --dataset <id> or --from <file> (see --help)');
    if (sources > 1) throw new Error('Use one of: locations/--file, --dataset, --from');
    const postal = o.locations.filter((l) => POSTAL_RE.test(l.split(',')[0].trim()));
    if (postal.length) {
        throw new Error(`Realtor.ca's agent search doesn't take postal codes: ${postal.join('; ')}. `
            + 'Use the city, district or neighbourhood instead.');
    }
    return o;
}

function readList(path) {
    return readFileSync(path, 'utf8').split('\n').map((s) => s.trim()).filter((s) => s && !s.startsWith('#'));
}

const today = () => new Date().toISOString().slice(0, 10);

// One Actor row (or one row this script wrote earlier) -> one flat output row.
function toRow(r) {
    const row = 'individual_id' in r
        ? Object.fromEntries(COLUMNS.map((c) => [c, r[c] ?? null]))
        : {
            individual_id: r.individualId != null ? String(r.individualId) : null,
            name: r.name ?? null,
            position: r.position ?? null,
            brokerage: r.brokerage ?? null,
            office_address: r.officeAddress ?? null,
            phone: r.phone ?? null,
            profile_url: r.profileUrl ?? null,
            photo_url: r.photoUrl ?? null,
            searched: r.location ?? null,
            first_seen: today(),
        };
    if (!row.first_seen) row.first_seen = today();
    return row;
}

const isAgent = (r) => 'individual_id' in r || 'individualId' in r;

// CSV cell. Text that starts with = + - @, a tab or a carriage return gets a leading ' so Excel and
// Google Sheets show it as text instead of running it as a formula.
function csvCell(v) {
    if (v == null) return '';
    let s = String(v);
    if (typeof v === 'string' && /^[=+\-@\t\r]/.test(s)) s = `'${s}`;
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

// Minimal RFC 4180 reader, enough to find the individual_id column of a CSV this script wrote.
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
        return new Set(text.split('\n').filter(Boolean).map((l) => JSON.parse(l).individual_id).filter(Boolean));
    }
    const [header, ...rows] = parseCsv(text);
    if (header.join(',') !== COLUMNS.join(',')) {
        throw new Error(`${path} has different columns than this script writes; use a new --out file`);
    }
    const idx = COLUMNS.indexOf('individual_id');
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
        if (col === 'first_seen') return { value: new Date(`${v}T00:00:00Z`), type: Date, format: 'yyyy-mm-dd' };
        return { value: clean(v), type: String };
    };
    const data = [
        COLUMNS.map((value) => ({ value, fontWeight: 'bold' })),
        ...rows.map((r) => COLUMNS.map((c) => cell(c, r[c]))),
    ];
    const widths = { name: 26, position: 22, brokerage: 40, office_address: 50, phone: 14, profile_url: 44,
        photo_url: 44, searched: 16 };
    const columns = COLUMNS.map((c) => ({ width: widths[c] ?? 12 }));
    await writeExcelFile(data, { columns, stickyRowsCount: 1, sheet: 'Agents' }).toFile(path);
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
    // UTF-8 with a BOM, so Excel shows accents correctly when it opens the file.
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
            const input = { locations: o.locations, limitPerLocation: o.limit };
            console.error(`Running ${ACTOR}: ${o.locations.length} location(s), up to ${o.limit} agents each, `
                + `cost cap $${o.maxCharge} ...`);
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
        console.error(`No agents, so ${o.out} was left as it was.`);
        process.exit(1);
    }

    let rows = items.filter(isAgent).map(toRow);
    let skipped = 0;
    if (o.append) {
        const seen = existingIds(o.out);
        const before = rows.length;
        rows = rows.filter((r) => !r.individual_id || !seen.has(String(r.individual_id)));
        skipped = before - rows.length;
    }
    await writeRows(o.out, rows, o.append);

    const verb = o.append ? 'Appended' : 'Wrote';
    console.error(`${verb} ${rows.length} agents to ${o.out}${skipped ? ` (${skipped} already there)` : ''}`);
}

main().catch((err) => {
    console.error(err.message);
    process.exit(1);
});
