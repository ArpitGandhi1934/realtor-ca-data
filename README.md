# Realtor.ca Listings to CSV and Excel (Python + Node.js)

[![Realtor.ca Property Scraper on Apify][badge-listings]][listings]

**Export Realtor.ca listings** to **Excel (.xlsx), CSV or JSONL**: MLS® number, price, beds, baths, size, address, postal code, coordinates, listing date, main photo, listing link and the listing brokerage. Listings can be for sale or for rent, residential or commercial, in 18 Canadian cities.

Realtor.ca has no public API for this data. Instead of running your own Realtor.ca scraper in Python, the scripts here (Python and Node.js) call a hosted Actor on Apify, the [Realtor.ca Property Scraper][listings]. You need an Apify API token. You don't need a Realtor.ca account, an MLS® data feed or CREA membership, and there is no scraper code for you to maintain.

> Not affiliated with or endorsed by The Canadian Real Estate Association (CREA) or REALTOR.ca. Public listing pages only.

**Coming soon:** examples for the Realtor.ca agent directory. For now this repo covers listings.

## What you get

Real output from one run on 2026-10-06: the 20 newest residential listings for sale in the Toronto search area, from `node listings.mjs "Toronto, ON" --limit 20`. In [`samples/`](samples) the listing agents are pseudonymized and their phone numbers removed ([how](samples/redact.jq)); everything else is as published.

| mls_number | price | bedrooms | bathrooms | size_interior | city_province | brokerages |
|---|---:|---|---|---|---|---|
| C13865220 | $980,000 | 3 | 2 | 65.0316 m2 | Toronto (Oakwood Village), Ontario M6E1H2 | BOSLEY REAL ESTATE LTD. |
| C13865188 | $16,880,000 | 5 + 1 | 9 | 464.5113 m2 | Toronto (Forest Hill South), Ontario M5P2N4 | SOTHEBY'S INTERNATIONAL REALTY CANADA |
| C13865156 | $449,900 | 1 | 1 | 46.4511 m2 | Toronto (Waterfront Communities), Ontario M5A0S6 | RE/MAX EXCEL TITAN |
| C13865170 | $849,000 | 3 + 2 | 4 | 102.1925 m2 | Toronto (Oakwood Village), Ontario M6E2E6 | KELLER WILLIAMS PORTFOLIO REALTY |
| C13864980 | $525,000 | 1 | 1 | 46.4511 m2 | Toronto (Waterfront Communities), Ontario M5V0N4 | ROYAL LEPAGE VISION REALTY |

The same 20 rows as [Excel](samples/toronto-listings.xlsx), [CSV](samples/toronto-listings.csv) and the Actor's [raw JSON rows](samples/toronto-listings.json). They were listed between 22:25 UTC on October 5 and 04:24 UTC on October 6: the Actor returns the newest listings first.

### Columns (`listings.mjs` / `listings.py`)

| Column | Meaning |
|---|---|
| `mls_number` | MLS® number. Stable, so it's the key `--append` uses |
| `price`, `price_value` | Price as shown (`$980,000`) and as a number |
| `bedrooms`, `bathrooms` | As published, for example `5 + 1` |
| `size_interior` | Interior size as published (often in m2); empty when the listing has none |
| `property_type`, `ownership_type` | For example `Single Family` and `Freehold` or `Condominium/Strata` |
| `stories`, `parking` | Storeys and parking spaces, when published |
| `address`, `city_province`, `postal_code`, `province` | Street address; city and neighbourhood as Realtor.ca shows them |
| `latitude`, `longitude` | Coordinates, ready for a map or a radius query |
| `listed_date` | When the listing went up on Realtor.ca (UTC; a date-time cell in the .xlsx) |
| `days_on_realtor` | Whole days from `listed_date` to the moment you exported |
| `photo_url` | The main listing photo (one photo per listing) |
| `listing_url` | The listing's page on Realtor.ca |
| `brokerages` | The listing brokerage(s) |
| `listing_agents`, `agent_phones` | The listing agent(s) and their phone numbers, as published on the listing. `--no-agents` leaves them out |
| `searched` | The city you asked for |
| `first_seen` | The date (UTC) this script first wrote the row. With `--append` it marks when a listing was new |

There's no email column: Realtor.ca doesn't publish agent email addresses.

## Quick start

1. Create a free Apify account and copy your API token ([sign up][signup], then Console > Settings > API & Integrations).
2. `cp .env.example .env` and paste the token after `APIFY_TOKEN=`.

**Python** (3.10 or newer):

```bash
pip install -r requirements.txt
python listings.py "Toronto, ON" --limit 200 --out toronto.xlsx
python listings.py "Vancouver, BC" --type rent --out vancouver-rentals.csv
```

**Node.js** (20.6 or newer):

```bash
npm install
node --env-file=.env listings.mjs "Calgary, AB" "Edmonton, AB" --limit 300 --out alberta.xlsx
node --env-file=.env listings.mjs "Ottawa, ON" --property commercial --out ottawa-commercial.csv
node --env-file=.env listings.mjs --file cities.example.txt --out gta-ottawa.csv
```

Both languages take the same options and write byte-identical CSV and JSONL (`--help` prints them). The CSV is UTF-8 with a BOM so Excel shows accents correctly, and text that starts with `=`, `+`, `-` or `@` gets a leading `'` so a spreadsheet never runs it as a formula.

### Options

| Option | What it does | Default |
|---|---|---|
| cities | `"Toronto, ON"` or `Toronto`. One or more of: Toronto, Mississauga, Brampton, Vaughan, Markham, Oakville, Hamilton, Ottawa, London, Vancouver, Surrey, Burnaby, Victoria, Calgary, Edmonton, Montreal, Winnipeg, Halifax | |
| `--file cities.txt` | One city per line ([example](cities.example.txt)) | |
| `--type sale\|rent` | For sale or for rent | `sale` |
| `--property residential\|commercial` | Residential or commercial | `residential` |
| `--limit N` | Listings per city, newest first. 500 at most | `100` |
| `--out FILE` | `.xlsx`, `.csv` or `.jsonl` | `listings.csv` |
| `--append` | Add only listings whose `mls_number` isn't in `--out` yet (`.csv` or `.jsonl`) | |
| `--new-out FILE` | With `--append`: also write just this run's new listings to their own file | |
| `--no-agents` | Leave the listing agents' names and phones out; the brokerage stays | |
| `--max-charge USD` | Hard cost cap for the run, enforced by Apify. Rows past the cap aren't collected | `2` |
| `--dataset ID` | Export a run you already made in Apify Console. No new run starts | |
| `--from FILE` | Convert a JSON or JSONL file of the Actor's rows (a Console export) or of this script's rows. Needs no token | |
| `--any-city` | Skip the script's city check, in case the Actor adds cities | |

The Actor filters by city, sale or rent, and residential or commercial. Price, bedroom and size aren't Actor inputs, so filter those in the file; every listing a run returns is billed.

## New listings since your last run

The Actor has no "new since" switch, but it returns the newest listings first and every listing has an MLS® number. So run the export on a schedule, append to one file, and only listings the file hasn't seen are added:

```bash
# crontab -e   (every morning at 7)
0 7 * * * cd /path/to/realtor-ca-data && node --env-file=.env listings.mjs "Toronto, ON" --limit 200 --append --out toronto.csv --new-out toronto-new-today.csv
```

- `toronto.csv` grows by the listings that weren't in it yet; `first_seen` says which day each one appeared.
- `toronto-new-today.csv` holds only today's new ones, ready to send to a sheet, a CRM or an email.
- If every listing a run returns is new, the script warns that the run may not reach back to the previous one. Raise `--limit` (up to 500) or run more often.
- You pay for every listing a run returns, including the ones already in your file. Pick a `--limit` a little above how many listings your area gets between runs: in the Toronto sample, the 20 newest residential listings for sale covered about 6 hours of a Monday evening.

Prefer to run it on Apify? Save the input as a task in Apify Console, attach a daily Schedule, and append each day's result locally:

```bash
curl -s "https://api.apify.com/v2/actor-tasks/<TASK_ID>/runs/last/dataset/items?status=SUCCEEDED&clean=true" \
  -H "Authorization: Bearer $APIFY_TOKEN" -o today.json
node listings.mjs --from today.json --append --out toronto.csv --new-out toronto-new-today.csv
```

## Use cases

- **Canadian real estate market data:** asking prices, sizes and new supply per city, grouped by neighbourhood or postal code in your spreadsheet, refreshed on a schedule.
- **Businesses that serve listings** (photographers, stagers, mortgage brokers): each morning's new listings in your area, with the listing brokerage.

## Price

Pay per listing, billed to your own Apify account. From the Actor's pricing in the Apify API on 2026-10-06 (the Store page always shows the current price):

| | Realtor.ca Property Scraper |
|---|---:|
| Per listing | $0.001 ($1.00 per 1,000) |
| Start fee per run, until 2026-10-11 (19:57 UTC) | $0.00005 |
| Start fee per run from 2026-10-11 (19:57 UTC) | $0.002 |

The per-listing price is the same on every Apify plan and doesn't change on October 11; only the start fee does. The start fee is charged once per run per GB of memory, and runs use 1 GB by default. The Actor's compute isn't billed on top.

| Example | Rows | Until 2026-10-11 | From 2026-10-11 |
|---|---:|---:|---:|
| One city, `--limit 100` (the default) | 100 listings | $0.10 | $0.10 |
| One city, `--limit 500` (the maximum) | 500 listings | $0.50 | $0.50 |
| All 18 cities, `--limit 500 --max-charge 10` | up to 9,000 listings | up to $9.00 | up to $9.00 |
| Daily new-listings check, one city, `--limit 200`, 30 days | 6,000 listings | $6.00 | $6.06 |

Amounts are rounded to the cent. Apify's Free plan includes monthly usage credit you can spend on this Actor. `--max-charge` (default $2) is a hard cap Apify enforces for each run: rows past the cap aren't collected, and the scripts warn before starting a run that could go past it. Raise it for big runs, as in the 18-city row.

## Limits

- **18 cities:** Toronto, Mississauga, Brampton, Vaughan, Markham, Oakville, Hamilton, Ottawa, London, Vancouver, Surrey, Burnaby, Victoria, Calgary, Edmonton, Montreal, Winnipeg and Halifax. No postal codes, provinces or neighbourhoods. Write Montreal without the accent.
- **Each city is a fixed map area around its centre,** about 24 km north to south. It can miss the outer edges of a large city and take in part of a neighbour, so check `city_province` when you need exact city limits.
- **Up to 500 listings per city per run, newest first.** The Toronto area had 8,114 active residential listings for sale during the sample run (the Actor's log), so one run gets the newest 500 of them. The "new listings" recipe above is how you keep a complete feed over time.
- **Active listings only.** No sold prices and no price history.
- One photo per listing: the main photo URL.
- The sample run above used listings for sale. For rentals or commercial listings, start with `--limit 5` and check that the columns you need are filled.

## FAQ

**Does Realtor.ca have an API?**
Not a public one for listing data. CREA's REALTOR.ca DDF® "is a managed service for REALTORS® and broker owners who are looking to share listings with multiple third-party websites" ([CREA][ddf], read 2026-10-06). If you are a REALTOR® member or build for one, DDF® is the licensed route. Otherwise, the Actor here reads what Realtor.ca shows any visitor and returns it as data you can call from Python, Node.js, `curl` or any HTTP client.

**Is scraping Realtor.ca legal?**
Here are the facts; this is not legal advice. The Actor reads only the listing data Realtor.ca shows any visitor without an account, and never signs in. The rows include the listing agents' names and phone numbers, as brokerages publish them on Realtor.ca. That is business contact information, and still personal information under PIPEDA (Canada), GDPR and similar laws. You're responsible for having a lawful reason to collect and keep it, for following REALTOR.ca's terms of use, and for getting consent before you send commercial electronic messages (CASL). For a commercial product, ask a lawyer. More: [Is web scraping legal?][legal]

**Do I need an MLS® feed, CREA membership or a brokerage login?**
No. The Actor doesn't use any of them.

**Can I get agent email addresses?**
No. Realtor.ca doesn't publish them; an agent's page offers a contact form instead.

**Can I get sold prices or price history?**
No. The Actor returns active listings only. To track asking-price changes, save each day's export to its own file and compare `price_value` by `mls_number`.

**Can I filter by price or bedrooms?**
Not in the Actor. Its inputs are the city, sale or rent, residential or commercial, and the number of listings. Filter the CSV or the Excel sheet.

**Why was my city rejected?**
The Actor covers the 18 cities above. Postal codes, provinces and neighbourhoods aren't inputs, and "Montréal" has to be written "Montreal".

**How fresh is the data?**
Read live from Realtor.ca when the run starts. Schedule the export to keep a file current.

**Can I call it from Claude, ChatGPT or Cursor?**
Yes, through Apify's hosted MCP server. Add `https://mcp.apify.com/?tools=fetch-actor-details,yugenox/realtor-ca-property-scraper` to your assistant as a custom connector and sign in with your Apify account; runs bill to that account. In Claude Code: `claude mcp add --transport http yugenox-realtor-ca-listings "https://mcp.apify.com/?tools=fetch-actor-details,yugenox/realtor-ca-property-scraper"`. Steps for each client: [use it in ChatGPT, Claude and Cursor][ai].

## Privacy

The listing agents' names and phone numbers are business contact details, published on Realtor.ca so clients can reach the agent, and they're personal information all the same. Keep only what you need, don't resell agent lists, and don't send commercial email or texts without consent (CASL). `--no-agents` drops the listing agents' names and phones from an export. This repo commits no real agent details: the listing sample is pseudonymized.

## More

- Guides and other public-data examples: [yugenox-data.vercel.app](https://yugenox-data.vercel.app)
- The Actor, with its input form, output table and API page: [Realtor.ca Property Scraper][listings] ([API][listings-api]). n8n and Make can run it with their Apify node.
- More Canadian data (Kijiji, grocery prices, store stock, jobs): [yugenox on Apify][store]
- Use it from AI assistants: [setup guide][ai]

## Samples, testing and legal

- The listing sample comes from one run on 2026-10-06 (Property Scraper build 0.1.11, 20 listings) through `listings.mjs`. The raw rows went through [`samples/redact.jq`](samples/redact.jq), and the .csv and .xlsx were written from them with `--from`. Input keys were checked against the Actor's live input schema the same day.
- MLS®, Multiple Listing Service® and the associated logos are trademarks owned by The Canadian Real Estate Association (CREA). REALTOR®, REALTORS® and the REALTOR® logo are trademarks controlled by CREA. This repo and the Actor are not affiliated with or endorsed by CREA or REALTOR.ca.
- The MIT licence covers the code. The listing data in `samples/` is a short excerpt of public Realtor.ca listings, included to show the output format; it isn't ours to license.
- Made by Yugenox Corporation.

<!-- All apify.com links for this README live below. When the Apify affiliate id exists, append ?fpr=<id> to these URLs only. -->
[listings]: https://apify.com/yugenox/realtor-ca-property-scraper
[listings-api]: https://apify.com/yugenox/realtor-ca-property-scraper/api
[badge-listings]: https://apify.com/actor-badge?actor=yugenox/realtor-ca-property-scraper
[store]: https://apify.com/yugenox
[signup]: https://console.apify.com/sign-up
[legal]: https://blog.apify.com/is-web-scraping-legal/
[ddf]: https://www.crea.ca/technology/realtor-ca-for-realtors/realtor-ca-tools/realtor-ca-ddf/
[ai]: https://yugenox-data.vercel.app/ai/use-in-chatgpt-claude-cursor
