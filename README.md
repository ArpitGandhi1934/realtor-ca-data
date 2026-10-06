# Realtor.ca Listings and Agents to CSV and Excel (Python + Node.js)

[![Realtor.ca Property Scraper on Apify][badge-listings]][listings] [![Realtor.ca Agent Scraper on Apify][badge-agents]][agents]

**Export Realtor.ca listings** to **Excel (.xlsx), CSV or JSONL**: MLS® number, price, beds, baths, size, address, postal code, coordinates, listing date, main photo, listing link and the listing brokerage. Listings can be for sale or for rent, residential or commercial, in 18 Canadian cities. A second script exports the **Realtor.ca agent directory** by city, neighbourhood or province: name, title, brokerage, office address and phone.

Realtor.ca has no public API for this data. Instead of running your own Realtor.ca scraper in Python, the scripts here (Python and Node.js) call two hosted Actors on Apify, the [Realtor.ca Property Scraper][listings] and the [Realtor.ca Agent Scraper][agents]. You need an Apify API token. You don't need a Realtor.ca account, an MLS® data feed or CREA membership, and there is no scraper code for you to maintain.

> Not affiliated with or endorsed by The Canadian Real Estate Association (CREA) or REALTOR.ca. Public listing and directory pages only.

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

### Listing columns (`listings.mjs` / `listings.py`)

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

### Agent columns (`agents.mjs` / `agents.py`)

| Column | Meaning |
|---|---|
| `individual_id` | Realtor.ca's agent ID. Stable, so it's the key `--append` uses |
| `name` | As CREA stores it, often in capitals |
| `position` | For example Sales Representative, Broker, Broker of Record |
| `brokerage`, `office_address` | The brokerage and its office address (empty when the agent hides the address on Realtor.ca too) |
| `phone` | The agent's phone, when published |
| `profile_url`, `photo_url` | Realtor.ca profile page and profile photo |
| `searched`, `first_seen` | The location you asked for, and the date (UTC) this script first wrote the row |

There's no email column: Realtor.ca doesn't publish agent email addresses. The agent sample in [`samples/agents-synthetic.csv`](samples/agents-synthetic.csv) is **made up** (fictional people, brokerages and 555-01xx numbers) to show the format; this repo contains no real agent rows.

## Quick start

1. Create a free Apify account and copy your API token ([sign up][signup], then Console > Settings > API & Integrations).
2. `cp .env.example .env` and paste the token after `APIFY_TOKEN=`.

**Python** (3.10 or newer):

```bash
pip install -r requirements.txt
python listings.py "Toronto, ON" --limit 200 --out toronto.xlsx
python listings.py "Vancouver, BC" --type rent --out vancouver-rentals.csv
python agents.py "Oakville, ON" "North York" --limit 100 --out agents.xlsx
```

**Node.js** (20.6 or newer):

```bash
npm install
node --env-file=.env listings.mjs "Calgary, AB" "Edmonton, AB" --limit 300 --out alberta.xlsx
node --env-file=.env listings.mjs "Ottawa, ON" --property commercial --out ottawa-commercial.csv
node --env-file=.env agents.mjs --file locations.example.txt --out agents.csv
```

Both languages take the same options and write byte-identical CSV and JSONL (`--help` prints them). The CSV is UTF-8 with a BOM so Excel shows accents correctly, and text that starts with `=`, `+`, `-` or `@` gets a leading `'` so a spreadsheet never runs it as a formula.

### Listing options

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

### Agent options

| Option | What it does | Default |
|---|---|---|
| locations | A city (`"Toronto, ON"`, `Mississauga`), district or neighbourhood (`"North York"`, `Yorkville`) or province (`Ontario`, `ON`). Spelling must match Realtor.ca | |
| `--file locations.txt` | One location per line ([example](locations.example.txt)) | |
| `--limit N` | Agents per location | `50` |
| `--out FILE`, `--append`, `--max-charge USD`, `--dataset ID`, `--from FILE` | Same as for listings; `--append` keys on `individual_id` | |

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
- **Brokerage research:** agent counts by brokerage and title in a city, from the agent export.

## Price

Pay per row, billed to your own Apify account. From the Actors' pricing in the Apify API on 2026-10-06 (the Store pages always show the current price):

| | Realtor.ca Property Scraper | Realtor.ca Agent Scraper |
|---|---:|---:|
| Per listing / per agent | $0.001 ($1.00 per 1,000) | $0.001 ($1.00 per 1,000) |
| Start fee per run, until 2026-10-11 (19:57 UTC) | $0.00005 | $0.00005 |
| Start fee per run from 2026-10-11 (19:57 UTC) | $0.002 | $0.003 |

The per-row price is the same on every Apify plan and doesn't change on October 11; only the start fee does. The start fee is charged once per run per GB of memory, and runs use 1 GB by default. The Actors' compute isn't billed on top.

| Example | Rows | Until 2026-10-11 | From 2026-10-11 |
|---|---:|---:|---:|
| One city, `--limit 100` (the default) | 100 listings | $0.10 | $0.10 |
| One city, `--limit 500` (the maximum) | 500 listings | $0.50 | $0.50 |
| All 18 cities, `--limit 500 --max-charge 10` | up to 9,000 listings | up to $9.00 | up to $9.00 |
| Daily new-listings check, one city, `--limit 200`, 30 days | 6,000 listings | $6.00 | $6.06 |
| One location, `--limit 50` (the default) | 50 agents | $0.05 | $0.05 |
| Three cities, 1,000 agents | 1,000 agents | $1.00 | $1.00 |

Amounts are rounded to the cent. Apify's Free plan includes monthly usage credit you can spend on these Actors. `--max-charge` (default $2) is a hard cap Apify enforces for each run: rows past the cap aren't collected, and the scripts warn before starting a run that could go past it. Raise it for big runs, as in the 18-city row.

## Limits

**Listings**

- **18 cities:** Toronto, Mississauga, Brampton, Vaughan, Markham, Oakville, Hamilton, Ottawa, London, Vancouver, Surrey, Burnaby, Victoria, Calgary, Edmonton, Montreal, Winnipeg and Halifax. No postal codes, provinces or neighbourhoods. Write Montreal without the accent.
- **Each city is a fixed map area around its centre,** about 24 km north to south. It can miss the outer edges of a large city and take in part of a neighbour, so check `city_province` when you need exact city limits.
- **Up to 500 listings per city per run, newest first.** The Toronto area had 8,114 active residential listings for sale during the sample run (the Actor's log), so one run gets the newest 500 of them. The "new listings" recipe above is how you keep a complete feed over time.
- **Active listings only.** No sold prices and no price history.
- One photo per listing: the main photo URL.
- The sample run above used listings for sale. For rentals or commercial listings, start with `--limit 5` and check that the columns you need are filled.

**Agents**

- Cities, districts, neighbourhoods and provinces. **No postal codes**: Realtor.ca's agent search doesn't take them, and the script stops before starting a run for one. A location Realtor.ca can't resolve (often a spelling variation, like "Godrich" for "Goderich") is reported back to you, not swapped for national results.
- One search returns at most about 500 agents; the Actor sweeps larger cities in slices on its own. About 1,000 agents across three cities take about 30 seconds. Provinces work but are slower than naming cities.
- No emails, ratings or buyer/seller splits: Realtor.ca doesn't publish them.

## FAQ

**Does Realtor.ca have an API?**
Not a public one for listing data. CREA's REALTOR.ca DDF® "is a managed service for REALTORS® and broker owners who are looking to share listings with multiple third-party websites" ([CREA][ddf], read 2026-10-06). If you are a REALTOR® member or build for one, DDF® is the licensed route. Otherwise, the Actors here read what Realtor.ca shows any visitor and return it as data you can call from Python, Node.js, `curl` or any HTTP client.

**Is scraping Realtor.ca legal?**
Here are the facts; this is not legal advice. The Actors read only the listing and directory data Realtor.ca shows any visitor without an account, and never sign in. The listing rows include the listing agents' names and phone numbers, and the agent rows are business contact details that agents publish; both are still personal information under PIPEDA (Canada), GDPR and similar laws. You're responsible for having a lawful reason to collect and keep them, for following REALTOR.ca's terms of use, and for getting consent before you send commercial electronic messages (CASL). For a commercial product, ask a lawyer. More: [Is web scraping legal?][legal]

**Do I need an MLS® feed, CREA membership or a brokerage login?**
No. The Actors don't use any of them.

**Can I get agent email addresses?**
No. Realtor.ca doesn't publish them; an agent's page offers a contact form instead.

**Can I get sold prices or price history?**
No. The Property Scraper returns active listings only. To track asking-price changes, save each day's export to its own file and compare `price_value` by `mls_number`.

**Can I filter by price or bedrooms?**
Not in the Actor. Its inputs are the city, sale or rent, residential or commercial, and the number of listings. Filter the CSV or the Excel sheet.

**Why was my city rejected?**
The Property Scraper covers the 18 cities above. Postal codes, provinces and neighbourhoods aren't inputs, and "Montréal" has to be written "Montreal".

**How fresh is the data?**
Read live from Realtor.ca when the run starts. Schedule the export to keep a file current.

**Can I call it from Claude, ChatGPT or Cursor?**
Yes. The `yugenox-realtor-ca` server in [yugenox-mcp](https://github.com/ArpitGandhi1934/yugenox-mcp) gives AI assistants both Actors as tools through Apify's MCP server.

## Privacy

Agent names and phone numbers are business contact details that agents publish on Realtor.ca, and they're personal information all the same. Keep only what you need, don't resell agent lists, and don't send commercial email or texts without consent (CASL). `--no-agents` drops the listing agents' names and phones from a listings export. This repo commits no real agent rows: the listing sample is pseudonymized and the agent sample is fictional.

## More

- Guides and other public-data examples: [yugenox-data.vercel.app](https://yugenox-data.vercel.app)
- The Actors, with their input forms, output tables and API pages: [Realtor.ca Property Scraper][listings] ([API][listings-api]) and [Realtor.ca Agent Scraper][agents] ([API][agents-api]). n8n and Make can run them with their Apify node.
- More Canadian data (Kijiji, grocery prices, store stock, jobs): [yugenox on Apify][store]
- Use them from AI assistants: [yugenox-mcp](https://github.com/ArpitGandhi1934/yugenox-mcp)

## Samples, testing and legal

- The listing sample comes from one run on 2026-10-06 (Property Scraper build 0.1.11, 20 listings) through `listings.mjs`. The raw rows went through [`samples/redact.jq`](samples/redact.jq), and the .csv and .xlsx were written from them with `--from`. Input keys were checked against both Actors' live input schemas the same day (Agent Scraper build 0.1.22). The agent sample is synthetic.
- MLS®, Multiple Listing Service® and the associated logos are trademarks owned by The Canadian Real Estate Association (CREA). REALTOR®, REALTORS® and the REALTOR® logo are trademarks controlled by CREA. This repo and the Actors are not affiliated with or endorsed by CREA or REALTOR.ca.
- The MIT licence covers the code. The listing data in `samples/` is a short excerpt of public Realtor.ca listings, included to show the output format; it isn't ours to license.
- Made by Yugenox Corporation.

<!-- All apify.com links for this README live below. When the Apify affiliate id exists, append ?fpr=<id> to these URLs only. -->
[listings]: https://apify.com/yugenox/realtor-ca-property-scraper
[listings-api]: https://apify.com/yugenox/realtor-ca-property-scraper/api
[agents]: https://apify.com/yugenox/realtor-ca-agent-scraper
[agents-api]: https://apify.com/yugenox/realtor-ca-agent-scraper/api
[badge-listings]: https://apify.com/actor-badge?actor=yugenox/realtor-ca-property-scraper
[badge-agents]: https://apify.com/actor-badge?actor=yugenox/realtor-ca-agent-scraper
[store]: https://apify.com/yugenox
[signup]: https://console.apify.com/sign-up
[legal]: https://blog.apify.com/is-web-scraping-legal/
[ddf]: https://www.crea.ca/technology/realtor-ca-for-realtors/realtor-ca-tools/realtor-ca-ddf/
