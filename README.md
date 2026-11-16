# Rate Matrix Builder

A small website (a Streamlit app) that turns a carrier **rate card** — a normal
Excel file full of shipping prices — into clean **pricing matrices** that our
transport system, **CargoWrite**, can read.

You upload one Excel file, click a button, and download ready-to-use price
tables. That's the whole job. You do **not** need to write code or understand
the formulas to use it.

---

# Part 1 — For first-time users

If you have never opened this tool before, read this part and nothing else.
Everything below "Part 2" is for the person who maintains the code.

## What problem does this solve?

Carriers (UPS, DHL, DPD, PostNord…) send us their prices as messy spreadsheets,
each laid out differently. CargoWrite needs them in **one consistent format**,
with fuel and other surcharges already added in, and with the cheapest valid
option listed first. Doing that by hand is slow and error-prone. This tool does
it in a few seconds and produces the exact layout CargoWrite expects.

## Words you'll see (plain-language glossary)

| Term | What it means here |
|------|--------------------|
| **Rate card** | The Excel price file a carrier gives us. The thing you upload. |
| **Matrix** | The finished price table this tool produces. One row = one priced shipping option. |
| **Carrier** | The shipping company: UPS DE, UPS NL, UPS GB, DHL, DPD, PostNord, DHL-FENDER (pallets). |
| **Parcel** | A normal box shipment, priced by weight and number of parcels. |
| **Pallet** | A freight shipment on a wooden pallet, priced by destination zone and weight band. Different carrier (DHL-FENDER), different pricing. |
| **Surcharge** | An extra cost added on top of the base rate — most importantly **fuel %**, plus toll, MAUT, admin, mobility. |
| **MAUT** | A road-toll surcharge that some countries (mostly in central Europe) add. Read automatically from the master file; you don't type it in. |
| **CargoWrite** | Our transport management system. It reads the matrix you download and uses it to price real orders. |
| **Bucket / catch-all row** | A safety-net row that catches unusual orders (oversized, too heavy, or a postcode the carrier didn't list) so nothing falls through with no price. Shown highlighted for review. |
| **Variables sheet** | A second tab inside the downloaded file listing the surcharge percentages, so they can be checked or tweaked in Excel later. |

## Before you start

You need two things:

1. **The website link** — the person who set this up will give you a
   `…streamlit.app` URL. Open it in any browser. There's nothing to install.
2. **A rate card to upload.** The usual one is the DSV master workbook
   (file name like `MDK__FENDER__PARCEL_RATES__S2026.xlsx`). For pallets you
   also have a second file (the DHL "pricing with factor" workbook).

## Using the website — step by step

1. **Open the link.** You'll see a sidebar on the left and a main panel.
2. **Upload your rate card.** Use the upload box at the top of the sidebar. The
   tool figures out the format on its own — you don't pick a type.
3. *(Pallets only)* **Upload the pallet rate card** in the separate
   "Pallet rate card (optional)" box just below. Skip this if you only need
   parcels.
4. **Check the fuel %.** Fuel changes monthly, so confirm the box in the sidebar
   shows the current rate. Most other surcharges are read from the file for you.
5. **Pick the countries** you want matrices for. Only countries actually present
   in your file appear in the list.
6. *(Optional)* Open **Advanced** or **📐 Exceptions & buckets** only if someone
   has told you to change carriers, postcode ranges, or oversized-parcel rules.
   You can safely ignore these the first time.
7. **Click ▶ Generate matrices.** Wait a few seconds.
8. **Download** — either one country at a time, **all of them as a single
   combined file**, or everything zipped together.

That's it. If you only ever do steps 1, 4, 5, 7, 8, you're using it correctly.

## Understanding what you downloaded

**Each country produces three files.** They are the *same* prices filtered three
different ways:

| File | What it is | Use it? |
|------|-----------|---------|
| `*_extended` | Every possible row, nothing removed. | For auditing only. |
| `*_optimized` | Cheaper-or-equal duplicates removed *within* each carrier. | Intermediate. |
| `*_minimal` | Cheaper options removed *across all carriers* — only the genuinely best rows survive. | **This is the one to give CargoWrite.** |

There's also a **Combined** download that merges every selected country's
`minimal` table into one sheet, sorted by country then price — handy when you
want a single file instead of many.

### Row colours in the Excel

Some rows are highlighted so you can spot them at a glance:

- **Blue rows** = **pallet** shipments (DHL-FENDER freight). They're priced on
  a completely different basis from parcels, so the colour keeps the two modes
  visually separate in a mixed sheet.
- **Pale amber rows** = **catch-all / bucket** rows — the safety-net rows for
  oversized, overweight, or unlisted-postcode orders. They carry an extra
  surcharge, so ops should glance over them.

Everything else is a normal parcel row.

### The "Variables" tab

Most downloads include a second tab called **Variables** listing the surcharge
percentages (fuel, MAUT, mobility, toll, admin). In the per-country files these
feed live Excel formulas, so changing a number there recalculates the whole
sheet. In the **combined** and **pallet** files the prices are written as fixed
numbers (because surcharges differ by country and can't all share one cell), and
the Variables tab is there for reference only.

## Quick troubleshooting

- **A country I expected isn't in the list.** It isn't present in the uploaded
  file, or it has no data for the carriers you selected.
- **Pallet rows didn't appear.** You didn't upload the separate pallet rate
  card, or that country has no pallet data (those stay parcel-only).
- **Prices look off after I changed fuel.** Re-generate — the sidebar value is
  only applied when you click Generate.
- **It looks stale after an update.** Re-upload the file; the app keys its cache
  on the file's name *and* size.

---

# Part 2 — For whoever maintains the tool

Everything below is developer reference. Day-to-day users can stop here.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Streamlit UI — the only file users interact with |
| `pipeline.py` | Core parcel logic: build, compute, optimize, write Excel |
| `pallet_pipeline.py` | Pallet (DHL freight / EUROCONNECT) matrix builder |
| `master_parser.py` | Parses the DSV master rate card (all countries in one file) |
| `pallet_parser.py` | Parses the DHL "pricing with factor" pallet workbook |
| `requirements.txt` | Python dependencies |

## Two input formats (auto-detected)

**1. Master file** — the DSV "MDK - FENDER - PARCEL RATES" workbook with one
sheet per rate table and all countries in one file. Upload this and the app:
- reads every carrier and country in one pass
- shows only the countries actually present in the file
- reads **MAUT surcharges per country, per carrier** directly from the file
  (the MAUT inputs disappear from the sidebar)
- you still set **fuel %** in the sidebar (fuel is a separate monthly surcharge)

**2. Per-country file** — the older one-country-per-workbook format (tabs named
UPSDE, DHL, DPD, UPSNL, POSTNORD). Still fully supported.

## Carriers supported

UPS DE, UPS NL, UPS GB (UK domestic), DHL, DPD, PostNord, and DHL-FENDER
(pallets).

Special handling baked in:
- **UPS WorldEase (WEA)** — flat per-parcel rate for CH and NO
- **DHL BNL** — "1st parcel + each additional" pricing for BE, LU, NL
- **PostNord** — flat rate per service (B2B / Home / PUDO), multi-country table
- **UPS DE zones** — postcode/zone-based, including alphanumeric zones (ES4/ES5/ES6)

## Deploy to Streamlit Cloud (free)

1. Create a GitHub repo (can be private).
2. Upload all files to the root: `app.py`, `pipeline.py`, `pallet_pipeline.py`,
   `master_parser.py`, `pallet_parser.py`, `requirements.txt`, `README.md`.
3. Go to [share.streamlit.io](https://share.streamlit.io), sign in with GitHub.
4. New app → pick repo → branch → main file `app.py` → Deploy.
5. Share the URL.

## Changing defaults

Carrier defaults and country configs live in `pipeline.py`
(`CARRIER_DEFAULTS`, `COUNTRY_CONFIG`). The master-file sheet names and parsing
logic live in `master_parser.py`. Pallet economics live in `PALLET_DEFAULTS`,
`PALLET_COUNTRY_OVERRIDES`, and `PALLET_MAUT` in `pipeline.py`. Edit and redeploy.

## Exceptions & buckets (add-on)

CargoWrite picks the first matrix row (cheapest-first) whose constraints all fit
an order. A constraint like a **size limit** means a row only matches parcels
at/under that size — so oversized orders need a **bucket** row lower down that
drops the limit, adds a surcharge, and is flagged for review. Without buckets,
some orders match nothing.

In the app, open **📐 Exceptions & buckets** and edit the table — one line per
carrier (or country). Each line says: *for this carrier, the normal size limit
is X metres; anything bigger gets a €Y-per-parcel surcharge.* On Generate, every
output gets:
- the **normal limit** stamped on the cheap base rows
- a **catch-all bucket twin** (limit blank, surcharge added), highlighted in
  **amber** for oversight

Defaults: every UPS carrier (UPDE incl. its 7R9W62 account, UPSNL, UPSGB,
UPSWEA) 1.5 m, +€6 per parcel above it; DPD 1.75 m, +€46.50; DHL-ROS 1 m,
+€8.63. The limit is written to `USER_DEF_NUM_1`.

> **Note (changed):** the old standalone `AWKWARD` flag column has been removed
> from the output. Bucket / catch-all rows are now identified by their **amber
> highlight** rather than a `Y` in a column. The rule engine still supports a
> `flag_col` (it defaults to the now-unexported `AWKWARD` name); point it at a
> real, exported column such as `USER_DEF_TYPE_2` if you want a machine-readable
> marker back.

The mechanism is general: in `pipeline.py`, `apply_exceptions()` takes a list of
rule dicts (constraint column, normal value, surcharge, scope). Today the UI
exposes the size case; the same function already supports flat surcharges and
other constraint columns.

Leave the table empty for a plain matrix with no buckets.

### Future buckets (implemented, off by default)

Two more bucket types live under the same expander, both **off by default**:

**Overflow buckets** — catch orders heavier or with more parcels than the grid.
For each parcel count *n* the matrix gets a row with `MIN_PARCEL=n`,
`MIN_WEIGHT=n×max-each-weight` and **no upper caps**, priced
`overflow rate × n + surcharge × n` and flagged. The overflow rate is the
contract heavy/per-kg rate **you enter** — it is never guessed from the grid.

**Postcode catch-all** — for zoned carriers (e.g. UPS DE), adds a blank-postcode
fallback at the worst zone's rate so a prefix not present in any zone still
matches something, flagged for review.

All three bucket types stack and are added to the *surviving* rows after
optimization, keeping the final list as short as possible. The engine functions
are `apply_exceptions`, `add_overflow_buckets`, and `add_postcode_catchall` in
`pipeline.py` — each takes generic rule dicts, so new constraint columns or
scopes need a new rule, not new logic.

## Pallets (major add-on)

The builder handles **pallets** as a second shipping mode alongside parcels, in
the same output matrix. Pallets use carrier **DHL-FENDER** (service EUROCONNECT)
and are priced on two axes — destination **postcode zone** and a **weight
bucket** — instead of the parcel each-weight/parcel-count grid. Pallet rows are
written in **orange** so they stand out from parcel rows in a mixed sheet.

### Pallet rate card (separate upload)

Pallets come from their own file (the DHL "pricing with factor" workbook): one
sheet, `Country | Zip | Country+Zip | 0,1-100 kg | 100,1-200 kg | … | FTL`. It
covers 31 countries. Upload it in the sidebar under **Pallet rate card
(optional)** — parsed by `pallet_parser.py`. Pallet rows are added for every
selected country that has pallet data; countries without it stay parcel-only.

### Bucket map

The source has 58 fine weight bands; the matrix collapses them into **21
operationally-meaningful pallet buckets**. The collapse is encoded in
`PALLET_BUCKET_MAP` in `pipeline.py` as `(output MAX_WEIGHT, source band
upper-bound)` pairs and is applied to every country. `MIN_WEIGHT` is left blank
so CargoWrite's cheapest-first matching picks the smallest bucket that fits.

### Pricing chain (matches the combi reference exactly)

For each pallet row:
- `RATE_BASE = FACTORED RATE PALLET ÷ FACTOR`
- `Mobility = mobility% × RATE_BASE` (default 4%)
- `BASE_TOTAL = RATE_BASE + Mobility`
- `FUEL = fuel% × BASE_TOTAL` (default 15% — DHL freight)
- `TOLL = toll% × BASE_TOTAL` (per-country, UK 0.43%)
- `ADMIN = flat € per shipment` (per-country, UK €46.51)
- `TOTAL = RATE_BASE + RATE_EXTRA + Mobility + FUEL + TOLL + ADMIN`

Fuel/mobility/factor are written as **formulas** referencing the Variables sheet
in per-country files (tweak once, recalculates everywhere). Toll/admin are
**per-country literal values** so a single matrix can carry several countries
with different pallet surcharges. All four are editable in the sidebar
(**Pallet surcharges**) and configurable per country via
`PALLET_COUNTRY_OVERRIDES`.

### Columns

Pallet output carries: `FACTORED RATE PALLET`, `USER_DEF_TYPE_1` (PARCEL/PALLET),
`USER_DEF_TYPE_2` (Single/Multi — mirrors RATE_TYPE for parcels), `Mobility`,
`TOLL`, `ADMIN`. Parcel rows leave the pallet-only columns blank; pallet rows
leave the parcel-grid columns (MAX_PARCEL/EACH_WEIGHT/volumes/MAUT/Linehaul)
blank.

### Architecture notes

- `build_pallet_df()` emits zone × bucket rows; `PALLET_DEFAULTS` hold the
  economics.
- Pallets **bypass the parcel optimizer** — every (zone, bucket) is a distinct
  CargoWrite match target, and the parcel dominance logic assumes numeric
  postcodes/weights that pallets don't share. They are split out before
  optimization and re-attached before writing.
- `RATE_TYPE` (Single/Multi) was added for UPDE and UPSGB STANDARD rows; it is
  mirrored into `USER_DEF_TYPE_2`.

## Combined export (all countries in one sheet)

Alongside the per-country files and the ZIP, the app offers a **combined**
download: every selected country's minimal matrix merged into a single sheet
(`Combined_Matrix_minimal.xlsx`), sorted by country then price.

Key design point: the combined sheet writes computed columns as **numeric
values, not formulas**. A single sheet can't reference one Variables cell for a
surcharge that varies by country (MAUT differs per country; pallet toll/admin
are UK-only), so formulas would silently apply one country's rate to all.
Numeric values keep every country correct in one file. A Variables sheet is
still included for reference. The engine function is
`pipeline.write_combined_matrix(frames, path, variables_layout)`; the app builds
it from each country's persisted minimal frame (`result['minimal_df']`).

---

# Appendix — Full function reference (current PALLET2 build)

What every function does, file by file. `_name` = internal helper, not meant to be
called from outside its module. Read this **with** the architecture notes above,
not instead of them.

**Live import graph:** `app.py` → `pipeline.py` + `master_parser.py` +
`pallet_parser.py`. That is the whole running app.

> ⚠️ **`pallet_pipeline.py` is NOT used by the app** — nothing imports it. It is an
> earlier, standalone pallet builder kept in the repo for history. The pallets you
> see in real output are built by `pipeline.build_pallet_df()`, fed by
> `pallet_parser`. Safe to delete `pallet_pipeline.py` on handover; it only causes
> confusion (two modules that look like they build pallets — only one does).

---

## `app.py` — the Streamlit UI (the only file users touch)

Module constants: `ALL_COUNTRIES`, `CARRIER_LABELS`, `FUEL_CARRIERS`,
`MAUT_CARRIERS`, `EXPRESS_CARRIERS`, `DEFAULT_EXCEPTIONS`, `DEFAULT_OVERFLOW`,
`DEFAULT_POSTCODE`, `DEFAULT_HEAVY`, `COLS`.

| Function | What it does |
|----------|--------------|
| `_pct_input(label, key, default)` | One sidebar percentage field; stores/reads the value as a fraction in session state. |
| `variables_layout(fuel_vals, maut_dhl, maut_dpd, pallet_vals=None)` | Assembles the **Variables sheet** rows (fuel, MAUT, mobility, toll, admin, factor) from current sidebar values, to embed in every workbook. |
| `carrier_defaults(fuel_vals, maut_dhl, maut_dpd)` | Turns sidebar values into the per-carrier defaults dict the pipeline expects. |
| `country_cfg_with_overrides(country)` | Base country config from `pipeline.COUNTRY_CONFIG` merged with any UI overrides. |
| `persist(result)` | Copies a country's generated files out of the build temp dir into a longer-lived one so downloads survive Streamlit reruns. |
| `express_rename(result)` | Renames the three matrix files with an `_EXPRESS_` marker for the express-only build. |
| `file_bytes(p)` | Reads a file into bytes for a download button. |
| `exception_rules_from_editor(edited_df)` | Converts the **size-limit** editor table into `apply_exceptions` rule dicts. |
| `overflow_rules_from_editor(edited_df)` | Converts the **overflow** editor table into `add_overflow_buckets` rule dicts. |
| `postcode_rules_from_editor(edited_df)` | Converts the **postcode** editor table into `add_postcode_catchall` rule dicts. |
| `heavy_rules_from_editor(edited_df)` | Converts the **heavy per-kg rate** editor table into overflow rate rules. |
| `_default_pallet_maut_df()` | Seed table for the editable pallet-MAUT grid. |
| `pallet_maut_from_editor(edited_df)` | Converts the pallet-MAUT editor table into the `{country: …}` map `build_pallet_df` expects. |
| `make_zip(results)` | Bundles every country's three files into one ZIP. |
| `make_combined(results, variables_layout_rows, stage='minimal', …)` | Builds the single **all-countries combined** workbook for the chosen stage. |
| `render_results(results, heading, kp, fname_prefix, *, caption=None)` | Renders the summary + per-country + bulk-download UI block for one result set (normal or express). |

---

## `pipeline.py` — parcel core: parse → build → compute → optimize → write

Config constants you'd actually edit: `CARRIER_DEFAULTS`, `COUNTRY_CONFIG`,
`PALLET_DEFAULTS`, `PALLET_COUNTRY_OVERRIDES`, `PALLET_MAUT`. Layout/order
constants (rarely touched): `VARIABLES_LAYOUT`, `CARRIER_BUILDERS`,
`COLUMN_ORDER`, `COL_LETTER`, `PALLET_COLUMN_ORDER`, `EUROCONNECT_BUCKET_*`.

### Parsing — legacy per-country workbook
| Function | What it does |
|----------|--------------|
| `_default_country_cfg(iso2)` | Fallback config for a country not in `COUNTRY_CONFIG`. |
| `_norm(v)` | Normalise any cell: lowercase, collapse punctuation/whitespace to single spaces. |
| `_cell_match(cell_value, *needles)` | True if the normalised cell contains/equals any normalised needle. |
| `_parse_float(v)` | Robust float parse — comma decimals, currency symbols, `None`. |
| `_find_sheet(wb, *name_hints)` | First sheet whose name fuzzy-matches any hint. |
| `_scan_anchor(ws, *anchor_texts, …)` | First cell fuzzy-matching an anchor; used to locate tables. |
| `_find_from_to(ws, …)` | Locate the From/To weight-band header row (multi-language, wide window). |
| `_extract_tiers(ws, …)` | Read weight-band tiers, tolerating blank spacers, comma-decimals, "over X". |
| `_extract_rates_by_zone(ws, hrow, from_col, to_col)` | Read per-zone rate columns into tiers. |
| `_parse_upsde_zones(ws)` | Parse the UPS DE postcode→zone table (incl. alphanumeric ES4/5/6 zones). |
| `_parse_postnord_sheet(ws)` | Four-strategy PostNord sheet parser. |
| `parse_rate_cards(excel_path)` | **Entry point** for the legacy per-country file; returns the parsed rate dict. |

### Row building — one builder per carrier
| Function | What it does |
|----------|--------------|
| `lookup_tier_rate(tiers, weight)` | Rate for a given weight from a tier list. |
| `collapse_same_rate_tiers(tiers, weight_cap=None)` | Merge adjacent tiers sharing a rate. |
| `_upde_service_buckets(rate_data, service_key, country_cfg)` | Group a UPDE service into `(postcode_prefix, tiers)`. |
| `_common(site, client, carrier, iso2)` | Shared CargoWrite key fields for a row. |
| `build_combined_weight_rows(…)` | Emit **Model-B** rows: one rate on the whole-shipment payweight. |
| `build_rows_upde(rate_data, country_cfg)` | UPS DE — zoned, Single/Multi `RATE_TYPE`. |
| `build_rows_dhl(rate_data, country_cfg)` | DHL — BNL "first + each additional", else Other-countries tiers. |
| `build_rows_dpd(rate_data, country_cfg)` | DPD — genuine per-parcel pricing. |
| `build_rows_upsnl(rate_data, country_cfg)` | UPS NL — Express Saver by zone. |
| `build_rows_postnord(rate_data, country_cfg)` | PostNord — flat rate per service (B2B / Home / PUDO). |
| `build_rows_upsgb(rate_data, country_cfg)` | UPS GB (UK domestic) — STDS single/per-parcel, STDM combined, EXPS; rates in **GBP**. |
| `build_extended_matrix(parsed, country_cfg)` | Run every applicable builder → the full **extended** row set. |

### Compute & write (Excel)
| Function | What it does |
|----------|--------------|
| `compute_numeric_totals(df, carrier_defaults=None)` | Fill `TOTAL_PRICE` and surcharge columns as plain numbers. |
| `_build_formulas_for_row(row_dict, excel_row, carrier_defaults=None)` | Build the live Excel formulas for one row. |
| `write_matrix_excel(df, output_path, country_cfg, …)` | Write a per-country workbook with formulas + Variables sheet. |
| `_letter_map`, `_ensure_var`, `_update_formula`, `_write_filtered_excel`, `_ensure_numeric`, `_recompute_total` | Excel/formula plumbing helpers (column letters, ensuring a Variables row exists, rewriting cell refs after row deletion, numeric coercion, total recompute). |

### Optimizers — the three stages
| Function | What it does |
|----------|--------------|
| `optimize_matrix(df)` | **Within-carrier** dominance (groups by carrier+service+postcode). Produces the *optimized* stage. |
| `optimize_globally(input_path, output_path)` | **Cross-carrier** dominance via an Excel round-trip. Produces the *minimal* stage. |
| `optimize_globally_df(df)` | Same cross-carrier logic in pure pandas (postcode-partitioned for speed) so buckets can be added before writing. **This is the one the live path uses.** |

> Known structural limit (carried over): the global optimizer is **service-level
> blind**, so cheaper STANDARD rows can dominate out EXPRESS SAVER rows. The
> express-only build mode is the current workaround.

### Buckets & exceptions (all take generic rule dicts)
| Function | What it does |
|----------|--------------|
| `apply_exceptions(df, rules)` | Stamp size limits on in-scope base rows + append amber **bucket twins**. |
| `add_overflow_buckets(df, rules, …)` | Append heavy / extra-parcel overflow rows (rate×n + surcharge×n, no upper caps). |
| `add_postcode_catchall(df, rules, …)` | For zoned carriers, append a blank-postcode fallback at the worst-zone rate. |

### Pallets & orchestration
| Function | What it does |
|----------|--------------|
| `_pallet_maut_for(country, ceiling_kg, maut_table)` | Pick the right two-tier MAUT for a pallet weight. |
| `build_pallet_df(country, zip_rate_map, band_ceilings, …)` | **The live pallet builder** — emits DHL-FENDER zone×bucket rows; returns `(df, maut_known)`. |
| `collapse_pallet_bands(df_pallet)` | Within each (country, zip), drop a band whose `RATE_BASE` equals the band above (redundant). |
| `_align_columns(frames)` | Concat frames on a union of columns (pallet order when pallet rows present). |
| `write_matrix_numeric(df, …)` | Write a matrix as **numeric values** — used whenever pallet rows are present. |
| `write_matrix_with_formulas(df, …)` | Write a pallet-inclusive matrix with **formulas** referencing the Variables sheet. |
| `write_combined_matrix(frames, …)` | Merge all countries' minimal frames into one sheet; always `formulas=False`. |
| `append_euroconnect_buckets(df, …)` | Append one EUROCONNECT catch-all row per country (`MAX_WEIGHT 24000`, `TOTAL_PRICE 999999`). |
| `run_pipeline(input_path, country, …)` | Full build for one country from a **legacy file path**. |
| `run_pipeline_from_parsed(parsed, country, output_dir, cfg, …)` | **Main orchestrator** — full build for one country from an already-parsed dict (master-file path; supports `express_only`). |

---

## `master_parser.py` — the DSV master workbook (all countries, one file)

The current real-world input. One sheet per rate table; parsed once, then sliced
per country into the same shape `parse_rate_cards()` produces.

| Function | What it does |
|----------|--------------|
| `is_master_file(path)` | True if the workbook looks like the DSV master (one sheet per table). |
| `_scan_from_to(ws)` | Find the From/To header anywhere in a sheet. |
| `_find_label_row(ws, *labels)` | First row containing **all** the given labels → `(row, {label: col})`. |
| `_extract_zone_tiers(ws, hrow, from_col, to_col)` | `{zone_key: tiers}` for every zone column right of "To". |
| `_flat_country_rates(ws, value_col_offset=1)` | For "XX rate" sheets (WEA, etc.) → `{ISO2: rate}`. |
| `_parse_upsde_zones(ws)` | `ZONES UPSDE` → per-country postcode bands with zone per service. |
| `_parse_upsnl_zones(ws)` | `ZONES UPSNL EXPRESS` → `{ISO2: zone}` (Express Saver). |
| `_parse_dpd(ws)` | `PARCEL - DPD` → `{ISO2: {normal, small}}`. |
| `_parse_maut(ws)` | `MAUT SURCHARGE` → `{ISO2: {DHL-ROS, DPD}}` (read automatically). |
| `_parse_dhl_other(ws)` | `PARCEL - DHL - Other countries` → `{ISO2: tiers}`. |
| `_parse_dhl_bnl(ws)` | `PARCEL - DHL - BNL` → `{ISO2: {first, after}}`. |
| `_parse_postnord(ws)` | `PARCEL - POSTNORD - STD` → `{ISO2: {B2B, HOME, PUDO}}`. |
| `_parse_gb_tiers(ws)` | Single-rate-column tier table (UPSGB STDS/STDM/EXPS). |
| `_parse_linehaul(ws)` | First numeric value next to a Carrier/UPS label. |
| `parse_master_rate_card(path)` | **Entry point** — parse the whole master workbook once into a structured dict. |
| `country_rate_data(master, iso2)` | Slice the master into one country's parsed dict (pipeline-ready shape). |
| `country_maut(master, iso2)` | `{DHL-ROS, DPD}` MAUT for a country (0.0 if absent). |
| `available_countries(master)` | All ISO2 codes that have any rate data. |

---

## `pallet_parser.py` — the DHL "pricing with factor" workbook

| Function | What it does |
|----------|--------------|
| `is_pallet_factor_file(path)` | True if this is the DHL factor workbook (Country / Zip / "<n> kg" columns). |
| `_band_ceiling(header_value)` | `"100,1 - 200 kg"` → `200`; `"FTL"` → `None`. |
| `parse_pallet_factor_file(path)` | **Entry point** — parse the whole workbook once. |
| `available_pallet_countries(parsed)` | Countries that have pallet data. |
| `country_pallet_data(parsed, iso2)` | `{zip_prefix: {band_kg: rate}}` for one country. |
| `band_ceilings(parsed)` | Ordered list of weight-band ceilings. |

---

## `pallet_pipeline.py` — ⚠️ legacy, not wired in

Standalone pallet builder from before pallets were folded into `pipeline.py`.
**Not imported anywhere.** Functions (`build_rows_pallet`, `optimize_pallet`,
`write_pallet_excel`, `run_pallet_pipeline`, `_maut_for`) duplicate logic now
living in `pipeline.build_pallet_df` / `collapse_pallet_bands` /
`write_matrix_*`. Delete on handover unless you want it as a historical reference.
