"""
app.py — Rate Matrix Builder  ·  Streamlit UI

Supports two input formats, auto-detected on upload:
  • MASTER file  — the DSV "MDK - FENDER - PARCEL RATES" workbook (all countries
    in one file). MAUT is read per-country from the file automatically.
  • Per-country file — the older one-country-per-workbook format.

Optional pallet rate card (DHL-FENDER / EUROCONNECT) can be uploaded alongside;
pallet rows are merged into every selected country's matrix.

Run locally:   streamlit run app.py
Deploy:        push to GitHub, connect on share.streamlit.io
"""

import io
import logging
import shutil
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

import pandas as pd
import streamlit as st

import pipeline as pl
import master_parser as mp
import pallet_parser as pp

st.set_page_config(page_title="Rate Matrix Builder", page_icon="📦",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
  .block-container { padding-top: 1.5rem; }
  .section-title { font-size:.75rem; font-weight:700; letter-spacing:.12em;
                   text-transform:uppercase; color:#888; margin:1.4rem 0 .5rem 0; }
  div[data-testid="column"] .stCheckbox label { font-size:.85rem; }
</style>
""", unsafe_allow_html=True)

logging.basicConfig(level=logging.INFO)

ALL_COUNTRIES  = sorted(pl.COUNTRY_CONFIG.keys())
CARRIER_LABELS = {cid: cfg['label'] for cid, cfg in pl.CARRIER_DEFAULTS.items()}
FUEL_CARRIERS  = ['UPDE', 'DHL-ROS', 'DPD', 'UPSNL', 'POSTNORD', 'UPSGB']
MAUT_CARRIERS  = ['DPD', 'DHL-ROS']
EXPRESS_CARRIERS = ['UPDE', 'UPSNL', 'UPSGB']   # the only carriers quoting EXPRESS SAVER


# ── Helpers ───────────────────────────────────────────────────────────────────

def _pct_input(label, key, default):
    return st.number_input(label, min_value=0.0, max_value=1.0,
                           value=float(default), step=0.01, format="%.2f", key=key)


def variables_layout(fuel_vals, maut_dhl, maut_dpd, pallet_vals=None):
    pv = pallet_vals or {}
    return [
        ('FUEL UPSDE',    fuel_vals.get('UPDE',     0.27)),   # B1
        ('FUEL DHL',      fuel_vals.get('DHL-ROS',  0.27)),   # B2
        ('FUEL DPD',      fuel_vals.get('DPD',      0.27)),   # B3
        ('FUEL UPSNL',    fuel_vals.get('UPSNL',    0.27)),   # B4
        ('FUEL POSTNORD', fuel_vals.get('POSTNORD', 0.27)),   # B5
        ('FUEL UPSGB',    fuel_vals.get('UPSGB',    0.27)),   # B6
        (None, None),                                         # B7
        ('MAUT DPD',      maut_dpd),                          # B8
        ('MAUT DHL',      maut_dhl),                          # B9
        (None, None),                                         # B10
        ('FUEL DHL PALLET', pv.get('fuel',     0.155)),       # B11
        ('MOBILITY PALLET', pv.get('mobility', 0.04)),        # B12
        ('TOLL UK PALLET',  pv.get('toll',     0.0043)),      # B13
        ('ADMIN PALLET',    pv.get('admin',    46.51)),       # B14
        ('FACTOR DHL',      pv.get('factor',   4.13)),         # B15
    ]


def carrier_defaults(fuel_vals, maut_dhl, maut_dpd):
    cd = deepcopy(pl.CARRIER_DEFAULTS)
    for cid, pct in fuel_vals.items():
        if cid in cd:
            cd[cid]['fuel_pct'] = pct
    cd['DHL-ROS']['maut_pct'] = maut_dhl
    cd['DPD']['maut_pct']     = maut_dpd
    return cd


def country_cfg_with_overrides(country):
    cfg = deepcopy(pl.COUNTRY_CONFIG[country])
    ov  = st.session_state.get('country_overrides', {}).get(country)
    if ov:
        cfg.update({
            'carriers':              ov['carriers'],
            'max_parcel_count':      ov['max_parcel_count'],
            'max_each_weight_kg':    ov['max_each_weight_kg'],
            'postcode_prefix_range': tuple(ov['postcode_prefix_range']),
            'each_weight_grid':      sorted(set(
                list(range(1, int(ov['max_each_weight_kg']) + 1))
                + [ov['max_each_weight_kg']])),
        })
    return cfg


def persist(result):
    """Copy result files out of a temp dir into a longer-lived temp dir."""
    out = tempfile.mkdtemp()
    for key in ('extended', 'optimized', 'minimal'):
        dst = Path(out) / Path(result[key]).name
        shutil.copy(result[key], dst)
        result[key] = str(dst)
    # Carry the numeric stage frames (used to build the combined workbooks)
    for key in ('extended_df', 'optimized_df', 'minimal_df'):
        if result.get(key):
            dst = Path(out) / Path(result[key]).name
            shutil.copy(result[key], dst)
            result[key] = str(dst)
    return result


def express_rename(result):
    """Rename the three matrix files to carry an _EXPRESS_ marker so the
    express downloads can't be confused with a standard run."""
    for key in ('extended', 'optimized', 'minimal'):
        p = Path(result[key])
        new = p.with_name(p.name.replace('_Matrix_', '_EXPRESS_Matrix_'))
        shutil.move(str(p), str(new))
        result[key] = str(new)
    return result


def file_bytes(p):
    return Path(p).read_bytes()


DEFAULT_EXCEPTIONS = pd.DataFrame([
    {'Enabled': True, 'Carrier': 'UPSGB', 'Country (blank=all)': '',
     'Service level (blank=all)': 'STANDARD',
     'Size limit (m)': 1.5,  'Surcharge €/parcel': 6.0},
    {'Enabled': True, 'Carrier': 'DPD',  'Country (blank=all)': '',
     'Service level (blank=all)': '',
     'Size limit (m)': 1.75, 'Surcharge €/parcel': 46.5},
     {'Enabled': True, 'Carrier': 'DHL-ROS',  'Country (blank=all)': '',
     'Service level (blank=all)': '',
     'Size limit (m)': 1, 'Surcharge €/parcel': 8.63},
])


def exception_rules_from_editor(edited_df):
    rules = []
    for _, row in edited_df.iterrows():
        if not bool(row.get('Enabled', False)):
            continue
        carrier = str(row.get('Carrier', '') or '').strip()
        country = str(row.get('Country (blank=all)', '') or '').strip().upper()
        service = str(row.get('Service level (blank=all)', '') or '').strip().upper()
        try:
            limit = float(row.get('Size limit (m)'))
        except (TypeError, ValueError):
            continue
        try:
            sur = float(row.get('Surcharge €/parcel') or 0)
        except (TypeError, ValueError):
            sur = 0.0
        rules.append({
            'enabled':        True,
            'label':          f'Oversize {carrier or "ALL"}',
            'carriers':       [carrier] if carrier and carrier != '(all)' else [],
            'countries':      [c.strip() for c in country.split(',') if c.strip()],
            'service_levels': [s.strip() for s in service.split(',') if s.strip()],
            'constraint_col': 'USER_DEF_TYPE_1',
            'normal_value':   limit,
            'bucket_value':   None,
            'flag_col':       'AWKWARD',
            'flag_value':     'y',
            'surcharge':      sur,
            'surcharge_mode': 'per_parcel',
        })
    return rules


DEFAULT_OVERFLOW = pd.DataFrame([
    {'Enabled': False, 'Carrier': 'UPDE', 'Country (blank=all)': '',
     'Overflow rate €/parcel': 0.0, 'Surcharge €/parcel': 6.0},
])

DEFAULT_POSTCODE = pd.DataFrame([
    {'Enabled': False, 'Carrier': 'UPDE', 'Country (blank=all)': '',
     'Surcharge €': 0.0},
])


def overflow_rules_from_editor(edited_df):
    rules = []
    for _, row in edited_df.iterrows():
        if not bool(row.get('Enabled', False)):
            continue
        try:
            rate = float(row.get('Overflow rate €/parcel'))
        except (TypeError, ValueError):
            continue
        if rate <= 0:
            continue
        carrier = str(row.get('Carrier','') or '').strip()
        country = str(row.get('Country (blank=all)','') or '').strip().upper()
        rules.append({
            'enabled': True,
            'carriers': [carrier] if carrier and carrier != '(all)' else [],
            'countries': [c.strip() for c in country.split(',') if c.strip()],
            'overflow_rate': rate,
            'surcharge': float(row.get('Surcharge €/parcel') or 0),
            'flag_col': 'AWKWARD', 'flag_value': 'Y',
        })
    return rules


def postcode_rules_from_editor(edited_df):
    rules = []
    for _, row in edited_df.iterrows():
        if not bool(row.get('Enabled', False)):
            continue
        carrier = str(row.get('Carrier','') or '').strip()
        country = str(row.get('Country (blank=all)','') or '').strip().upper()
        rules.append({
            'enabled': True,
            'carriers': [carrier] if carrier and carrier != '(all)' else [],
            'countries': [c.strip() for c in country.split(',') if c.strip()],
            'surcharge': float(row.get('Surcharge €') or 0),
            'flag_col': 'AWKWARD', 'flag_value': 'Y',
        })
    return rules


# ── Heavy / oversized parcel (per-parcel weight surcharge) ────────────────────
# A per-parcel weight threshold. Unlike the oversize (size) rule it must not
# overwrite EACH_WEIGHT, so it maps to a mode='threshold' exception rule that
# surcharges, in place, every row whose per-box cap reaches the threshold.
DEFAULT_HEAVY = pd.DataFrame([
    {'Enabled': True, 'Carrier': 'DHL-ROS', 'Country (blank=all)': '',
     'Weight threshold kg': 20.0, 'Surcharge €/parcel': 4.89},
])


def heavy_rules_from_editor(edited_df):
    rules = []
    for _, row in edited_df.iterrows():
        if not bool(row.get('Enabled', False)):
            continue
        try:
            thr = float(row.get('Weight threshold kg'))
        except (TypeError, ValueError):
            continue
        try:
            sur = float(row.get('Surcharge €/parcel') or 0)
        except (TypeError, ValueError):
            sur = 0.0
        carrier = str(row.get('Carrier', '') or '').strip()
        country = str(row.get('Country (blank=all)', '') or '').strip().upper()
        rules.append({
            'enabled':        True,
            'mode':           'threshold',
            'label':          f'Heavy parcel {carrier or "ALL"} ≥{thr:g}kg',
            'carriers':       [carrier] if carrier and carrier != '(all)' else [],
            'countries':      [c.strip() for c in country.split(',') if c.strip()],
            'constraint_col': 'EACH_WEIGHT',
            'threshold':      thr,
            'flag_col':       'AWKWARD',
            'flag_value':     'y',
            'surcharge':      sur,
            'surcharge_mode': 'per_parcel',
        })
    return rules


# ── Pallet MAUT editor helpers ────────────────────────────────────────────────

def _default_pallet_maut_df():
    rows = []
    for iso, (low, high, tier) in pl.PALLET_MAUT.items():
        rows.append({'Country': iso, 'MAUT % (≤ tier)': low,
                     'MAUT % (> tier)': high, 'Tier kg': tier})
    return pd.DataFrame(rows)


def pallet_maut_from_editor(edited_df):
    table = {}
    for _, row in edited_df.iterrows():
        iso = str(row.get('Country', '') or '').strip().upper()
        if not iso:
            continue
        try:
            low  = float(row.get('MAUT % (≤ tier)') or 0)
            high = float(row.get('MAUT % (> tier)') or 0)
            tier = float(row.get('Tier kg') or 2500)
        except (TypeError, ValueError):
            continue
        table[iso] = (low, high, tier)
    return table


def make_zip(results):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for country, r in results.items():
            for key, label in [('extended', 'extended'),
                               ('optimized', 'optimized'),
                               ('minimal', 'minimal')]:
                zf.write(r[key], f'{country}/{country}_Matrix_{label}.xlsx')
    buf.seek(0)
    return buf.read()


def make_combined(results, variables_layout_rows, stage='minimal',
                  pallet_maut=None, pallet_defaults=None, carrier_defaults=None):
    """Build ONE workbook with every country's matrix (for the given stage:
    'extended', 'optimized' or 'minimal') in a single sheet. Written numerically
    so each country keeps its own per-country surcharges (MAUT differs by country
    and can't be a single shared Variables formula)."""
    key = f'{stage}_df'
    frames = []
    for country, r in results.items():
        p = r.get(key)
        if p and Path(p).exists():
            try:
                frames.append(pd.read_pickle(p))
            except Exception:
                pass
    if not frames:
        return None
    out = Path(tempfile.mkdtemp()) / f'Combined_Matrix_{stage}.xlsx'
    # Numeric, not formulas: a single shared sheet can't carry per-country MAUT
    # (DPD/DHL differ by country). Formulas would reference one Variables cell and
    # apply the same % to every country. Numeric values keep each country correct.
    pl.write_combined_matrix(frames, out, variables_layout_rows,
                             pallet_maut=pallet_maut, pallet_defaults=pallet_defaults,
                             carrier_defaults=carrier_defaults, formulas=False)
    return Path(out).read_bytes()


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 📦 Rate Matrix Builder")
    st.caption("Fender Musical Instruments · Logistics")
    st.divider()

    st.markdown('<p class="section-title">Rate card</p>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload Excel file", type=["xlsx", "xls"],
                                label_visibility="collapsed",
                                help="Upload the DSV master rate card, or an older "
                                     "per-country file. The format is detected automatically.")

    is_master = False
    master = None
    if uploaded is not None:
        if st.session_state.get('uploaded_name') != uploaded.name:
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                tmp.write(uploaded.read())
                st.session_state['input_path']    = tmp.name
                st.session_state['uploaded_name'] = uploaded.name
                for key in ('master', 'parsed_per_country', 'parse_warnings',
                            'country_overrides'):
                    st.session_state.pop(key, None)
        input_path = st.session_state['input_path']

        is_master = mp.is_master_file(input_path)
        if is_master:
            if 'master' not in st.session_state:
                with st.spinner("Reading master rate card…"):
                    st.session_state['master']         = mp.parse_master_rate_card(input_path)
                    st.session_state['parse_warnings'] = []
            master = st.session_state['master']
            avail = mp.available_countries(master)
            st.success(f"Master rate card detected — {len(avail)} countries available. "
                       "MAUT is read per-country from the file.")
        else:
            if 'parsed_per_country' not in st.session_state:
                with st.spinner("Reading rate card…"):
                    st.session_state['parsed_per_country'] = pl.parse_rate_cards(input_path)
                    st.session_state['parse_warnings']     = []
            st.info("Per-country rate card detected.")

    # ── Pallet rate card (optional, separate file) ────────────────────────────
    st.markdown('<p class="section-title">Pallet rate card (optional)</p>',
                unsafe_allow_html=True)
    pallet_uploaded = st.file_uploader(
        "Upload DHL pallet Excel", type=["xlsx", "xls"],
        label_visibility="collapsed", key="pallet_uploader",
        help="Upload the DHL pallet rate card (Country/Zip × weight bands). "
             "Pallet rows (carrier DHL-FENDER) are added for every selected "
             "country that has pallet data.")
    if pallet_uploaded is not None:
        if st.session_state.get('pallet_uploaded_name') != pallet_uploaded.name:
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                tmp.write(pallet_uploaded.read())
                st.session_state['pallet_path']          = tmp.name
                st.session_state['pallet_uploaded_name'] = pallet_uploaded.name
                st.session_state.pop('pallets', None)
        if 'pallets' not in st.session_state:
            with st.spinner("Reading pallet rate card…"):
                # canonical name in pallet_parser; alias parse_pallet_rate_card also works
                st.session_state['pallets'] = pp.parse_pallet_factor_file(
                    st.session_state['pallet_path'])
        _pcount = len(pp.available_pallet_countries(st.session_state['pallets']))
        st.success(f"Pallet card loaded — {_pcount} countries with pallet rates.")

    st.markdown('<p class="section-title">Fuel surcharge (%)</p>', unsafe_allow_html=True)
    fuel_vals = {}
    cols = st.columns(2)
    for i, cid in enumerate(FUEL_CARRIERS):
        with cols[i % 2]:
            fuel_vals[cid] = _pct_input(CARRIER_LABELS[cid], f'fuel_{cid}',
                                        pl.CARRIER_DEFAULTS[cid]['fuel_pct'])

    maut_dhl = pl.CARRIER_DEFAULTS['DHL-ROS']['maut_pct']
    maut_dpd = pl.CARRIER_DEFAULTS['DPD']['maut_pct']
    if not is_master:
        st.markdown('<p class="section-title">MAUT surcharge (%)</p>', unsafe_allow_html=True)
        cols = st.columns(2)
        with cols[0]:
            maut_dpd = _pct_input('DPD', 'maut_DPD', maut_dpd)
        with cols[1]:
            maut_dhl = _pct_input('DHL', 'maut_DHL', maut_dhl)

    # Pallet surcharge inputs — only shown when a pallet card is loaded
    pallet_vals = {}
    pallet_maut_table = dict(pl.PALLET_MAUT)
    pallet_max_band_kg = 0          # 0 = no cap; set by the control below
    if st.session_state.get('pallets'):
        st.markdown('<p class="section-title">Pallet surcharges (DHL-FENDER)</p>',
                    unsafe_allow_html=True)
        _pd = pl.PALLET_DEFAULTS['DHL-FENDER']
        cols = st.columns(2)
        with cols[0]:
            pallet_vals['fuel'] = _pct_input('Fuel %', 'pal_fuel', _pd['fuel_pct'])
        with cols[1]:
            pallet_vals['mobility'] = _pct_input('Mobility %', 'pal_mob', _pd['mobility_pct'])
        cols = st.columns(2)
        with cols[0]:
            pallet_vals['toll'] = st.number_input(
                'UK toll %', min_value=0.0, max_value=1.0,
                value=float(pl.PALLET_COUNTRY_OVERRIDES.get('GB', {}).get('toll_pct', 0.0043)),
                step=0.0001, format="%.4f", key='pal_toll')
        with cols[1]:
            pallet_vals['admin'] = st.number_input(
                'Admin € / shipment', min_value=0.0,
                value=float(_pd['admin_per_shipment']), step=1.0, format="%.2f", key='pal_admin')
        pallet_vals['factor'] = st.number_input(
            'Factor', min_value=0.0001,
            value=float(_pd['factor']), step=0.01, format="%.4f", key='pal_factor')
        pallet_max_band_kg = st.number_input(
            'Max pallet weight (kg)  —  0 = no cap', min_value=0,
            value=0, step=500, key='pal_max_band',
            help="Drops every rate-card weight band whose ceiling is above this. "
                 "Bands step e.g. 3500 → 4000 → 4500, so 4025 keeps bands up to "
                 "4000 kg (the 4000,1–4500 band is dropped); 4500 keeps that band. "
                 "0 keeps all bands.")

        st.markdown('<p class="section-title">Pallet MAUT (% of rate)</p>',
                    unsafe_allow_html=True)
        st.caption("Per country. Low = ≤ tier kg, high = above. Add a row for any "
                   "new country — unlisted countries get 0 MAUT (with a warning).")
        if 'pallet_maut_df' not in st.session_state:
            st.session_state.pallet_maut_df = _default_pallet_maut_df()
        _maut_edit = st.data_editor(
            st.session_state.pallet_maut_df, num_rows="dynamic",
            use_container_width=True, hide_index=True,
            column_config={
                'Country': st.column_config.TextColumn(width="small"),
                'MAUT % (≤ tier)': st.column_config.NumberColumn(format="%.4f"),
                'MAUT % (> tier)': st.column_config.NumberColumn(format="%.4f"),
                'Tier kg': st.column_config.NumberColumn(format="%d"),
            }, key='pallet_maut_editor')
        st.session_state.pallet_maut_df = _maut_edit
        pallet_maut_table = pallet_maut_from_editor(_maut_edit)

    st.divider()
    run_btn = st.button("▶ Generate matrices", type="primary",
                        use_container_width=True, disabled=(uploaded is None))
    exp_btn = st.button("⚡ Express-only matrices",
                        use_container_width=True, disabled=(uploaded is None),
                        help="Build a matrix with ONLY the EXPRESS SAVER options — "
                             "UPS DE (7R9W62 + EXPRESS SAVER), UPS NL EXPRESS SAVER, "
                             "UPS GB EXPRESS SAVER. No STANDARD, no DPD/DHL, no pallets. "
                             "Results appear in their own section so a standard run "
                             "isn't overwritten.")
    if uploaded is None:
        st.caption("Upload a rate card first.")


# ── Country selection ─────────────────────────────────────────────────────────

st.markdown("## Select countries")

_parse_warnings = st.session_state.get('parse_warnings', [])
if _parse_warnings:
    with st.expander(f"⚠️ {len(_parse_warnings)} parse warning(s) — some carriers or sheets "
                     f"were not found. Click to see details.", expanded=True):
        for w in _parse_warnings:
            st.warning(w)
elif uploaded is not None:
    st.success("✅ All carriers parsed cleanly — no warnings.")

if is_master and master is not None:
    avail = mp.available_countries(master)
    selectable = [c for c in ALL_COUNTRIES if c in avail]
    if 'GB' in avail and 'GB' not in selectable:
        selectable.append('GB')
    selectable = sorted(set(selectable))
    st.caption(f"Showing the {len(selectable)} countries present in this rate card.")
else:
    selectable = ALL_COUNTRIES
    st.caption("Choose which countries to generate matrices for.")

# Checkbox state lives under the widget keys themselves (chk_<country>). The
# Select all / Clear buttons render *before* the checkboxes, so they may seed
# those keys directly; a keyed checkbox ignores `value=` on rerun and reads its
# state from session_state, which is why writing to a separate dict did nothing.
for c in selectable:
    st.session_state.setdefault(f'chk_{c}', False)

ca, cb, *_ = st.columns([1, 1, 8])
if ca.button("Select all"):
    for c in selectable:
        st.session_state[f'chk_{c}'] = True
if cb.button("Clear"):
    for c in selectable:
        st.session_state[f'chk_{c}'] = False

COLS = 10
grid = st.columns(COLS)
for i, country in enumerate(selectable):
    with grid[i % COLS]:
        st.checkbox(country, key=f'chk_{country}')

selected = [c for c in selectable if st.session_state.get(f'chk_{c}')]

# ── Advanced per-country settings ─────────────────────────────────────────────
if selected:
    with st.expander("⚙️ Advanced — carrier & postcode settings per country", expanded=False):
        st.caption("Defaults are pre-configured. Change only if a country differs.")
        if 'country_overrides' not in st.session_state:
            st.session_state.country_overrides = {}
        for country in selected:
            base = pl.COUNTRY_CONFIG.get(country, pl._default_country_cfg(country))
            ov = st.session_state.country_overrides.setdefault(country, {
                'carriers':              list(base['carriers']),
                'max_parcel_count':      base['max_parcel_count'],
                'max_each_weight_kg':    base['max_each_weight_kg'],
                'postcode_prefix_range': list(base['postcode_prefix_range']),
            })
            st.markdown(f"**{country}**")
            c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 1, 1])
            with c1:
                ov['carriers'] = st.multiselect("Carriers", list(pl.CARRIER_DEFAULTS),
                                                default=ov['carriers'], key=f'car_{country}')
            with c2:
                ov['max_parcel_count'] = st.number_input("Max parcels", 1, 20,
                                                         value=ov['max_parcel_count'], key=f'mp_{country}')
            with c3:
                ov['max_each_weight_kg'] = st.number_input("Max kg", 1.0, 70.0,
                                                           value=float(ov['max_each_weight_kg']),
                                                           step=0.5, key=f'mw_{country}')
            with c4:
                ov['postcode_prefix_range'][0] = st.number_input("PC from", 0, 99,
                                                                 value=ov['postcode_prefix_range'][0], key=f'p0_{country}')
            with c5:
                ov['postcode_prefix_range'][1] = st.number_input("PC to", 0, 99,
                                                                 value=ov['postcode_prefix_range'][1], key=f'p1_{country}')
            st.session_state.country_overrides[country] = ov

# ── Exceptions & buckets (add-on) ─────────────────────────────────────────────
with st.expander("📐 Exceptions & buckets — oversize / surcharges per carrier",
                 expanded=False):
    st.caption(
        "CargoWrite matches each order top-down and takes the first row that "
        "fits. A size limit means a row only matches parcels at/under that size; "
        "anything bigger must fall through to a **bucket** row that drops the "
        "limit, flags it (AWKWARD), and adds a surcharge. Buckets are added to "
        "every output and shown in amber. Leave the table empty for no buckets."
    )
    if 'exceptions_df' not in st.session_state:
        st.session_state.exceptions_df = DEFAULT_EXCEPTIONS.copy()
    edited = st.data_editor(
        st.session_state.exceptions_df,
        num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            'Enabled': st.column_config.CheckboxColumn(width="small"),
            'Carrier': st.column_config.SelectboxColumn(
                options=['(all)'] + list(pl.CARRIER_DEFAULTS), width="small"),
            'Country (blank=all)': st.column_config.TextColumn(
                help="ISO2 code(s), comma-separated. Blank = all countries.", width="small"),
            'Service level (blank=all)': st.column_config.TextColumn(
                help="e.g. STANDARD. Comma-separated. Blank = all services. "
                     "UPDE oversize defaults to STANDARD only (excludes Express Saver).",
                width="small"),
            'Size limit (m)': st.column_config.NumberColumn(
                format="%.2f", help="Stamped on the normal (cheap) rows."),
            'Surcharge €/parcel': st.column_config.NumberColumn(format="%.2f"),
        },
        key='exceptions_editor',
    )
    st.session_state.exceptions_df = edited

    st.markdown('---')
    st.markdown("**Overflow buckets** — catch orders heavier or with more "
                "parcels than the grid. Off by default. The overflow rate (€ per "
                "parcel) is your contract heavy/per-kg rate — it is **not** guessed.")
    if 'overflow_df' not in st.session_state:
        st.session_state.overflow_df = DEFAULT_OVERFLOW.copy()
    ov_edit = st.data_editor(
        st.session_state.overflow_df, num_rows="dynamic",
        use_container_width=True, hide_index=True,
        column_config={
            'Enabled': st.column_config.CheckboxColumn(width="small"),
            'Carrier': st.column_config.SelectboxColumn(
                options=['(all)'] + list(pl.CARRIER_DEFAULTS), width="small"),
            'Country (blank=all)': st.column_config.TextColumn(width="small"),
            'Overflow rate €/parcel': st.column_config.NumberColumn(format="%.2f"),
            'Surcharge €/parcel': st.column_config.NumberColumn(format="%.2f"),
        }, key='overflow_editor')
    st.session_state.overflow_df = ov_edit

    st.markdown("**Postcode catch-all** — for zoned carriers, add a blank-postcode "
                "fallback at the worst zone's rate so an unlisted prefix still matches. "
                "Off by default.")
    if 'postcode_df' not in st.session_state:
        st.session_state.postcode_df = DEFAULT_POSTCODE.copy()
    pc_edit = st.data_editor(
        st.session_state.postcode_df, num_rows="dynamic",
        use_container_width=True, hide_index=True,
        column_config={
            'Enabled': st.column_config.CheckboxColumn(width="small"),
            'Carrier': st.column_config.SelectboxColumn(
                options=['(all)'] + list(pl.CARRIER_DEFAULTS), width="small"),
            'Country (blank=all)': st.column_config.TextColumn(width="small"),
            'Surcharge €': st.column_config.NumberColumn(format="%.2f"),
        }, key='postcode_editor')
    st.session_state.postcode_df = pc_edit

    st.markdown('---')
    st.markdown("**Heavy / oversized parcel** — per-parcel **weight** surcharge. "
                "Any row whose per-box cap (EACH_WEIGHT) reaches the threshold is "
                "surcharged €/parcel × parcels, flagged (AWKWARD) and shown in amber. "
                "It does **not** overwrite the weight grid. Default: DHL parcels "
                "≥ 20 kg → €4.89/parcel. Note: a heavy parcel that is *also* over a "
                "size limit is charged the size surcharge only (rare).")
    if 'heavy_df' not in st.session_state:
        st.session_state.heavy_df = DEFAULT_HEAVY.copy()
    hv_edit = st.data_editor(
        st.session_state.heavy_df, num_rows="dynamic",
        use_container_width=True, hide_index=True,
        column_config={
            'Enabled': st.column_config.CheckboxColumn(width="small"),
            'Carrier': st.column_config.SelectboxColumn(
                options=['(all)'] + list(pl.CARRIER_DEFAULTS), width="small"),
            'Country (blank=all)': st.column_config.TextColumn(width="small"),
            'Weight threshold kg': st.column_config.NumberColumn(
                format="%.1f", help="Surcharge rows whose per-box cap can hold a "
                                    "parcel at/over this weight."),
            'Surcharge €/parcel': st.column_config.NumberColumn(format="%.2f"),
        }, key='heavy_editor')
    st.session_state.heavy_df = hv_edit

st.divider()

# ── Run ────────────────────────────────────────────────────────────────────────
if 'results' not in st.session_state:
    st.session_state.results = {}
if 'results_express' not in st.session_state:
    st.session_state.results_express = {}

if (run_btn or exp_btn) and uploaded and selected:
    express_mode = bool(exp_btn)
    _store = 'results_express' if express_mode else 'results'
    st.session_state[_store] = {}
    input_path = st.session_state['input_path']
    errors = {}
    rules    = exception_rules_from_editor(st.session_state.get('exceptions_df', DEFAULT_EXCEPTIONS))
    hv_rules = heavy_rules_from_editor(st.session_state.get('heavy_df', DEFAULT_HEAVY))
    rules    = rules + hv_rules   # heavy (threshold) rules evaluated after oversize (stamp)
    ov_rules = overflow_rules_from_editor(st.session_state.get('overflow_df', DEFAULT_OVERFLOW))
    pc_rules = postcode_rules_from_editor(st.session_state.get('postcode_df', DEFAULT_POSTCODE))
    pallets  = st.session_state.get('pallets')   # None if no pallet card uploaded

    # Pallet surcharge config: globals (fuel/mobility/admin/factor) + UK-only toll.
    pal_defaults = deepcopy(pl.PALLET_DEFAULTS)
    pal_overrides = deepcopy(pl.PALLET_COUNTRY_OVERRIDES)
    if pallet_vals:
        pal_defaults['DHL-FENDER'].update({
            'fuel_pct':           pallet_vals.get('fuel',     pal_defaults['DHL-FENDER']['fuel_pct']),
            'mobility_pct':       pallet_vals.get('mobility', pal_defaults['DHL-FENDER']['mobility_pct']),
            'factor':             pallet_vals.get('factor',   pal_defaults['DHL-FENDER']['factor']),
            'admin_per_shipment': pallet_vals.get('admin',    pal_defaults['DHL-FENDER']['admin_per_shipment']),
        })
        # Toll is GB-only in the contract; apply the sidebar value to GB.
        pal_overrides.setdefault('GB', {})
        pal_overrides['GB']['toll_pct'] = pallet_vals.get('toll',
                                                          pal_overrides['GB'].get('toll_pct', 0.0043))
        pal_overrides['GB']['admin_per_shipment'] = pallet_vals.get(
            'admin', pal_defaults['DHL-FENDER']['admin_per_shipment'])

    _maut_dhl_ref = pl.CARRIER_DEFAULTS['DHL-ROS']['maut_pct']
    _maut_dpd_ref = pl.CARRIER_DEFAULTS['DPD']['maut_pct']
    st.session_state['variables_layout_rows'] = variables_layout(
        fuel_vals, _maut_dhl_ref, _maut_dpd_ref, pallet_vals)
    # Stash pallet config so the combined export can write matching formulas.
    st.session_state['pallet_maut_table'] = pallet_maut_table
    st.session_state['pallet_defaults_used'] = pal_defaults
    st.session_state['carrier_defaults_used'] = carrier_defaults(
        fuel_vals, _maut_dhl_ref, _maut_dpd_ref)
    progress = st.progress(0, text="Starting…")

    for idx, country in enumerate(selected):
        progress.progress(idx / len(selected),
                          text=f"Processing {country}…  ({idx+1}/{len(selected)})")
        try:
            cfg = country_cfg_with_overrides(country)
            if express_mode:
                cfg['carriers'] = [c for c in cfg['carriers']
                                   if c in EXPRESS_CARRIERS]
            pallet_zones = (pp.country_pallet_data(pallets, country)
                            if pallets and not express_mode else None)
            # Postcode list for parcel expansion: sourced from the pallet file's
            # per-country zip prefixes regardless of the pallet-merge / express
            # toggles (express rows are parcel rows and still need postcodes).
            parcel_postcodes = (list(pp.country_pallet_data(pallets, country).keys())
                                if pallets else None)

            if is_master:
                parsed = mp.country_rate_data(master, country)
                maut   = mp.country_maut(master, country)
                cd = carrier_defaults(fuel_vals, maut['DHL-ROS'], maut['DPD'])
                vl = variables_layout(fuel_vals, maut['DHL-ROS'], maut['DPD'], pallet_vals)
            else:
                parsed = st.session_state.get('parsed_per_country', {})
                cd = carrier_defaults(fuel_vals, maut_dhl, maut_dpd)
                vl = variables_layout(fuel_vals, maut_dhl, maut_dpd, pallet_vals)

            missing = [cid for cid in cfg['carriers'] if cid not in parsed]
            for cid in missing:
                errors.setdefault(country, []).append(
                    f"⚠️ {cid}: no rate data found for {country} — carrier skipped")

            with tempfile.TemporaryDirectory() as tmp:
                result = pl.run_pipeline_from_parsed(
                    parsed, country, tmp, cfg, cd, vl,
                    exceptions=rules, overflow_rules=ov_rules, postcode_rules=pc_rules,
                    pallet_zones=pallet_zones, pallet_defaults=pal_defaults,
                    pallet_overrides=pal_overrides, pallet_maut=pallet_maut_table,
                    pallet_max_band_kg=(pallet_max_band_kg or None),
                    express_only=express_mode,
                    parcel_postcodes=parcel_postcodes)
                result = persist(result)
            if express_mode:
                result = express_rename(result)

            if result.get('pallet_warning'):
                errors.setdefault(country, []).append(result['pallet_warning'])
            if result.get('postcode_warning'):
                errors.setdefault(country, []).append(result['postcode_warning'])
            st.session_state[_store][country] = result
        except Exception as e:
            errors.setdefault(country, []).append(str(e))

    progress.progress(1.0, text="Done.")
    for country, msgs in errors.items():
        for msg in (msgs if isinstance(msgs, list) else [msgs]):
            if msg.startswith("⚠️"):
                st.warning(f"**{country}**: {msg[2:].strip()}")
            else:
                st.error(f"**{country}**: {msg}")

elif (run_btn or exp_btn) and not selected:
    st.warning("Please select at least one country.")

# ── Results ──────────────────────────────────────────────────────────────────
def render_results(results, heading, kp, fname_prefix, *, caption=None):
    """Render the summary + per-country + bulk download blocks for one result set.
    `kp` (key-prefix) keeps Streamlit widget keys unique across the two sections;
    `fname_prefix` marks express download filenames."""
    st.markdown(heading)
    if caption:
        st.caption(caption)
    summary = [{'Country': c, 'Extended': f"{r['rows_extended']:,}",
                'Optimized': f"{r['rows_optimized']:,}", 'Minimal': f"{r['rows_minimal']:,}"}
               for c, r in results.items()]
    st.dataframe(pd.DataFrame(summary).set_index('Country'), use_container_width=True)
    st.caption("**Extended** = all combinations · **Optimized** = per-carrier dominance "
               "removed · **Minimal** = cross-carrier dominance removed (use this one)")

    st.markdown("#### Download individual countries")
    for country, r in results.items():
        c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
        c1.markdown(f"**{country}**")
        for col, key, label in [(c2, 'extended', '📥 Extended'),
                                (c3, 'optimized', '📥 Optimized'),
                                (c4, 'minimal', '📥 Minimal')]:
            col.download_button(label, data=file_bytes(r[key]),
                                file_name=Path(r[key]).name,
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                key=f'dl_{kp}_{country}_{key}')

    st.markdown("#### Download everything")
    dc1, dc2 = st.columns(2)
    with dc1:
        st.download_button("📦 Download all countries as ZIP", data=make_zip(results),
                           file_name=f"{fname_prefix}rate_matrices.zip", mime="application/zip",
                           type="primary", key=f'dl_{kp}_zip')
    with dc2:
        _vl_rows = st.session_state.get('variables_layout_rows', pl.VARIABLES_LAYOUT)
        _xlsx_mime = ('application/vnd.openxmlformats-officedocument.'
                      'spreadsheetml.sheet')
        st.caption("🧩 **Combined** — every selected country merged into one sheet, "
                   "sorted by country then price. Numeric values so per-country "
                   "surcharges (e.g. MAUT) stay correct.")
        for _stage, _label in [('extended',  '🧩 Combined extended'),
                               ('optimized', '🧩 Combined optimized'),
                               ('minimal',   '🧩 Combined minimal')]:
            _combined = make_combined(
                results, _vl_rows, stage=_stage,
                pallet_maut=st.session_state.get('pallet_maut_table'),
                pallet_defaults=st.session_state.get('pallet_defaults_used'),
                carrier_defaults=st.session_state.get('carrier_defaults_used'))
            if _combined is not None:
                st.download_button(
                    _label, data=_combined,
                    file_name=f"{fname_prefix}Combined_Matrix_{_stage}.xlsx", mime=_xlsx_mime,
                    key=f'dl_{kp}_combined_{_stage}',
                    type=('primary' if _stage == 'minimal' else 'secondary'))
            else:
                st.caption(f"Combined {_stage} unavailable — re-run to regenerate.")


if st.session_state.get('results'):
    render_results(st.session_state.results, "## Results", "std", "")

if st.session_state.get('results_express'):
    st.divider()
    render_results(
        st.session_state.results_express, "## ⚡ Express-only results", "exp", "EXPRESS_",
        caption="EXPRESS SAVER options only — UPS DE (7R9W62 + EXPRESS SAVER), "
                "UPS NL EXPRESS SAVER, UPS GB EXPRESS SAVER. No STANDARD / pallets.")



