"""
Reverse Geocoding Web App (Streamlit) — v3
==========================================
Reads an Excel file with latitude / longitude columns and produces a new Excel
file with the road-level physical address, Plus Code, and Google Maps link for
every point.

Providers:
  * Nominatim (free, no key) — unreliable on Streamlit Cloud (shared IP → 429)
  * LocationIQ (free key)   — RECOMMENDED for cloud; 5,000 req/day, 2 req/s
  * Google Maps (paid key)  — most accurate, especially in Pakistan

Run locally:
    pip install -r requirements.txt
    python -m streamlit run app.py
"""

import io
import os
import time
from collections import deque

# --------------------------------------------------------------------------- #
#  Dependency check — fail with a clear message instead of a cryptic traceback
# --------------------------------------------------------------------------- #
_MISSING = []

try:
    import pandas as pd
except ImportError:
    _MISSING.append("pandas")

try:
    import requests
except ImportError:
    _MISSING.append("requests")

try:
    import streamlit as st
except ImportError:
    raise SystemExit(
        "Streamlit is not installed. Run: pip install streamlit"
    )

_HAS_GEOPY = True
try:
    from geopy.exc import (GeocoderServiceError, GeocoderTimedOut,
                           GeocoderUnavailable)
    from geopy.geocoders import Nominatim
except ImportError:
    _HAS_GEOPY = False

if _MISSING or not _HAS_GEOPY:
    st.set_page_config(page_title="Missing dependencies", page_icon="⚠️")
    st.error(
        "### ⚠️ Missing Python packages\n\n"
        "This app needs: `pandas`, `requests`, `geopy`, `openpyxl`.\n\n"
        "**Fix:** add a file named `requirements.txt` to the root of your repo "
        "with this content:\n\n"
        "```txt\n"
        "streamlit>=1.32.0\n"
        "pandas>=2.0.0\n"
        "openpyxl>=3.1.0\n"
        "geopy>=2.4.0\n"
        "requests>=2.31.0\n"
        "openlocationcode>=1.0.1\n"
        "```\n\n"
        "Then click **Manage app → Reboot** in Streamlit Cloud.\n\n"
        f"**Detected missing:** {', '.join(_MISSING) or 'geopy'}"
    )
    st.stop()

# Optional Plus Code support
try:
    from openlocationcode import openlocationcode as olc
    HAS_OLC = True
except Exception:
    HAS_OLC = False


# --------------------------------------------------------------------------- #
#  Page config
# --------------------------------------------------------------------------- #
st.set_page_config(
    page_title="Lat/Long → Physical Address Extractor",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# --------------------------------------------------------------------------- #
#  Watermark — fixed bottom-right corner
# --------------------------------------------------------------------------- #
st.markdown(
    """
    <style>
        #dev-watermark {
            position: fixed; right: 16px; bottom: 14px; z-index: 2147483647;
            padding: 8px 14px; border-radius: 10px;
            background: linear-gradient(135deg, rgba(17,24,39,.92), rgba(55,65,81,.92));
            color: #fff !important;
            font-family: "Segoe UI", Roboto, Arial, sans-serif;
            font-size: 12.5px; font-weight: 600; letter-spacing: .2px;
            line-height: 1.4; text-align: right;
            box-shadow: 0 6px 18px rgba(0,0,0,.28);
            border: 1px solid rgba(255,255,255,.16);
            pointer-events: none; user-select: none;
        }
        #dev-watermark .nm { color: #4ade80; font-weight: 700; }
        #dev-watermark .ph { color: #e5e7eb; font-weight: 500; font-size: 11.5px; }
        @media (max-width: 640px) {
            #dev-watermark { font-size: 10.5px; padding: 6px 10px;
                             right: 8px; bottom: 8px; }
        }
        .block-container { padding-bottom: 4rem; }
    </style>
    <div id="dev-watermark">
        Developed By <span class="nm">Engr Muhammad Asghar</span><br>
        <span class="ph">+92-345-8383838</span>
    </div>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------- #
#  Helpers
# --------------------------------------------------------------------------- #
def to_float(value):
    """Convert a cell value to float, tolerating comma decimal separators."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        raise ValueError("empty")
    s = str(value).strip().replace(",", ".")
    if s == "" or s.lower() in ("nan", "none", "null"):
        raise ValueError("empty")
    return float(s)


@st.cache_data(show_spinner=False)
def get_sheet_names(data: bytes):
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


@st.cache_data(show_spinner=False)
def get_sheet(data: bytes, sheet: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(data), sheet_name=sheet)


def guess_column(columns, keys):
    """Best-effort auto-detection of a latitude / longitude column."""
    for c in columns:
        cl = str(c).lower()
        if any(k in cl for k in keys):
            return c
    return None


def make_plus_code(lat, lon):
    """10-digit Plus Code (~14 m precision). No API needed."""
    if not HAS_OLC:
        return ""
    try:
        return olc.encode(lat, lon, 10)
    except Exception:
        return ""


def make_gmaps_link(lat, lon):
    """Universal Google Maps link — always opens the exact pin."""
    return f"https://www.google.com/maps?q={lat:.6f},{lon:.6f}"


# Nominatim address-component priority — most specific first.
_ADDR_ORDER = [
    "house_number", "road", "pedestrian", "footway", "path", "cycleway",
    "neighbourhood", "suburb", "hamlet", "village", "town",
    "city_district", "city", "municipality", "county",
    "state_district", "state", "postcode", "country",
]


def build_road_level_address(raw_address: dict) -> str:
    """
    Rebuild a road-level address from Nominatim / LocationIQ components.
    Returns a string ordered from most-specific to least-specific.
    """
    if not raw_address:
        return ""
    parts, seen = [], set()
    for key in _ADDR_ORDER:
        v = raw_address.get(key)
        if v and str(v) not in seen:
            parts.append(str(v))
            seen.add(str(v))
    return ", ".join(parts)


# --------------------------------------------------------------------------- #
#  Geocoding back-ends
# --------------------------------------------------------------------------- #
def lookup_nominatim(geolocator, lat, lon):
    """Return (road_level_addr, full_addr, status)."""
    for attempt in range(3):
        try:
            loc = geolocator.reverse((lat, lon), exactly_one=True,
                                     language="en", addressdetails=True)
            if loc is None:
                return "", "", "No result"
            raw = loc.raw.get("address", {}) if hasattr(loc, "raw") else {}
            detailed = build_road_level_address(raw)
            return detailed or loc.address, loc.address, "OK"
        except (GeocoderTimedOut, GeocoderUnavailable):
            time.sleep(2)
        except GeocoderServiceError as e:
            msg = str(e)
            if "429" in msg:
                wait = min(60, (2 ** attempt) * 5)
                time.sleep(wait)
                continue
            return "", "", f"Service error: {e}"
        except Exception as e:
            return "", "", f"Error: {e}"
    return "", "", "Rate limited (429)"


def lookup_locationiq(api_key, lat, lon):
    """LocationIQ reverse geocoding — Nominatim-compatible, generous free tier.

    Free plan: 5,000 requests/day, 2 requests/second.
    Sign up at https://locationiq.com/register
    """
    url = "https://us1.locationiq.com/v1/reverse"
    for attempt in range(3):
        try:
            r = requests.get(
                url,
                timeout=15,
                params={
                    "key": api_key,
                    "lat": f"{lat:.7f}",
                    "lon": f"{lon:.7f}",
                    "format": "json",
                    "addressdetails": 1,
                    "accept-language": "en",
                },
            )
            if r.status_code == 429:
                wait = min(60, (2 ** attempt) * 5)
                time.sleep(wait)
                continue
            if r.status_code != 200:
                return "", "", f"HTTP {r.status_code}"
            data = r.json()
            raw = data.get("address", {})
            detailed = build_road_level_address(raw)
            full = data.get("display_name", "")
            return detailed or full, full, "OK"
        except Exception as e:
            return "", "", f"Error: {e}"
    return "", "", "Rate limited (429)"


def lookup_google(api_key, lat, lon):
    """Return (road_level_addr, full_addr, status)."""
    url = "https://maps.googleapis.com/maps/api/geocode/json"
    for attempt in range(3):
        try:
            r = requests.get(
                url, timeout=15,
                params={"latlng": f"{lat},{lon}", "key": api_key},
            )
            data = r.json()
            status = data.get("status", "UNKNOWN")
            if status == "OK":
                result = data["results"][0]
                full = result.get("formatted_address", "")
                parts, seen = [], set()
                for comp in result.get("address_components", []):
                    name = comp.get("long_name")
                    if name and name not in seen:
                        parts.append(name)
                        seen.add(name)
                return ", ".join(parts) or full, full, "OK"
            if status == "ZERO_RESULTS":
                return "", "", "No result"
            if status in ("OVER_QUERY_LIMIT", "OVER_DAILY_LIMIT"):
                wait = min(60, (2 ** attempt) * 5)
                time.sleep(wait)
                continue
            return "", "", status
        except Exception as e:
            return "", "", f"Error: {e}"
    return "", "", "Rate limited"


# --------------------------------------------------------------------------- #
#  Session state
# --------------------------------------------------------------------------- #
defaults = {
    "file_id": None,
    "file_bytes": None,
    "cache": {},
    "result_bytes": None,
    "result_name": "geocoded.xlsx",
    "summary": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# --------------------------------------------------------------------------- #
#  Header
# --------------------------------------------------------------------------- #
st.title("🌍 Lat/Long → Physical Address Extractor")
st.caption(
    "Upload an Excel file with latitude / longitude columns. "
    "Every row gets a **road-level address**, a **Plus Code**, and a "
    "**Google Maps link** — so anyone can actually reach the location."
)

# --------------------------------------------------------------------------- #
#  1 · Input
# --------------------------------------------------------------------------- #
st.markdown("### 1 · Input Excel file")
uploaded = st.file_uploader(
    "Choose an Excel file",
    type=["xlsx", "xls", "xlsm"],
    label_visibility="collapsed",
)

if uploaded is None:
    st.info("⬆️  Upload an Excel file containing latitude / longitude columns.")
    st.stop()

file_id = f"{uploaded.name}:{uploaded.size}"
if st.session_state.file_id != file_id:
    st.session_state.file_id = file_id
    st.session_state.file_bytes = uploaded.getvalue()
    st.session_state.result_bytes = None
    st.session_state.summary = None
    st.session_state.cache = {}

data_bytes = st.session_state.file_bytes

try:
    sheet_names = get_sheet_names(data_bytes)
except Exception as e:
    st.error(f"Could not open the file: {e}")
    st.stop()

c1, c2, c3 = st.columns([1, 1, 1])
sheet = c1.selectbox("Sheet", sheet_names)

try:
    df = get_sheet(data_bytes, sheet)
except Exception as e:
    st.error(f"Could not read the sheet: {e}")
    st.stop()

cols = list(df.columns)
if len(cols) < 2:
    st.error("The selected sheet needs at least two columns (latitude & longitude).")
    st.stop()

lat_guess = guess_column(cols, ["lat", "y_coord", "ycoord", "breite"])
lon_guess = guess_column(cols, ["lon", "lng", "long", "x_coord", "xcoord", "länge"])
lat_idx = cols.index(lat_guess) if lat_guess in cols else 0
lon_idx = cols.index(lon_guess) if lon_guess in cols else min(1, len(cols) - 1)

lat_col = c2.selectbox("Latitude column", cols, index=lat_idx, key=f"lat_{sheet}")
lon_col = c3.selectbox("Longitude column", cols, index=lon_idx, key=f"lon_{sheet}")

with st.expander(f"👀 Preview — '{sheet}'  ({len(df):,} rows × {len(cols)} columns)"):
    st.dataframe(df.head(15), use_container_width=True)

st.divider()

# --------------------------------------------------------------------------- #
#  2 · Service
# --------------------------------------------------------------------------- #
st.markdown("### 2 · Geocoding service")

s1, s2 = st.columns([2, 1])

provider = s1.radio(
    "Provider",
    ["LocationIQ (free key) — recommended",
     "Nominatim (free, no key)",
     "Google Maps (API key)"],
    horizontal=False,
)

api_key = ""
if provider.startswith("LocationIQ"):
    api_key = s2.text_input(
        "LocationIQ API key",
        type="password",
        help="Free key from https://locationiq.com/register — "
             "5,000 requests/day, 2 req/sec. Solves 429 errors on Streamlit Cloud.",
    )
elif provider.startswith("Google"):
    api_key = s2.text_input(
        "Google API key",
        type="password",
        help="Enable the Geocoding API in Google Cloud Console.",
    )

d1, d2, d3 = st.columns([1, 1, 2])
default_delay = 0.55 if provider.startswith("LocationIQ") else 1.1
delay = d1.number_input(
    "Delay between requests (s)",
    min_value=0.0, max_value=10.0, value=float(default_delay),
    step=0.05, format="%.2f",
    help="Nominatim: 1.1s (1 req/s). LocationIQ: 0.55s (2 req/s). Google: 0.0s.",
)
limit = d2.number_input(
    "Limit rows (0 = all)",
    min_value=0, max_value=1_000_000, value=0, step=10,
    help="Process only the first N rows — handy for a quick test run.",
)

if not HAS_OLC:
    d3.markdown(
        "<div style='padding-top:28px;color:#b45309;font-size:13px;'>"
        "⚠️  <code>openlocationcode</code> not installed — Plus Codes will be "
        "blank. Add it to <code>requirements.txt</code>.</div>",
        unsafe_allow_html=True,
    )

if provider.startswith("Nominatim"):
    st.warning(
        "⚠️ **Nominatim on Streamlit Cloud is unreliable.** Because Streamlit "
        "Cloud apps share public IP addresses, you will frequently get "
        "`429 Too Many Requests`. For a dependable cloud deployment, use "
        "**LocationIQ (free key)** — 5,000 requests/day, 2 req/sec."
    )
elif provider.startswith("LocationIQ"):
    st.info(
        "ℹ️ **LocationIQ free tier:** 5,000 requests/day, 2 req/sec. "
        "Get a key at https://locationiq.com/register"
    )

if st.session_state.cache:
    cc1, cc2 = st.columns([4, 1])
    cc1.caption(f"🧠 Cached look-ups: {len(st.session_state.cache):,}")
    if cc2.button("Clear cache", use_container_width=True):
        st.session_state.cache = {}
        st.rerun()

st.divider()

# --------------------------------------------------------------------------- #
#  3 · Run
# --------------------------------------------------------------------------- #
st.markdown("### 3 · Run")
start = st.button("▶  Start geocoding", type="primary")

if start:
    # ---- Validation
    if provider.startswith("Google") and not api_key.strip():
        st.warning("Please enter your Google API key.")
        st.stop()
    if provider.startswith("LocationIQ") and not api_key.strip():
        st.warning(
            "Please enter your LocationIQ API key. "
            "Get a free one at https://locationiq.com/register"
        )
        st.stop()
    if lat_col == lon_col:
        st.warning("Latitude and longitude columns must be different.")
        st.stop()

    work_df = df if limit == 0 else df.head(int(limit))
    total = len(work_df)
    if total == 0:
        st.warning("There is nothing to process.")
        st.stop()

    progress = st.progress(0.0)
    status_line = st.empty()
    log_box = st.empty()
    log_lines = deque(maxlen=14)

    geolocator = None
    if provider.startswith("Nominatim"):
        geolocator = Nominatim(
            user_agent="excel_reverse_geocoder_streamlit/3.0",
            timeout=15,
        )

    addresses, details, plus_codes, gm_links, statuses = [], [], [], [], []
    cache = st.session_state.cache
    ok = fail = 0
    start_t = time.time()

    for pos, (_, row) in enumerate(work_df.iterrows(), start=1):
        # ---- Parse coordinates
        try:
            lat = to_float(row[lat_col])
            lon = to_float(row[lon_col])
            if not (-90 <= lat <= 90 and -180 <= lon <= 180):
                raise ValueError("out of range")
        except Exception:
            addresses.append("")
            details.append("")
            plus_codes.append("")
            gm_links.append("")
            statuses.append("Invalid coordinates")
            fail += 1
            log_lines.append(f"Row {pos}/{total}: ⚠️ invalid coordinates")
            progress.progress(pos / total)
            status_line.markdown(
                f"**{pos} / {total}** · ✅ {ok} · ❌ {fail}"
            )
            log_box.code("\n".join(log_lines), language=None)
            continue

        # ---- Cached?
        key = (round(lat, 6), round(lon, 6))
        if key in cache:
            road_addr, full_addr, status = cache[key]
            status = "OK (cached)"
        else:
            if provider.startswith("LocationIQ"):
                road_addr, full_addr, status = lookup_locationiq(
                    api_key.strip(), lat, lon
                )
            elif provider.startswith("Nominatim"):
                road_addr, full_addr, status = lookup_nominatim(
                    geolocator, lat, lon
                )
            else:
                road_addr, full_addr, status = lookup_google(
                    api_key.strip(), lat, lon
                )

            # Only cache successful lookups — failures get retried next run.
            if status == "OK":
                cache[key] = (road_addr, full_addr, status)

            # Respect the provider rate limit.
            if delay > 0:
                time.sleep(delay)

        addresses.append(road_addr)
        details.append(full_addr)
        plus_codes.append(make_plus_code(lat, lon))
        gm_links.append(make_gmaps_link(lat, lon))
        statuses.append(status)

        if road_addr:
            ok += 1
            log_lines.append(f"Row {pos}/{total}: ✅ {road_addr[:95]}")
        else:
            fail += 1
            log_lines.append(f"Row {pos}/{total}: ❌ FAILED ({status})")

        # ---- Progress + ETA
        elapsed = time.time() - start_t
        rate = pos / elapsed if elapsed > 0 else 0
        eta = ""
        if rate > 0 and pos < total:
            rem = int((total - pos) / rate)
            eta = f" · ETA {rem // 60}m {rem % 60}s"

        progress.progress(pos / total)
        status_line.markdown(
            f"**{pos} / {total}** · ✅ {ok} found · ❌ {fail} failed{eta}"
        )
        log_box.code("\n".join(log_lines), language=None)

    # ---- Build output workbook
    out_df = df.copy()
    while len(addresses) < len(out_df):
        addresses.append("")
        details.append("")
        plus_codes.append("")
        gm_links.append("")
        statuses.append("Not processed")

    lat_out, lon_out = [], []
    for _, r in out_df.iterrows():
        try:
            lat_out.append(to_float(r[lat_col]))
            lon_out.append(to_float(r[lon_col]))
        except Exception:
            lat_out.append(None)
            lon_out.append(None)

    out_df["Physical_Address"]   = addresses     # road-level address
    out_df["Full_Admin_Address"] = details       # admin fallback / full string
    out_df["Plus_Code"]          = plus_codes    # e.g. 7JQ6+XX Quetta
    out_df["Google_Maps_Link"]   = gm_links      # always opens the exact pin
    out_df["Latitude"]           = lat_out
    out_df["Longitude"]          = lon_out
    out_df["Geocode_Status"]     = statuses

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False)

    base = os.path.splitext(uploaded.name)[0]
    st.session_state.result_bytes = buf.getvalue()
    st.session_state.result_name = f"{base}_with_addresses.xlsx"
    st.session_state.summary = (ok, fail, total)

    progress.progress(1.0)
    status_line.markdown(
        f"### ✅ Finished — {ok} addresses found, {fail} failed (of {total})"
    )

# --------------------------------------------------------------------------- #
#  4 · Download
# --------------------------------------------------------------------------- #
if st.session_state.result_bytes:
    ok, fail, total = st.session_state.summary
    m1, m2, m3 = st.columns(3)
    m1.metric("✅ Addresses found", f"{ok:,}")
    m2.metric("❌ Failed / empty", f"{fail:,}")
    m3.metric("📄 Rows processed", f"{total:,}")

    st.download_button(
        "⬇️  Download result Excel file",
        data=st.session_state.result_bytes,
        file_name=st.session_state.result_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

st.divider()
st.caption(
    "ℹ️ Data © OpenStreetMap contributors (via Nominatim / LocationIQ) or "
    "Google Maps Platform. In areas with sparse map data, the "
    "**Google Maps link** and **Plus Code** are the most reliable way to "
    "reach a location."
)