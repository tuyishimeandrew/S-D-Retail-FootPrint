"""
S&D Retail FootPrint — MTN-branded Streamlit viewer

Reads:
    Retail_Footprint_Prepared.xlsx        (sheets: Shops, Boundaries, ShopAreaAssignments)
    Retail_Footprint_Boundaries.parquet   (one polygon per boundary key)
"""
import io
import json
import re
from pathlib import Path

import branca.colormap as cm
import folium
import numpy as np
import pandas as pd
import pyarrow.parquet as pq
import streamlit as st
from shapely.geometry import mapping, shape
from shapely import contains_xy
from shapely.ops import unary_union
from streamlit_folium import st_folium

# =========================================================
# MTN BRAND TOKENS
# =========================================================
MTN_YELLOW = "#FFCC08"
MTN_BLACK = "#000000"
MTN_GREY = "#F4F4F2"
MTN_MID_GREY = "#6B6B6B"
MTN_BORDER = "#E3E3E0"

st.set_page_config(
    layout="wide",
    page_title="S&D Retail FootPrint | MTN",
    page_icon="🟡",
)

# =========================================================
# GLOBAL STYLING
# =========================================================
st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800;900&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Montserrat', -apple-system, BlinkMacSystemFont, sans-serif;
    }}
    .stApp {{ background: {MTN_GREY}; }}

    .block-container {{
        padding-top: 0.5rem !important;
        padding-bottom: 2rem !important;
        max-width: 100% !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
    }}

    header[data-testid="stHeader"] {{
        display: none !important;
        height: 0 !important;
    }}

    /* ---------- STICKY ---------- */
    .st-key-sticky_top {{
        position: -webkit-sticky !important;
        position: sticky !important;
        top: 0 !important;
        z-index: 2147483647 !important;
        background: {MTN_GREY} !important;
        padding: 14px 1.5rem 16px 1.5rem !important;
        margin: 0 -1.5rem 1rem -1.5rem !important;
        border-bottom: 4px solid {MTN_YELLOW} !important;
        box-shadow: 0 8px 16px -6px rgba(0,0,0,0.15) !important;
    }}

    [data-testid="stAppViewContainer"],
    [data-testid="stAppViewContainer"] > .main,
    [data-testid="stAppViewContainer"] > [data-testid="stMain"],
    [data-testid="stMain"],
    [data-testid="stMainBlockContainer"],
    section.main,
    .block-container,
    .block-container > div,
    div[data-testid="stVerticalBlock"] {{
        overflow: visible !important;
    }}
    [data-testid="stAppViewContainer"] {{
        overflow-y: auto !important;
        overflow-x: hidden !important;
        height: 100vh !important;
    }}

    /* ---------- Masthead ---------- */
    .mtn-header {{
        display: flex; align-items: center; gap: 20px;
        background: {MTN_BLACK};
        padding: 20px 26px;
        border-bottom: 6px solid {MTN_YELLOW};
        border-radius: 10px;
        margin-bottom: 14px;
    }}
    .mtn-logo-chip {{
        background: {MTN_YELLOW};
        border-radius: 8px;
        height: 64px; min-width: 84px;
        display: flex; align-items: center; justify-content: center;
    }}
    .mtn-logo-text {{
        font-weight: 900; font-size: 1.7rem;
        color: {MTN_BLACK}; letter-spacing: -1.5px; line-height: 1;
    }}
    .mtn-title {{
        font-size: 2.1rem; font-weight: 900;
        color: #FFFFFF; letter-spacing: -0.5px;
        margin: 0; line-height: 1.1;
    }}
    .mtn-title span {{ color: {MTN_YELLOW}; }}
    .mtn-subtitle {{
        font-size: 0.72rem; font-weight: 600;
        color: {MTN_YELLOW}; letter-spacing: 2.2px;
        text-transform: uppercase; margin-top: 5px;
    }}

    /* ---------- Sidebar ---------- */
    section[data-testid="stSidebar"] {{
        background: {MTN_BLACK};
        border-right: 4px solid {MTN_YELLOW};
    }}
    section[data-testid="stSidebar"] > div {{ background: {MTN_BLACK}; }}
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {{
        color: {MTN_YELLOW} !important;
        font-weight: 800 !important;
        letter-spacing: 0.8px;
        text-transform: uppercase;
        font-size: 0.82rem !important;
    }}
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] .stMarkdown {{
        color: #F2F2F2 !important;
        font-weight: 500;
    }}
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div {{
        background-color: #141414 !important;
        border: 1.5px solid {MTN_YELLOW} !important;
        color: #FFFFFF !important;
    }}
    section[data-testid="stSidebar"] div[data-baseweb="select"] * {{ color: #FFFFFF !important; }}
    section[data-testid="stSidebar"] div[data-baseweb="select"] svg {{ fill: {MTN_YELLOW} !important; }}
    span[data-baseweb="tag"] {{
        background-color: {MTN_YELLOW} !important;
        color: {MTN_BLACK} !important;
        font-weight: 800 !important;
    }}
    span[data-baseweb="tag"] span {{ color: {MTN_BLACK} !important; }}
    span[data-baseweb="tag"] svg {{ fill: {MTN_BLACK} !important; }}

    /* ---------- Custom KPI cards ---------- */
    .mtn-kpi {{
        background: #FFFFFF;
        border: 1px solid {MTN_BORDER};
        border-left: 5px solid {MTN_YELLOW};
        border-radius: 8px;
        padding: 14px 16px 12px 16px;
        height: 100%;
        min-height: 92px;
        box-sizing: border-box;
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
    }}
    .mtn-kpi-value {{
        color: {MTN_BLACK};
        font-weight: 900;
        font-size: 1.5rem;
        line-height: 1.1;
        margin: 0;
        padding: 0;
        word-break: break-word;
    }}
    .mtn-kpi-label {{
        color: {MTN_MID_GREY};
        font-weight: 700;
        font-size: 0.63rem;
        letter-spacing: 0.3px;
        text-transform: uppercase;
        line-height: 1.35;
        margin-top: 8px;
        padding: 0;
        white-space: normal;
        word-wrap: break-word;
        overflow-wrap: anywhere;
        word-break: break-word;
    }}

    /* ---------- Section labels & buttons ---------- */
    .mtn-section-label {{
        display: inline-block;
        background: {MTN_YELLOW};
        color: {MTN_BLACK};
        font-weight: 800;
        font-size: 0.75rem;
        letter-spacing: 1px;
        text-transform: uppercase;
        padding: 5px 12px;
        border-radius: 4px;
        margin: 8px 0 10px 0;
    }}
    .stDownloadButton > button {{
        background: {MTN_YELLOW};
        color: {MTN_BLACK};
        font-weight: 800;
        border: 0;
        border-radius: 6px;
    }}
    .stDownloadButton > button:hover {{
        background: #E6B800;
        color: {MTN_BLACK};
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# CONFIG
# =========================================================
PREPARED_NAME = "Retail_Footprint_Prepared.xlsx"
GEOJSON_NAME = "Retail_Footprint_Boundaries.parquet"

APP_DIR = Path(__file__).resolve().parent
SEARCH_DIRS = [APP_DIR, APP_DIR / "data"]

LEVEL_ORDER = ["country", "region", "territory", "district", "county", "subcounty"]
PARENT_COLUMNS = [
    ("region", "region_name"),
    ("territory", "territory_name"),
    ("district", "district_name"),
    ("county", "county_name"),
]

DESIGN_STYLE = {
    "TOUCH POINT":    {"legend_color": "#27AE60", "fa_class": "fa-flag"},
    "CONNECT STORE":  {"legend_color": "#2980B9", "fa_class": "fa-circle"},
    "SERVICE CENTER": {"legend_color": "#FF6A00", "fa_class": "fa-house"},
}
FALLBACK_STYLE = {"legend_color": "#6C757D", "fa_class": "fa-location-dot"}


def design_style(design):
    return DESIGN_STYLE.get(design, FALLBACK_STYLE)


ICON_FONT_PX = 15
ICON_BOX_PX = 20
TOLERANCE = 0.0005

MTN_RAMP = ["#FFF8D6", "#FFE680", MTN_YELLOW, "#C9A000", "#6B5500", MTN_BLACK]


# =========================================================
# FILE LOADING
# =========================================================
def find_file(name):
    for folder in SEARCH_DIRS:
        candidate = folder / name
        if candidate.exists():
            return candidate
    return None


@st.cache_data(show_spinner=False)
def read_file_bytes(path_str, modified_time):
    return Path(path_str).read_bytes()


def get_source_bytes(name, uploader_label):
    path = find_file(name)
    if path is not None:
        return read_file_bytes(str(path), path.stat().st_mtime), str(path)
    uploaded = st.sidebar.file_uploader(uploader_label, key=f"up_{name}")
    if uploaded is not None:
        return uploaded.getvalue(), f"uploaded: {uploaded.name}"
    return None, None


@st.cache_data(show_spinner="Reading tables...")
def load_tables(xlsx_bytes):
    book = pd.ExcelFile(io.BytesIO(xlsx_bytes), engine="openpyxl")
    sheets = {
        name: book.parse(name)
        for name in ["Shops", "Boundaries", "ShopAreaAssignments"]
    }
    # Optional sheet: staff headcount per shop (copied from the 'Details' sheet of the workforce file)
    workforce = book.parse("Workforce") if "Workforce" in book.sheet_names else pd.DataFrame()
    shops = sheets["Shops"]
    bnd = sheets["Boundaries"]
    asg = sheets["ShopAreaAssignments"]
    for df in (shops, bnd, asg):
        df.columns = [str(c).strip() for c in df.columns]

    need = {
        "Shops": (shops, ["ShopID", "ShopName", "DealerName", "Design", "Latitude", "Longitude"]),
        "Boundaries": (bnd, ["key", "label", "admin_level"]),
        "ShopAreaAssignments": (asg, ["ShopID", "AdminLevel", "AreaKey", "AreaLabel"]),
    }
    for sheet, (df, cols) in need.items():
        missing = [c for c in cols if c not in df.columns]
        if missing:
            raise ValueError(f"Sheet '{sheet}' is missing columns: {missing}")

    for col in ["ShopID", "ShopName", "DealerName", "Design"]:
        shops[col] = shops[col].fillna("").astype(str).str.strip()
    shops["Design"] = shops["Design"].str.upper()
    shops["Latitude"] = pd.to_numeric(shops["Latitude"], errors="coerce")
    shops["Longitude"] = pd.to_numeric(shops["Longitude"], errors="coerce")

    for col in ["key", "label", "admin_level"]:
        bnd[col] = bnd[col].fillna("").astype(str).str.strip()
    bnd["admin_level"] = bnd["admin_level"].str.lower()

    for col in ["ShopID", "AdminLevel", "AreaKey", "AreaLabel"]:
        asg[col] = asg[col].fillna("").astype(str).str.strip()
    asg["AdminLevel"] = asg["AdminLevel"].str.lower()
    asg = asg[["ShopID", "AdminLevel", "AreaKey", "AreaLabel"]]

    workforce.columns = [str(c).strip() for c in workforce.columns]

    # Optional sheets: population per area (site catchment populations summed by polygon)
    # and the raw per-site points used for the heatmap
    pop_area = book.parse("AreaPopulation") if "AreaPopulation" in book.sheet_names else pd.DataFrame()
    pop_sites = book.parse("SitePopulation") if "SitePopulation" in book.sheet_names else pd.DataFrame()
    if not pop_area.empty:
        pop_area["AdminLevel"] = pop_area["AdminLevel"].astype(str).str.strip().str.lower()
        pop_area["AreaKey"] = pop_area["AreaKey"].astype(str).str.strip()
    return shops, bnd, asg, workforce, pop_area, pop_sites


@st.cache_data(show_spinner="Preparing boundary shapes...")
def level_geometries(parquet_bytes, level, tolerance):
    df = pq.read_table(io.BytesIO(parquet_bytes)).to_pandas()
    df = df[df["admin_level"] == level]

    out = {}
    for _, row in df.iterrows():
        g = row["geometry_json"]
        if not isinstance(g, str) or not g:
            continue
        geom = shape(json.loads(g))
        if tolerance > 0:
            geom = geom.simplify(tolerance, preserve_topology=True)
        if geom.is_empty:
            continue
        out[row["key"]] = {"geometry": mapping(geom), "bounds": geom.bounds}
    return out


# =========================================================
# SMALL HELPERS
# =========================================================
def slug(text):
    return re.sub(r"[^a-z0-9]+", "_", str(text).lower()).strip("_") or "x"


POP_RAMP = ["#FFF3C4", "#FFD27A", "#FCA94A", "#F27B2C", "#D94A1E", "#A8201A", "#5E0B15"]
POP_VIEWS = {
    "Dots (1 dot = 1,000 people)": "Dots",
    "Density (people per km²)": "Population density",
    "Total population": "Population",
}
DOT_PEOPLE = 1000
DOT_COLOR = (178, 34, 34)   # firebrick red; overlaps darken, so dense areas read as solid red
# typical catchment spread (km) around a site, by site area type
SPREAD_KM = {"dense urban": 0.6, "urban": 1.2, "sub urban": 2.5, "rural": 5.0}


@st.cache_data(show_spinner="Scattering population dots...")
def make_population_dots(sites_df, country_json, seed=7):
    """1 dot per 1,000 people, scattered around each site, kept inside Uganda."""
    rng = np.random.default_rng(seed)
    sites_df = sites_df.dropna(subset=["Latitude", "Longitude", "Population"])
    base = sites_df["Population"].to_numpy(float) / DOT_PEOPLE
    n = np.floor(base).astype(int) + (rng.random(len(base)) < (base - np.floor(base))).astype(int)
    idx = np.repeat(np.arange(len(sites_df)), n)
    lon0 = sites_df["Longitude"].to_numpy(float)[idx]
    lat0 = sites_df["Latitude"].to_numpy(float)[idx]
    kinds = sites_df["AreaType"].astype(str).str.strip().str.lower() if "AreaType" in sites_df else None
    sig_km = (kinds.map(SPREAD_KM).fillna(3.0).to_numpy(float) if kinds is not None
              else np.full(len(sites_df), 3.0))[idx]
    country = shape(json.loads(country_json))
    lon, lat = lon0.copy(), lat0.copy()
    todo = np.ones(len(lon), dtype=bool)
    for _ in range(8):   # resample dots that landed in the lake / outside the border
        k = todo.sum()
        if k == 0:
            break
        lon[todo] = lon0[todo] + rng.normal(0, 1, k) * sig_km[todo] / 111.0
        lat[todo] = lat0[todo] + rng.normal(0, 1, k) * sig_km[todo] / 111.0
        todo = ~contains_xy(country, lon, lat)
    keep = ~todo
    return lon[keep], lat[keep]


def add_dots_layer(fmap, lon, lat):
    """Draw the dots as canvas circle markers: constant pixel size, crisp at every zoom, ignores the mouse."""
    pts = json.dumps([[round(float(a), 4), round(float(b), 4)] for a, b in zip(lat, lon)], separators=(",", ":"))
    r, g, b = DOT_COLOR
    js = f"""
    (function() {{
      var map = {fmap.get_name()};
      if (!map.getPane('popDots')) {{
        map.createPane('popDots');
        map.getPane('popDots').style.zIndex = 450;
        map.getPane('popDots').style.pointerEvents = 'none';
      }}
      var renderer = L.canvas({{pane: 'popDots', padding: 0.3}});
      var pts = {pts};
      var layer = L.layerGroup();
      for (var i = 0; i < pts.length; i++) {{
        L.circleMarker(pts[i], {{renderer: renderer, pane: 'popDots', radius: 1.6, stroke: false,
          fillColor: 'rgb({r},{g},{b})', fillOpacity: 0.5, interactive: false}}).addTo(layer);
      }}
      layer.addTo(map);
    }})();
    """
    fmap.get_root().script.add_child(folium.Element(js))


def quantile_colormap(values, caption):
    """Equal-count (quantile) classes so a few huge areas don't wash out the rest."""
    vals = np.asarray([v for v in values if v is not None and np.isfinite(v) and v > 0], dtype=float)
    if vals.size == 0:
        return None
    qs = np.unique(np.quantile(vals, np.linspace(0, 1, len(POP_RAMP) + 1)))
    if len(qs) < 3:
        qs = np.array([vals.min(), vals.max() + 1.0])
    colors = POP_RAMP[: len(qs) - 1]
    cmap = cm.StepColormap(colors, index=list(qs), vmin=float(qs[0]), vmax=float(qs[-1]), caption=caption)
    return cmap


def fmt_big(v):
    if v is None or not np.isfinite(v):
        return "n/a"
    return f"{v:,.0f}"


def level_title(level):
    return {"subcounty": "Subcounty"}.get(level, level.title())


def icon_html(design):
    style = design_style(design)
    return (
        f'<div style="font-size:{ICON_FONT_PX}px;line-height:1;'
        f'color:{style["legend_color"]};'
        f'text-shadow:0 0 2px #fff,0 0 2px #fff,0 0 2px #fff;">'
        f'<i class="fa-solid {style["fa_class"]}"></i></div>'
    )


def region_norm(text):
    """'SOUTH WEST' and 'Southwest' must match -> compare letters/digits only."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def workforce_by_region(wf):
    """{normalised region: (permanent, contract, total)} plus the national total.

    Total staff = MTN Permanent Staff + MTN Contract Staff only.
    """
    if wf.empty or "Region" not in wf.columns or "MTN Permanent Staff" not in wf.columns:
        return {}, 0
    wf = wf.dropna(subset=["Region"]).copy()
    staff_cols = ["MTN Permanent Staff", "MTN Contract Staff"]
    wf[staff_cols] = wf[staff_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    wf["_total"] = wf[staff_cols].sum(axis=1)
    wf["_key"] = wf["Region"].map(region_norm)
    grp = wf.groupby("_key")[["MTN Permanent Staff", "MTN Contract Staff", "_total"]].sum()
    out = {k: (int(r.iloc[0]), int(r.iloc[1]), int(r.iloc[2])) for k, r in grp.iterrows()}
    return out, int(wf["_total"].sum())


def kpi_card_html(label, value):
    return (
        '<div class="mtn-kpi">'
        f'<div class="mtn-kpi-value">{value}</div>'
        f'<div class="mtn-kpi-label">{label}</div>'
        '</div>'
    )


# =========================================================
# STICKY TOP: MASTHEAD + KPIs
# =========================================================
with st.container(key="sticky_top"):
    st.markdown(
        """
        <div class="mtn-header">
            <div class="mtn-logo-chip">
                <span class="mtn-logo-text">MTN</span>
            </div>
            <div>
                <div class="mtn-title">S&amp;D Retail <span>FootPrint</span></div>
                <div class="mtn-subtitle">Business Intelligence · Uganda</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    xlsx_bytes, _ = get_source_bytes(PREPARED_NAME, f"Upload {PREPARED_NAME}")
    geo_bytes, _ = get_source_bytes(GEOJSON_NAME, f"Upload {GEOJSON_NAME}")

    if xlsx_bytes is None or geo_bytes is None:
        st.info(
            f"Place **{PREPARED_NAME}** and **{GEOJSON_NAME}** next to this app "
            "(or in a `data` folder in the repository), or upload them in the sidebar."
        )
        st.stop()

    try:
        shops, boundaries, assignments, workforce, pop_area, pop_sites = load_tables(xlsx_bytes)
    except Exception as exc:
        st.error(f"Could not read the prepared workbook: {exc}")
        st.stop()

    # Data-integrity check: shops whose ShopID is missing from ShopAreaAssignments
    # (e.g. ShopID edited in Shops sheet but not in assignments) silently vanish from the map.
    _unassigned = shops[~shops["ShopID"].isin(set(assignments["ShopID"]))]
    if not _unassigned.empty:
        st.sidebar.warning(
            f"{len(_unassigned)} shop(s) have no area assignment and are not drawn: "
            + ", ".join(f"{n} ({d})" for n, d in zip(_unassigned["ShopName"], _unassigned["Design"]))
        )

    st.sidebar.header("Filters")

    levels_present = [lv for lv in LEVEL_ORDER if lv in set(boundaries["admin_level"])]
    if not levels_present:
        st.error("The Boundaries sheet has no recognised admin levels.")
        st.stop()

    default_level = levels_present.index("region") if "region" in levels_present else 0
    level = st.sidebar.selectbox(
        "Administrative level",
        options=levels_present,
        index=default_level,
        format_func=level_title,
    )

    areas = boundaries[boundaries["admin_level"] == level].copy()
    areas_all_at_level = areas.copy()

    # --- Region dropdown when Region is the selected level ---
    # (Otherwise, region filtering happens through the parent-column loop below.)
    chosen_regions = []
    if level == "region":
        region_options = sorted(
            {str(v).strip() for v in areas["label"].dropna() if str(v).strip()}
        )
        chosen_regions = st.sidebar.multiselect("Region", options=region_options)
        if chosen_regions:
            areas = areas[
                areas["label"].astype(str).str.strip().isin(chosen_regions)
            ]

    for parent_level, column in PARENT_COLUMNS:
        if LEVEL_ORDER.index(parent_level) >= LEVEL_ORDER.index(level):
            continue
        if column not in areas.columns:
            continue
        options = sorted(areas[column].dropna().astype(str).str.strip().unique())
        options = [o for o in options if o]
        if not options:
            continue
        chosen = st.sidebar.multiselect(level_title(parent_level), options=options)
        if chosen:
            areas = areas[areas[column].astype(str).str.strip().isin(chosen)]

    all_designs = sorted(d for d in shops["Design"].unique() if d)
    selected_designs = st.sidebar.multiselect("Shop Design", options=all_designs, default=all_designs)
    if not selected_designs:
        st.warning("Select at least one Design.")
        st.stop()

    show_shops = st.sidebar.checkbox("Show shop points", value=True)

    has_pop = not pop_area.empty
    colour_by = "Shops"   # only used if the workbook has no population sheets
    if has_pop:
        st.sidebar.header("Population")
        _view = st.sidebar.radio(
            "Map view", list(POP_VIEWS),
            help="Density shows where people are packed in, whatever the size of the area. "
                 "Total shows raw headcount per area. Dots scatter 1 dot per 1,000 people.",
        )
        colour_by = POP_VIEWS[_view]

    if areas.empty:
        st.warning("No areas match these filters.")
        st.stop()

    asg_level = assignments[assignments["AdminLevel"] == level].merge(
        shops[["ShopID", "ShopName", "DealerName", "Design", "Latitude", "Longitude"]],
        on="ShopID", how="left",
    )
    asg_level = asg_level[asg_level["Design"].isin(selected_designs)]

    area_keys = areas["key"].tolist()
    counts = (
        asg_level.groupby(["AreaKey", "Design"])["ShopID"].nunique().unstack(fill_value=0)
        .reindex(columns=selected_designs, fill_value=0)
        .reindex(area_keys, fill_value=0)
    )
    counts["Total"] = counts[selected_designs].sum(axis=1)

    if has_pop:
        _pa = pop_area[pop_area["AdminLevel"] == level].set_index("AreaKey")
        pop_by_key = _pa["Population"].reindex(area_keys).fillna(0)
        dens_by_key = _pa["PopDensityKm2"].reindex(area_keys)
    else:
        pop_by_key = pd.Series(0.0, index=area_keys)
        dens_by_key = pd.Series(np.nan, index=area_keys)
    shops_total = counts["Total"].reindex(area_keys)
    pps_by_key = (pop_by_key / shops_total.replace(0, np.nan))

    n_areas = len(areas)
    areas_without_shops = int((counts["Total"] == 0).sum())
    shops_in_areas = int(counts["Total"].sum())
    level_label = level_title(level)

    # KPIs — custom HTML cards, guaranteed wrapping
    kpi_cols = st.columns(3 + len(selected_designs) + (1 if has_pop else 0), gap="small")
    kpi_cols[0].markdown(
        kpi_card_html(f"{level_label} areas shown", f"{n_areas:,}"),
        unsafe_allow_html=True,
    )
    kpi_cols[1].markdown(
        kpi_card_html("Shops in these areas", f"{shops_in_areas:,}"),
        unsafe_allow_html=True,
    )
    kpi_cols[2].markdown(
        kpi_card_html(f"{level_label} areas without shops", f"{areas_without_shops:,}"),
        unsafe_allow_html=True,
    )
    if has_pop:
        kpi_cols[-1].markdown(
            kpi_card_html("Population in these areas", fmt_big(pop_by_key.sum())),
            unsafe_allow_html=True,
        )
    for col, design in zip(kpi_cols[3:3 + len(selected_designs)], selected_designs):
        col.markdown(
            kpi_card_html(
                f"Avg {design.title()} per {level_label}",
                f"{counts[design].sum() / n_areas:.2f}",
            ),
            unsafe_allow_html=True,
        )

# =========================================================
# MAP + TABLE
# =========================================================
geoms = level_geometries(geo_bytes, level, TOLERANCE)

wf_region, wf_national_total = workforce_by_region(workforce)
# Workforce hover on every Region-level view (with or without a region filter)
show_workforce = level == "region" and bool(wf_region)

features, missing_geometry = [], 0
for _, row in areas.iterrows():
    geometry = geoms.get(row["key"])
    if geometry is None:
        missing_geometry += 1
        continue
    props = {
        "key": row["key"], "label": row["label"],
        "level": level_label,
        "total_shops": int(counts.loc[row["key"], "Total"]),
    }
    for design in selected_designs:
        props[f"n_{slug(design)}"] = int(counts.loc[row["key"], design])
    if has_pop:
        _p = float(pop_by_key.get(row["key"], 0))
        _d = dens_by_key.get(row["key"])
        _pp = pps_by_key.get(row["key"])
        _n = int(counts.loc[row["key"], "Total"])
        props["pop_txt"] = fmt_big(_p) if _p > 0 else "no site data"
        props["dens_txt"] = f"{_d:,.0f} / km²" if _d is not None and np.isfinite(_d) and _p > 0 else "n/a"
        props["pps_txt"] = (
            f"{_pp:,.0f}" if _pp is not None and np.isfinite(_pp)
            else ("no shops" if _p > 0 and _n == 0 else "n/a")
        )
        props["_metric"] = {
            "Population": _p if _p > 0 else None,
            "Population density": float(_d) if _d is not None and np.isfinite(_d) and _p > 0 else None,
        }.get(colour_by)
        props["_noshop"] = bool(_p > 0 and _n == 0)
    if show_workforce:
        perm, contract, total = wf_region.get(region_norm(row["label"]), (None, None, None))
        if perm is None:
            props["wf_perm"] = props["wf_contract"] = props["wf_total"] = "n/a"
        else:
            props["wf_perm"] = f"{perm:,}"
            props["wf_contract"] = f"{contract:,}"
            pct = (total / wf_national_total * 100) if wf_national_total else 0
            props["wf_total"] = f"{total:,} ({pct:.1f}% of national)"
    features.append(
        {"type": "Feature", "properties": props, "geometry": geometry["geometry"]}
    )

if not features:
    st.error("None of the selected areas have a polygon in the boundary file.")
    st.stop()
if missing_geometry:
    st.warning(f"{missing_geometry:,} area(s) have no polygon and are not drawn.")

vmax = max(1, int(counts["Total"].max()))
colormap = cm.LinearColormap(MTN_RAMP, vmin=0, vmax=vmax)
colormap.caption = f"Shops per {level_label.lower()}"

metric_cmap = None
if has_pop and colour_by in ("Population", "Population density"):
    _caption = {
        "Population": f"Population per {level_label.lower()} (equal-count classes)",
        "Population density": "People per km² (equal-count classes)",
    }[colour_by]
    metric_cmap = quantile_colormap([f["properties"].get("_metric") for f in features], _caption)


def style_function(feature):
    props = feature["properties"]
    if has_pop and colour_by == "Dots":
        return {"fillColor": "#FFFFFF", "color": "#555555", "weight": 0.8, "fillOpacity": 0.55}
    if metric_cmap is not None:
        v = props.get("_metric")
        if v is None or v <= 0:
            return {"fillColor": "#EFEFEF", "color": "#BDBDBD", "weight": 1, "fillOpacity": 0.45}
        return {"fillColor": metric_cmap(v), "color": MTN_BLACK, "weight": 1, "fillOpacity": 0.78}
    total = props["total_shops"]
    if total == 0:
        return {"fillColor": "#EFEFEF", "color": "#BDBDBD", "weight": 1, "fillOpacity": 0.45}
    return {"fillColor": colormap(total), "color": MTN_BLACK, "weight": 1, "fillOpacity": 0.72}


def highlight_function(feature):
    return {"weight": 3, "color": MTN_YELLOW, "fillOpacity": 0.85}


drawn_keys = {f["properties"]["key"] for f in features}
drawn_bounds = [g["bounds"] for k, g in geoms.items() if k in drawn_keys]
minx = min(b[0] for b in drawn_bounds)
miny = min(b[1] for b in drawn_bounds)
maxx = max(b[2] for b in drawn_bounds)
maxy = max(b[3] for b in drawn_bounds)

# Force OSM tiles — tiles=None prevents any default (CARTO) layer
fmap = folium.Map(
    location=[(miny + maxy) / 2, (minx + maxx) / 2],
    zoom_start=7,
    tiles=None,
)
folium.TileLayer(
    tiles="https://tile.openstreetmap.org/{z}/{x}/{y}.png",
    attr="&copy; OpenStreetMap contributors",
    name="OpenStreetMap",
    max_zoom=19,
    no_wrap=True,
    control=False,
).add_to(fmap)

fmap.get_root().header.add_child(
    folium.Element(
        '<link rel="stylesheet" '
        'href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">'
    )
)

tooltip_fields = ["label", "level"]
tooltip_aliases = ["Area", "Level"]
if has_pop:
    tooltip_fields += ["pop_txt", "dens_txt", "pps_txt"]
    tooltip_aliases += ["Population", "Density", "People per shop"]
tooltip_fields += ["total_shops"] + [f"n_{slug(d)}" for d in selected_designs]
tooltip_aliases += ["Total shops"] + [d.title() for d in selected_designs]
if show_workforce:
    tooltip_fields += ["wf_perm", "wf_contract", "wf_total"]
    tooltip_aliases += ["Permanent staff", "Contract staff", "Total staff"]

folium.GeoJson(
    {"type": "FeatureCollection", "features": features},
    name="Boundaries",
    style_function=style_function,
    highlight_function=highlight_function,
    tooltip=folium.GeoJsonTooltip(
        fields=tooltip_fields, aliases=tooltip_aliases, sticky=True, localize=True,
    ),
).add_to(fmap)
(metric_cmap if metric_cmap is not None else colormap).add_to(fmap)

if has_pop and colour_by == "Dots" and not pop_sites.empty:
    _country = level_geometries(geo_bytes, "country", TOLERANCE)
    if _country:
        _cj = json.dumps(next(iter(_country.values()))["geometry"])
        _dlon, _dlat = make_population_dots(pop_sites, _cj)
        _drawn = [shape(f["geometry"]) for f in features]
        if len(features) < len(areas_all_at_level):   # filtered view: keep dots inside the shown areas
            _keep = contains_xy(unary_union(_drawn), _dlon, _dlat)
            _dlon, _dlat = _dlon[_keep], _dlat[_keep]
        add_dots_layer(fmap, _dlon, _dlat)
        fmap.get_root().html.add_child(folium.Element(
            "<div style='position:fixed;bottom:30px;right:12px;z-index:9999;background:#fff;"
            "padding:8px 12px;border:2px solid #B1121B;border-radius:8px;font-size:12px;"
            "font-family:Montserrat,sans-serif'>"
            "<span style='display:inline-block;width:10px;height:10px;border-radius:50%;"
            "background:#B1121B;margin-right:6px'></span>"
            f"1 dot = {DOT_PEOPLE:,} people ({len(_dlon):,} dots)</div>"
        ))

if show_shops:
    shop_points = (
        asg_level[asg_level["AreaKey"].isin(area_keys)]
        .drop_duplicates("ShopID")
        .dropna(subset=["Latitude", "Longitude"])
    )
    shop_layer = folium.FeatureGroup(name="Shops")
    for _, shop in shop_points.iterrows():
        popup_html = (
            f"<b>{shop['ShopName']}</b><br/>Dealer: {shop['DealerName']}<br/>"
            f"Design: {str(shop['Design']).title()}<br/>{level_label}: {shop['AreaLabel']}"
        )
        folium.Marker(
            location=[shop["Latitude"], shop["Longitude"]],
            icon=folium.DivIcon(
                html=icon_html(shop["Design"]),
                icon_size=(ICON_BOX_PX, ICON_BOX_PX),
                icon_anchor=(ICON_BOX_PX / 2, ICON_BOX_PX / 2),
            ),
            tooltip=f"{shop['ShopName']} ({str(shop['Design']).title()})",
            popup=folium.Popup(popup_html, max_width=300),
        ).add_to(shop_layer)
    shop_layer.add_to(fmap)

    legend_items = "".join(
        f"<div><span style='display:inline-block;width:10px;height:10px;"
        f"border-radius:50%;background:{design_style(d)['legend_color']};"
        f"margin-right:6px'></span>{d}</div>"
        for d in selected_designs
    )
    fmap.get_root().html.add_child(
        folium.Element(
            "<div style='position:fixed;bottom:30px;left:12px;z-index:9999;"
            "background:#000;color:#fff;padding:10px 12px;border:2px solid #FFCC08;"
            "border-radius:8px;font-size:12px;font-family:Montserrat,sans-serif'>"
            "<div style='color:#FFCC08;font-weight:800;letter-spacing:1px;"
            "text-transform:uppercase;margin-bottom:6px'>Shop Design</div>"
            f"{legend_items}</div>"
        )
    )

folium.LayerControl(collapsed=True).add_to(fmap)
fmap.fit_bounds([[miny, minx], [maxy, maxx]], padding=(10, 10))

# ---- Full-width map ----
st.markdown(
    '<div class="mtn-section-label">Uganda Retail Footprint</div>',
    unsafe_allow_html=True,
)
st_folium(
    fmap,
    use_container_width=True,
    height=720,
    returned_objects=[],
    key=f"map_v5_{level}_{colour_by}",   # bumped so cached old map is discarded
)

# ---------------------------------------------------------
# Table
# ---------------------------------------------------------
st.markdown(
    f'<div class="mtn-section-label">Shops per {level_label.lower()}</div>',
    unsafe_allow_html=True,
)

table = areas[["key", "label"]].copy()
for _, column in PARENT_COLUMNS:
    if column in areas.columns and column not in table.columns:
        table[column] = areas[column]
table = table.merge(
    counts.reset_index().rename(columns={"index": "key", "AreaKey": "key"}),
    on="key", how="left",
)

if has_pop:
    table["Population"] = table["key"].map(pop_by_key).fillna(0).round(0).astype(int)
    table["Pop density (per km²)"] = table["key"].map(dens_by_key).round(0)
    table["People per shop"] = table["key"].map(pps_by_key).round(0)

status = table["Total"].apply(lambda t: "Without shops" if t == 0 else "With shops")
insert_at = list(table.columns).index("Total") + 1
table.insert(insert_at, "Shops status", status)

table = table.sort_values(["Total", "label"], ascending=[False, True]).reset_index(drop=True)

st.dataframe(table, hide_index=True, use_container_width=True)
st.download_button(
    "Download table (CSV)",
    data=table.to_csv(index=False).encode("utf-8"),
    file_name=f"shops_per_{level}.csv",
    mime="text/csv",
)
