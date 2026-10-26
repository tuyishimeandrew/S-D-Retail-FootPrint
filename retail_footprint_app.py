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
import pandas as pd
import pyarrow.parquet as pq
import streamlit as st
from shapely.geometry import mapping, shape
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
    return shops, bnd, asg, workforce


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
        shops, boundaries, assignments, workforce = load_tables(xlsx_bytes)
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

    n_areas = len(areas)
    areas_without_shops = int((counts["Total"] == 0).sum())
    shops_in_areas = int(counts["Total"].sum())
    level_label = level_title(level)

    # KPIs — custom HTML cards, guaranteed wrapping
    kpi_cols = st.columns(3 + len(selected_designs), gap="small")
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
    for col, design in zip(kpi_cols[3:], selected_designs):
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
# Workforce hover only on the national view: Region level with no region filter applied
show_workforce = level == "region" and not chosen_regions and bool(wf_region)

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


def style_function(feature):
    total = feature["properties"]["total_shops"]
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

tooltip_fields = ["label", "level", "total_shops"] + [f"n_{slug(d)}" for d in selected_designs]
tooltip_aliases = ["Area", "Level", "Total shops"] + [d.title() for d in selected_designs]
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
colormap.add_to(fmap)

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
    key=f"map_v3_{level}",   # bumped so cached old map is discarded
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
