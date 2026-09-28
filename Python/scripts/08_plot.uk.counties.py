#!/usr/bin/env python3
from __future__ import annotations

"""
08_plot.uk.counties.py

Purpose:
    Plot seroprevalence by UK county group from a user-selected Excel file.

Input:
    Excel workbook containing a positive/negative test result table with county names.

Output:
    PNG file containing a map of the UK with county groups colored by seroprevalence.
    .csv file containing a contingency table of positive/negative counts for each county group.

Thesis:
    Chapter: ERVEV biosurveillance survey

Author:
    Igor Starinskij

AI assistance:
    Generative artificial intelligence tools, primarily GitHub Copilot 0.63.0 (GPT-5.6 Luna), were used during the development of software used in this thesis.
	AI was used to assist with Python code generation, debugging, code refinement and explanation of programming concepts.
	Generated code was reviewed, modified, tested and validated by the author.
	The analytical methodology, statistical decisions, interpretation of results and final responsibility for the analyses remain with the author.
"""

import argparse
import re
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog

import geopandas as gpd
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, Normalize, to_hex
from matplotlib.cm import ScalarMappable
import numpy as np
import pandas as pd
import requests
from arcgis2geojson import arcgis2geojson
from shapely.geometry import MultiLineString, MultiPolygon, Polygon
from shapely.ops import unary_union


# ONS Open Geography ArcGIS service for UK county and unitary boundaries.
COUNTY_SERVICE_QUERY_URL = (
    "https://services1.arcgis.com/ESMARspQHYMw9BZ9/arcgis/rest/services/"
    "Counties_and_Unitary_Authorities_December_2023_Boundaries_UK_BFC/"
    "FeatureServer/0/query"
)

# Conservative simplification for faster, more reliable transfer.
DEFAULT_MAX_ALLOWABLE_OFFSET = 1000
MAX_NON_BORDER_GROUP_DISTANCE_METRES = 150_000

COUNTY_COLUMN_CANDIDATES = [
    "county",
    "ceremonial_county",
    "ceremonial county",
    "admin_county",
    "region_county",
]

CALL_COLUMN_CANDIDATES = [
    "call",
]

COUNTY_ALIASES = {
    "city of bristol": "bristol",
    "county durham": "durham",
    "ayrshire and arran": "north ayrshire",
    "berkshire": "west berkshire",
    "cumbria": "cumberland",
    "dumfries": "dumfries and galloway",
    "inverness": "highland",
    "kincardineshire": "aberdeenshire",
    "mid glamorgan": "rhondda cynon taf",
    "nairn": "highland",
    "northamptonshire": "west northamptonshire north northamptonshire",
    "ross and cromarty": "highland",
    "stirling and falkirk": "stirling",
    "sutherland": "highland",
    "west glamorgan": "swansea",
    "herefordshire county of": "herefordshire",
    "kingston upon hull": "east riding of yorkshire",
    "city of kingston upon hull": "east riding of yorkshire",
    "city of london": "greater london",
    "london": "greater london",
    "banffshire": "aberdeenshire moray",
    "bedfordshire": "bedford central bedfordshire",
    "clackmannan": "clackmannanshire",
    "roxburgh ettrick and lauder": "scottish borders",
    "roxburgh ettrick and lauderdale": "scottish borders",
    "tweeddale": "scottish borders",
    "tyne and wear": "tyne and wear",
}

HISTORIC_COMBINED_REGIONS = {
    "aberdeenshire moray": ("Aberdeenshire, Moray", ["Aberdeenshire", "Moray"]),
    "bedford central bedfordshire": (
        "Bedford, Central Bedfordshire",
        ["Bedford", "Central Bedfordshire"],
    ),
    "west northamptonshire north northamptonshire": (
        "Northamptonshire",
        ["West Northamptonshire", "North Northamptonshire"],
    ),
    "tyne and wear": (
        "Tyne and Wear",
        [
            "Newcastle upon Tyne",
            "North Tyneside",
            "South Tyneside",
            "Sunderland",
            "Gateshead",
        ],
    ),
}


def normalize_name(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[&]", " and ", text)
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return COUNTY_ALIASES.get(text, text)


def pick_column(columns: list[str], explicit: str | None, candidates: list[str]) -> str:
    if explicit:
        if explicit not in columns:
            raise ValueError(f"Column '{explicit}' not found. Available columns: {columns}")
        return explicit

    lower_to_original = {c.lower(): c for c in columns}
    for cand in candidates:
        if cand in lower_to_original:
            return lower_to_original[cand]

    for col in columns:
        col_l = col.lower()
        if any(cand in col_l for cand in candidates):
            return col

    raise ValueError(f"Could not detect required column. Available columns: {columns}")


def choose_input_file() -> Path:
    root = tk.Tk()
    root.withdraw()
    selected = filedialog.askopenfilename(
        title="Select DSS summary file",
        filetypes=[
            ("Spreadsheet files", "*.xlsx *.xls *.csv"),
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*"),
        ],
    )
    root.destroy()
    if not selected:
        raise RuntimeError("No input file selected.")
    return Path(selected)


def read_input_table(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        xls = pd.ExcelFile(path)
        for name in xls.sheet_names:
            df = pd.read_excel(path, sheet_name=name)
            cols = [str(c).lower() for c in df.columns]
            if any("county" in c for c in cols) and any("call" in c for c in cols):
                return df
        return pd.read_excel(path, sheet_name=xls.sheet_names[0])

    if suffix == ".csv":
        return pd.read_csv(path)

    raise ValueError("Unsupported input format. Use .xlsx, .xls, or .csv")


def load_or_download_uk_counties(cache_path: Path) -> gpd.GeoDataFrame:
    if cache_path.exists():
        gdf = gpd.read_file(cache_path)
        if gdf.crs is None:
            gdf = gdf.set_crs(27700)
        return gdf

    params = {
        "where": "1=1",
        "outFields": "CTYUA23NM,FID",
        "returnGeometry": "true",
        "outSR": 27700,
        "maxAllowableOffset": DEFAULT_MAX_ALLOWABLE_OFFSET,
        "f": "json",
    }
    response = requests.get(COUNTY_SERVICE_QUERY_URL, params=params, timeout=120)
    response.raise_for_status()
    esri_json = response.json()
    if "features" not in esri_json:
        raise RuntimeError("Failed to download UK county map features.")

    geojson = arcgis2geojson(esri_json)
    gdf = gpd.GeoDataFrame.from_features(geojson["features"])
    gdf = gdf.set_crs(27700)
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    gdf.to_file(cache_path, driver="GeoJSON")
    return gdf


def add_historic_analysis_regions(counties: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Add combined modern ONS regions needed for historic input names."""
    if "CTYUA23NM" not in counties.columns:
        return counties

    result = counties.copy()
    normalized_map_names = result["CTYUA23NM"].map(normalize_name)
    additions = []
    for combined_key, (display_name, source_names) in HISTORIC_COMBINED_REGIONS.items():
        if combined_key in normalized_map_names.tolist():
            continue
        source_keys = [normalize_name(name) for name in source_names]
        source_geometries = result.loc[
            normalized_map_names.isin(source_keys), "geometry"
        ]
        if len(source_geometries) != len(source_keys):
            missing = sorted(set(source_names) - set(result.loc[
                normalized_map_names.isin(source_keys), "CTYUA23NM"
            ]))
            raise ValueError(
                f"Could not build historic analysis region '{display_name}'; "
                f"missing map counties: {missing}"
            )
        repaired_geometries = [
            geometry if geometry.is_valid else geometry.buffer(0)
            for geometry in source_geometries
        ]
        additions.append(
            {
                "CTYUA23NM": display_name,
                "geometry": unary_union(repaired_geometries),
            }
        )

    if additions:
        result = pd.concat([result, gpd.GeoDataFrame(additions, crs=counties.crs)], ignore_index=True)
    return result


def build_seroprevalence_counts(
    table: pd.DataFrame,
    counties: gpd.GeoDataFrame,
    county_col: str,
    call_col: str,
) -> tuple[gpd.GeoDataFrame, list[str]]:
    """Calculate positive/tested counts for each county.

    Call values of pos are positive and neg are negative. Other values are
    excluded from the seroprevalence denominator.
    """
    df = table.copy()
    df["_county_key"] = df[county_col].map(normalize_name)
    df["_call_norm"] = df[call_col].astype(str).str.strip().str.lower()
    df = df[df["_county_key"].ne("")].copy()

    map_name_col = "CTYUA23NM" if "CTYUA23NM" in counties.columns else counties.columns[0]
    county_shapes = counties[[map_name_col, "geometry"]].copy()
    county_shapes["_county_key"] = county_shapes[map_name_col].map(normalize_name)
    county_shapes = county_shapes.drop_duplicates(subset=["_county_key"])

    unmatched = sorted(
        df.loc[~df["_county_key"].isin(county_shapes["_county_key"]), county_col]
        .dropna()
        .unique()
    )

    tested = df[df["_call_norm"].isin({"pos", "neg"})].copy()
    tested["_positive"] = tested["_call_norm"].eq("pos")
    counts = tested.groupby("_county_key").agg(
        positive_count=("_positive", "sum"),
        tested_count=("_positive", "size"),
    )

    county_shapes = county_shapes.join(counts, on="_county_key")
    county_shapes["positive_count"] = county_shapes["positive_count"].fillna(0).astype(int)
    county_shapes["tested_count"] = county_shapes["tested_count"].fillna(0).astype(int)
    county_shapes["seroprevalence"] = np.where(
        county_shapes["tested_count"] > 0,
        county_shapes["positive_count"] / county_shapes["tested_count"],
        np.nan,
    )
    return county_shapes, unmatched


def group_low_sample_counties(
    counties: gpd.GeoDataFrame,
    minimum_n: int = 20,
) -> tuple[gpd.GeoDataFrame, dict[str, list[str]]]:
    """Group sampled counties below minimum_n with nearby sampled counties.

    Bordering groups are preferred. If no border exists, the closest sampled
    group is used. Counties with zero tested samples remain ungrouped.
    """
    result = counties.copy().reset_index(drop=True)
    result["geometry"] = result.geometry.apply(
        lambda geometry: geometry if geometry.is_valid else geometry.buffer(0)
    )
    county_keys = result["_county_key"].tolist()
    sample_n = dict(zip(county_keys, result["tested_count"].astype(int)))
    positive_n = dict(zip(county_keys, result["positive_count"].astype(int)))
    geometries = dict(zip(county_keys, result.geometry))

    groups: dict[str, set[str]] = {key: {key} for key in county_keys}
    group_n = sample_n.copy()
    group_positive = positive_n.copy()

    # Build shared-border lengths once. The boundary length is a useful proxy
    # for geographic closeness and keeps Scottish/Welsh regions local.
    border_lengths: dict[tuple[str, str], float] = {}
    for left_index, left_key in enumerate(county_keys):
        for right_key in county_keys[left_index + 1:]:
            shared_length = geometries[left_key].boundary.intersection(
                geometries[right_key].boundary
            ).length
            if shared_length > 0:
                border_lengths[(left_key, right_key)] = shared_length
                border_lengths[(right_key, left_key)] = shared_length

    def group_geometry(group_id: str):
        return unary_union([geometries[key] for key in groups[group_id]])

    blocked_groups: set[str] = set()
    while True:
        low_groups = [
            group_id
            for group_id in groups
            if 0 < group_n[group_id] < minimum_n and group_id not in blocked_groups
        ]
        if not low_groups:
            break

        source = min(low_groups, key=lambda group_id: group_n[group_id])
        candidates = [
            group_id
            for group_id in groups
            if group_id != source and group_n[group_id] > 0
        ]
        if not candidates:
            break

        source_members = groups[source]

        def shared_border_with(candidate: str) -> float:
            return sum(
                border_lengths.get((source_member, candidate_member), 0.0)
                for source_member in source_members
                for candidate_member in groups[candidate]
            )

        bordering = [candidate for candidate in candidates if shared_border_with(candidate) > 0]
        source_shape = group_geometry(source)
        if bordering:
            search_candidates = bordering
        else:
            nearby = [
                candidate
                for candidate in candidates
                if source_shape.distance(group_geometry(candidate))
                <= MAX_NON_BORDER_GROUP_DISTANCE_METRES
            ]
            if not nearby:
                # Keep the small county as its own group rather than joining
                # it to a geographically implausible distant county.
                blocked_groups.add(source)
                continue
            search_candidates = nearby

        def candidate_score(candidate: str) -> tuple[int, float, float]:
            reaches_threshold = int(group_n[source] + group_n[candidate] >= minimum_n)
            border_score = shared_border_with(candidate)
            distance_score = -source_shape.distance(group_geometry(candidate))
            return reaches_threshold, border_score, distance_score

        target = max(search_candidates, key=candidate_score)
        groups[target].update(groups[source])
        group_n[target] += group_n[source]
        group_positive[target] += group_positive[source]
        del groups[source]
        del group_n[source]
        del group_positive[source]
        blocked_groups.discard(target)

    key_to_group = {
        county_key: group_id
        for group_id, members in groups.items()
        for county_key in members
    }
    result["group_id"] = result["_county_key"].map(key_to_group)
    result["group_n"] = result["group_id"].map(group_n).astype(int)
    result["group_positive_count"] = result["group_id"].map(group_positive).astype(int)
    result["group_seroprevalence"] = np.where(
        result["group_n"] > 0,
        result["group_positive_count"] / result["group_n"],
        np.nan,
    )
    group_members = {
        group_id: sorted(members)
        for group_id, members in groups.items()
        if len(members) > 1
    }
    return result, group_members


def build_contingency_table(counties: gpd.GeoDataFrame) -> pd.DataFrame:
    """Build one neg/pos contingency row for each sampled county group."""
    map_name_col = "CTYUA23NM" if "CTYUA23NM" in counties.columns else counties.columns[0]

    def format_group_names(names) -> str:
        unique_names = {str(name) for name in names}
        for display_name, source_names in HISTORIC_COMBINED_REGIONS.values():
            if display_name in unique_names:
                unique_names.difference_update(source_names)
        return ", ".join(sorted(unique_names))

    table = (
        counties[counties["group_n"] > 0]
        .groupby("group_id", sort=False)
        .agg(
            county_group=(
                map_name_col,
                format_group_names,
            ),
            pos=("group_positive_count", "first"),
            tested=("group_n", "first"),
        )
        .reset_index(drop=True)
    )
    table["neg"] = table["tested"] - table["pos"]
    table["pos%"] = (100 * table["pos"] / table["tested"]).round(1)
    return table[["county_group", "neg", "pos", "pos%"]]


def plot_data_boundaries(
    counties: gpd.GeoDataFrame,
    ax,
    border_color: str = "#222222",
) -> None:
    """Draw thick outlines around groups and well-sampled standalone counties."""
    group_sizes = counties.groupby("group_id")["group_id"].transform("size")
    highlighted = counties[
        (counties["group_n"] > 0)
        & ((group_sizes > 1) | (counties["group_n"] >= 20))
    ]
    if highlighted.empty:
        return

    # Standalone counties are drawn from their own exterior only. This avoids
    # adding any shared or interior line from the full county boundary.
    standalone = highlighted[group_sizes.loc[highlighted.index].eq(1)]
    if not standalone.empty:
        standalone = standalone.copy()
        standalone["geometry"] = standalone.geometry.apply(
            lambda geometry: geometry if geometry.is_valid else geometry.buffer(0)
        )
        standalone.boundary.plot(
            ax=ax,
            color=border_color,
            linewidth=1.6,
            zorder=3,
        )

    grouped = highlighted[group_sizes.loc[highlighted.index].gt(1)]
    if grouped.empty:
        return

    grouped = grouped.copy()
    grouped["geometry"] = grouped.geometry.apply(
        lambda geometry: geometry if geometry.is_valid else geometry.buffer(0)
    )
    dissolved = grouped.dissolve(by="group_id", as_index=False)
    # Plot only the exterior rings. In particular, do not plot the complete
    # dissolved boundary, which can expose tiny internal topology gaps.
    dissolved["geometry"] = dissolved.geometry.apply(
        lambda geometry: geometry if geometry.is_valid else geometry.buffer(0)
    )
    def exterior_rings(geometry):
        if isinstance(geometry, Polygon):
            return geometry.exterior
        if isinstance(geometry, MultiPolygon):
            return MultiLineString([polygon.exterior for polygon in geometry.geoms])
        return geometry.boundary

    dissolved["geometry"].apply(exterior_rings).plot(
        ax=ax,
        color=border_color,
        linewidth=1.6,
        zorder=3,
    )


def plot_seroprevalence_map(
    counties: gpd.GeoDataFrame,
    output_path: Path,
    title: str,
    scale_max: float = 1.0,
) -> None:
    """Plot seroprevalence without sample dots."""
    plot_counties = counties.to_crs(27700)
    prevalence_gradient = LinearSegmentedColormap.from_list(
        "seroprevalence_yellow_red",
        ["#fffdf5", "#fff3a6", "#fdae61", "#d7301f", "#7f0000"],
    )
    scale_max = max(float(scale_max), 0.000001)
    prevalence_norm = Normalize(vmin=0, vmax=scale_max)
    county_colors = [
        "#eeeeee"
        if pd.isna(prevalence)
        else to_hex(prevalence_gradient(prevalence_norm(prevalence)))
        for prevalence in plot_counties["group_seroprevalence"]
    ]

    fig, ax = plt.subplots(figsize=(10, 13))
    plot_counties.plot(
        ax=ax,
        color=county_colors,
        edgecolor="#7f8c8d",
        linewidth=0.5,
    )
    plot_data_boundaries(plot_counties, ax, border_color="#222222")
    ax.set_title(title)
    ax.set_axis_off()

    scalar_mappable = ScalarMappable(norm=prevalence_norm, cmap=prevalence_gradient)
    scalar_mappable.set_array([])
    colorbar = fig.colorbar(scalar_mappable, ax=ax, fraction=0.035, pad=0.02)
    colorbar.set_label("Putative seroprevalence")
    colorbar_ticks = np.linspace(0, scale_max, 5)
    colorbar.set_ticks(colorbar_ticks)
    colorbar.set_ticklabels([f"{tick:.0%}" for tick in colorbar_ticks])
    plt.tight_layout()
    fig.savefig(output_path, dpi=300)
    plt.close(fig)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Plot seroprevalence by UK county group")
    parser.add_argument("--input", type=Path, default=None, help="Input Excel/CSV file")
    parser.add_argument("--county-col", default=None, help="County column name")
    parser.add_argument("--title", default="Seroprevalence By UK County Group", help="Map title")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = args.input if args.input is not None else choose_input_file()

    table = read_input_table(input_path)
    county_col = pick_column(list(table.columns), args.county_col, COUNTY_COLUMN_CANDIDATES)
    call_col = pick_column(list(table.columns), None, CALL_COLUMN_CANDIDATES)

    cache_path = Path.home() / ".uk_county_map_cache" / "uk_counties_2023.geojson"
    counties = add_historic_analysis_regions(load_or_download_uk_counties(cache_path))

    prevalence_counties, prevalence_unmatched = build_seroprevalence_counts(
        table=table,
        counties=counties,
        county_col=county_col,
        call_col=call_col,
    )
    grouped_counties, county_groups = group_low_sample_counties(prevalence_counties)

    output_path = input_path.with_name(
        f"{input_path.stem}_uk_county_seroprevalence_map.png"
    )
    contingency_output_path = input_path.with_name(
        f"{input_path.stem}_county_contingency_table.csv"
    )
    observed_peak = grouped_counties["group_seroprevalence"].max()
    observed_peak = float(observed_peak) if pd.notna(observed_peak) else 1.0
    contingency_table = build_contingency_table(grouped_counties)
    contingency_table.to_csv(contingency_output_path, index=False)
    plot_seroprevalence_map(
        counties=grouped_counties,
        output_path=output_path,
        title=args.title,
        scale_max=observed_peak,
    )

    print(f"Input file: {input_path}")
    print(f"Output map: {output_path}")
    print(f"Contingency table: {contingency_output_path}")
    print(f"Observed peak seroprevalence: {observed_peak:.1%}")
    print(f"Samples counted for grouping/shading: {int(prevalence_counties['tested_count'].sum())}")
    print(f"Positive/negative samples used for seroprevalence: {int(prevalence_counties['tested_count'].sum())}")
    print(f"Low-sample county groups created: {len(county_groups)}")
    if prevalence_unmatched:
        print("Unmatched county names for seroprevalence:")
        for name in prevalence_unmatched:
            print(f"  - {name}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
