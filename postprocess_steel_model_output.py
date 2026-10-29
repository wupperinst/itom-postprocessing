"""Postprocessing of ITOM-Steel model output.

Reads the long-format csv files of one scenario
(<model_data_dir>/<scenario>/<scenario>_<Variable>.csv), derives the indicators of section 7.99 of
OutputAnalyzer_AKmod_v1_8.R and writes them in wide format (one column per year) to one excel file
with one sheet per indicator ("indicator_sheets"). Optionally the same results are also written in the
IAMC-like format ("iamc"): one sheet per variable (Production, Final Energy, Material Consumption) with the columns Model,
Scenario, Region, Name, Unit and one column per year. `output_formats` in config_steel.yaml (or --formats) selects what is written.

Usage:  python postprocess_steel_model_output.py [path/to/config_steel.yaml] [--formats indicator_sheets iamc]
        (default: config_steel.yaml next to this script)
"""

import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import yaml
from openpyxl.styles import Font

# Columns that describe an entry; the one remaining column of a csv holds the value
# (named like the file, e.g. LocalProduction)
BASE_VARIABLES = ["REGION", "LOCATION", "TECHNOLOGY", "MODE_OF_OPERATION", "PRODUCT", "EMISSION", "YEAR"]

NEEDED_VARIABLES = [
    "LocalProduction", "LocalProductionByTechnology", "LocalProductionByMode", "LocalUse",
    "LocalUseByTechnology", "LocalCapitalInvestment", "LocalTotalCapacity", "LocalTechnologyEmission",
    "NewCapacity", "TotalCapacity", "Use", "TotalDiscountedCost",
]

# Same order as in the R code
SHEET_ORDER = [
    "total_steel_production", "country_steel_production", "total_steel_production_agg",
    "country_steel_production_agg", "total_energy_consumption", "country_energy_consumption",
    "total_iron_production", "country_iron_production", "total_capacity_investments",
    "total_capital_investments", "country_capital_investment", "total_installed_capacities",
    "country_installed_capacities", "total_emissions", "captured_emissions_by_route",
    "captured_emissions_by_cluster", "total_material_consumption", "country_material_consumption",
    "country_mat_consumption_agg", "marginal_production_cost", "total_discounted_cost",
    "country_slag_production",
]


OUTPUT_FORMATS = ["indicator_sheets", "iamc"]


# 0. Configuration ######

parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
parser.add_argument("config", nargs="?", default=Path(__file__).resolve().parent / "config_steel.yaml",
                    help="path to config_steel.yaml (default: next to this script)")
parser.add_argument("--formats", nargs="+", choices=OUTPUT_FORMATS,
                    help="output formats; overrides output_formats in config_steel.yaml")
args = parser.parse_args()

config_path = Path(args.config).resolve()
with open(config_path, encoding="utf-8") as f:
    cfg = yaml.safe_load(f)
root_dir = config_path.parent


def resolve_path(p):
    p = Path(p).expanduser()
    return p if p.is_absolute() else root_dir / p


scenario = cfg["scenario"]
model_dir = resolve_path(Path(cfg["paths"]["model_data_dir"]) / scenario)
output_dir = resolve_path(cfg["paths"]["output_dir"])
output_file = output_dir / f"{scenario}_postprocessed_data.xlsx"
iamc_file = output_dir / f"{scenario}_IAMC.xlsx"

output_formats = args.formats or cfg.get("output_formats") or ["indicator_sheets"]
if unknown_formats := [f for f in output_formats if f not in OUTPUT_FORMATS]:
    sys.exit(f"Unknown output format(s): {', '.join(unknown_formats)} (available: {', '.join(OUTPUT_FORMATS)})")

cap2act = cfg["cap2act"]
coking_coal_energy_factor = cfg["coking_coal_energy_factor"]
years_all = cfg["years"]
years_investment = cfg["investment_years"]
years_ccs = cfg["ccs_years"]
digits = cfg["rounding_digits"]

if not model_dir.is_dir():
    sys.exit(f"Model output folder not found: {model_dir}")
print(f"Scenario: {scenario}")


# 1. Reading the model output ######

def model_file(variable):
    path = model_dir / f"{scenario}_{variable}.csv"
    if not path.is_file():
        sys.exit(f"Model output file not found: {path}")
    return path


def read_model_variable(variable):
    data = pd.read_csv(model_file(variable), dtype=str, keep_default_na=False)

    value_columns = [c for c in data.columns if c not in BASE_VARIABLES]
    if len(value_columns) != 1:
        sys.exit(f"{variable}.csv: expected exactly one value column, found: {', '.join(value_columns)}")

    data = data.rename(columns={value_columns[0]: "VALUE"})
    data["VALUE"] = pd.to_numeric(data["VALUE"])
    if "YEAR" in data.columns:
        data["YEAR"] = pd.to_numeric(data["YEAR"]).astype(int)
    return data


# 2. Region grouping and parameters ######

rg = cfg["region_grouping"]
region_grouping_file = resolve_path(cfg["paths"]["region_grouping_file"])

region_grouping = pd.read_excel(region_grouping_file, sheet_name=rg["sheet"]).rename(columns={
    rg["state_column"]: "STATE", rg["country_column"]: "Country",
    rg["region_column"]: "Region_steel", rg["dummy_share_column"]: "share",
})[["STATE", "Country", "Region_steel", "share"]]

if region_grouping["STATE"].duplicated().any():
    sys.exit("Region grouping: one state is assigned to two regions")

share_sums = region_grouping.groupby("Region_steel")["share"].sum()
if (bad_shares := share_sums[(share_sums - 1).abs() > 1e-6]).size:
    print(f"WARNING: dummy location shares do not sum to 1 in region(s): {', '.join(bad_shares.index)}")

# Dummy locations (e.g. "BeNeLux_000") are split to the states of their region
dummy_split = pd.DataFrame({
    "LOCATION": region_grouping["Region_steel"] + "_000",
    "STATE": region_grouping["STATE"],
    "share": region_grouping["share"],
})

parameters_sheet = pd.read_excel(region_grouping_file, sheet_name=cfg["parameters"]["sheet"], dtype=str)
parameters = {name: list(pd.unique(col.dropna())) for name, col in parameters_sheet.items()}


def get_list(name):
    values = parameters.get(name) or cfg.get(name)
    if not values:
        sys.exit(f"List '{name}' is neither in the Parameters sheet nor in config_steel.yaml")
    return list(values)


iron_products = get_list("iron_products")
iron_production_technologies = get_list("iron_production_technologies")
LS_products = get_list("LS_products")
primary_energy_products = get_list("primary_energy_products")
cs_production_technologies = get_list("cs_production_technologies")
fs_production_technologies = get_list("fs_production_technologies")

capacity_technologies = (iron_production_technologies + cfg["capacity_extra_technologies"]
                         + cs_production_technologies)
investment_technologies = (iron_production_technologies + cfg["investment_extra_technologies"]
                           + cs_production_technologies + fs_production_technologies)


# 3. Spatial aggregation ######

# Location-based variables are available on country and on Europe level, like OUTPUT_Country and
# OUTPUT_Europe in the R code:
#  - steel sites (e.g. DE719_Bremen) are assigned to their state (first two letters of the name)
#  - dummy locations (<Region>_000) are split to the states with the shares of the region grouping
#  - states are summed to countries / Europe
#  - Europe_000 becomes "Europe"; TRANSPORT_HUB_* is dropped; energy_center_location and the demand locations
#    keep their name (they are neither a country nor a part of the Europe total)
# Variables without a LOCATION column (region-level variables) are used as they are.

def aggregate_locations(data):
    if "LOCATION" not in data.columns:
        return data, data

    dims = [c for c in data.columns if c not in ("LOCATION", "VALUE")]
    loc = data["LOCATION"]

    # Non-site locations that keep their name: transport hubs, energy center and the demand locations
    # ("Europe_Demand_location" in older runs, "<Region>_demand_hub" in newer ones)
    is_other = (loc.str.contains("TRANSPORT_HUB", regex=False) | loc.str.contains("energy_center", regex=False)
                | (loc == "Europe_Demand_location") | loc.str.endswith("_demand_hub"))
    is_europe_dummy = loc == "Europe_000"
    is_dummy = loc.str.contains("_000", regex=False) & ~is_europe_dummy & ~is_other
    is_site = ~(is_other | is_europe_dummy | is_dummy)

    sites = data[is_site].assign(STATE=lambda d: d["LOCATION"].str[:2])

    dummies = data[is_dummy].merge(dummy_split, on="LOCATION", how="left")
    if dummies["STATE"].isna().any():
        sys.exit("Dummy location(s) not in the region grouping: "
                 + ", ".join(sorted(dummies.loc[dummies["STATE"].isna(), "LOCATION"].unique())))
    dummies["VALUE"] = dummies["VALUE"] * dummies["share"]

    states = (pd.concat([sites[dims + ["STATE", "VALUE"]], dummies[dims + ["STATE", "VALUE"]]])
              .groupby(dims + ["STATE"], dropna=False, as_index=False)["VALUE"].sum()
              .merge(region_grouping[["STATE", "Country"]], on="STATE", how="left"))
    if states["Country"].isna().any():
        sys.exit("State(s) not in the region grouping: " + ", ".join(sorted(states.loc[states["Country"].isna(), "STATE"].unique())))

    # Transport hubs only move products between sites; keeping them would count the flows twice
    not_sites = data[(is_other | is_europe_dummy) & ~loc.str.contains("TRANSPORT_HUB", regex=False)].copy()
    not_sites["LOCATION"] = not_sites["LOCATION"].replace({"Europe_000": "Europe"})

    country = (states.groupby(dims + ["Country"], dropna=False, as_index=False)["VALUE"].sum()
               .rename(columns={"Country": "LOCATION"}))
    europe = states.groupby(dims, dropna=False, as_index=False)["VALUE"].sum().assign(LOCATION="Europe")

    return pd.concat([country, not_sites], ignore_index=True), pd.concat([europe, not_sites], ignore_index=True)


model_data = {}
for variable in NEEDED_VARIABLES:
    print(f"Reading {variable}")
    model_data[variable] = aggregate_locations(read_model_variable(variable))


def C(variable):
    """Country level (LOCATION = country)"""
    return model_data[variable][0]


def E(variable):
    """Europe level (LOCATION = "Europe")"""
    return model_data[variable][1]


# 4. Helpers for the wide format ######

def to_wide(data, ids, sheet):
    """One row per `ids`, one column per year. Duplicate rows are summed; the sum is rounded to the
    digits of the sheet."""
    if data.empty:
        return pd.DataFrame(columns=ids)

    keys = ids if ids else ["_all"]
    data = data.assign(_all="")
    wide = (data.groupby(keys + ["YEAR"])["value"].sum().round(digits[sheet])
            .unstack("YEAR").reset_index())
    wide.columns = [c if isinstance(c, str) else str(int(c)) for c in wide.columns]
    wide.columns.name = None
    return wide if ids else wide.drop(columns="_all")


def finish_sheet(wide, columns, years, sheet, sort_by=()):
    """Fixes the column order (`columns` + years), adds missing years and replaces missing values by 0"""
    year_columns = [str(y) for y in years]
    for y in year_columns:
        if y not in wide.columns:
            wide[y] = 0.0
    wide[year_columns] = wide[year_columns].fillna(0)

    wide = wide[columns + year_columns]
    if sort_by:
        wide = wide.sort_values(list(sort_by), kind="stable")
    if cfg["drop_all_zero_rows"]:
        wide = wide[(wide[year_columns] != 0).any(axis=1)]
    if wide.empty:
        print(f"  Note: sheet '{sheet}' has no data in this scenario")
    return wide.reset_index(drop=True)


def rounded(values, sheet):
    return np.round(values, digits[sheet])


sheets = {}

# 5. Production volumes ######

# Europe totals only use the steel sites (LOCATION == "Europe"): production at transport hubs,
# the energy center and the demand location is not part of the Europe total.

d = E("LocalProduction")
d = d[(d["LOCATION"] == "Europe") & d["PRODUCT"].isin(LS_products)]
sheets["total_steel_production"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"] / cap2act, "total_steel_production")).rename(columns={"PRODUCT": "Product"}),
            ["Product"], "total_steel_production").assign(Sector="Steel", Region="Europe", Unit="Mt"),
    ["Sector", "Region", "Product", "Unit"], years_all, "total_steel_production", ["Product"])

d = C("LocalProduction")
d = d[d["PRODUCT"].isin(LS_products)]
sheets["country_steel_production"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"] / cap2act, "country_steel_production"))
            .rename(columns={"LOCATION": "Country", "PRODUCT": "Product"}),
            ["Country", "Product"], "country_steel_production").assign(Sector="Steel", Unit="Mt"),
    ["Sector", "Country", "Product", "Unit"], years_all, "country_steel_production", ["Country", "Product"])

# Steel production with all scrap-based routes summed to "liquid_steel_scrap-agg"
d = C("LocalProduction")
d = d[d["PRODUCT"].isin(LS_products)]
country_steel_agg = (
    d.assign(value=np.round(d["VALUE"] / cap2act, 2),
             Product=np.where(d["PRODUCT"].str.contains("scrap", regex=False), "liquid_steel_scrap-agg", d["PRODUCT"]))
    .groupby(["YEAR", "LOCATION", "Product"], as_index=False)["value"].sum())

sheets["total_steel_production_agg"] = finish_sheet(
    to_wide(country_steel_agg.groupby(["YEAR", "Product"], as_index=False)["value"].sum(),
            ["Product"], "total_steel_production_agg").assign(Sector="Steel", Country="Europe", Unit="Mt"),
    ["Sector", "Country", "Product", "Unit"], years_all, "total_steel_production_agg", ["Product"])

sheets["country_steel_production_agg"] = finish_sheet(
    to_wide(country_steel_agg.rename(columns={"LOCATION": "Country"}),
            ["Country", "Product"], "country_steel_production_agg").assign(Sector="Steel", Unit="Mt"),
    ["Sector", "Country", "Product", "Unit"], years_all, "country_steel_production_agg", ["Country", "Product"])

d = E("LocalProduction")
d = d[(d["LOCATION"] == "Europe") & d["PRODUCT"].isin(iron_products)]
sheets["total_iron_production"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"] / cap2act, "total_iron_production")).rename(columns={"PRODUCT": "Product"}),
            ["Product"], "total_iron_production").assign(Sector="Steel", Region="Europe", Unit="Mt"),
    ["Sector", "Region", "Product", "Unit"], years_all, "total_iron_production", ["Product"])

d = C("LocalProduction")
d = d[d["PRODUCT"].isin(iron_products)]
sheets["country_iron_production"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"] / cap2act, "country_iron_production"))
            .rename(columns={"LOCATION": "Country", "PRODUCT": "Product"}),
            ["Country", "Product"], "country_iron_production").assign(Sector="Steel", Unit="Mt"),
    ["Sector", "Country", "Product", "Unit"], years_all, "country_iron_production", ["Country", "Product"])


# 6. Energy ######

def per_country(data, keep, factor=1.0):
    """Sum per year / country / product of the rows selected by `keep`, divided by cap2act"""
    data = data[keep(data)]
    data = data.assign(value=data["VALUE"] * factor / cap2act)
    return data.groupby(["YEAR", "LOCATION", "PRODUCT"], as_index=False)["value"].sum()


def not_energy_center(data):
    return data["LOCATION"] != "energy_center_location"


# Coke: modelled in Mt, converted to PJ. Only the coke that is not produced in the country's own
# coke ovens counts as primary energy (coke import). Coking coal and the other primary energy
# products are already modelled in energy units.
def coke_rows(data):
    return ((data["PRODUCT"] == "coke") & not_energy_center(data)
            & ~data["LOCATION"].str.contains("TRANSPORT", regex=False))


coke = (per_country(C("LocalUseByTechnology"), coke_rows, coking_coal_energy_factor)
        .merge(per_country(C("LocalProductionByTechnology"), coke_rows, coking_coal_energy_factor),
               on=["YEAR", "LOCATION", "PRODUCT"], how="left", suffixes=("_use", "_production")))
coke_import_unrounded = coke.assign(
    value=np.maximum(coke["value_use"] - coke["value_production"].fillna(0), 0))[["YEAR", "LOCATION", "PRODUCT", "value"]]
country_coke_import = coke_import_unrounded.assign(value=np.round(coke_import_unrounded["value"],
                                                                  digits["country_energy_consumption"]))

# Electricity: consumption minus the autoproduction from off-gases is taken from the grid
electricity_use = per_country(C("LocalUseByTechnology"), lambda x: x["PRODUCT"] == "electricity")
electricity_production = per_country(C("LocalProductionByTechnology"),
                                     lambda x: (x["PRODUCT"] == "electricity") & not_energy_center(x))
electricity = electricity_use.merge(electricity_production, on=["YEAR", "LOCATION", "PRODUCT"], how="left",
                                    suffixes=("_use", "_production"))
electricity_from_grid_unrounded = electricity.assign(
    value=electricity["value_use"] - electricity["value_production"].fillna(0))[["YEAR", "LOCATION", "PRODUCT", "value"]]
country_electricity_from_grid = electricity_from_grid_unrounded.assign(
    value=np.round(electricity_from_grid_unrounded["value"], digits["country_energy_consumption"]))

non_electric_primary_energy = [p for p in primary_energy_products if "electricity" not in p]

d = E("LocalUseByTechnology")
d = d[d["PRODUCT"].isin(non_electric_primary_energy)]
total_energy = pd.concat([
    d.assign(value=rounded(d["VALUE"] / cap2act, "total_energy_consumption"))
    .groupby(["YEAR", "PRODUCT"], as_index=False)["value"].sum(),
    pd.concat([country_coke_import, country_electricity_from_grid]).groupby(["YEAR", "PRODUCT"], as_index=False)["value"].sum(),
])
sheets["total_energy_consumption"] = finish_sheet(
    to_wide(total_energy.rename(columns={"PRODUCT": "Product"}), ["Product"], "total_energy_consumption")
    .assign(Sector="Steel", Variable="Primary energy consumption", Unit="PJ/a", Region="Europe"),
    ["Sector", "Region", "Variable", "Product", "Unit"], years_all, "total_energy_consumption", ["Product"])

d = C("LocalUseByTechnology")
d = d[d["PRODUCT"].isin(non_electric_primary_energy)]
country_energy = pd.concat([
    d.assign(value=rounded(d["VALUE"] / cap2act, "country_energy_consumption"))
    .groupby(["YEAR", "LOCATION", "PRODUCT"], as_index=False)["value"].sum(),
    country_coke_import, country_electricity_from_grid,
])
sheets["country_energy_consumption"] = finish_sheet(
    to_wide(country_energy.rename(columns={"LOCATION": "Country", "PRODUCT": "Product"}),
            ["Country", "Product"], "country_energy_consumption")
    .assign(Sector="Steel", Variable="Primary Energy consumption", Unit="PJ/a"),
    ["Sector", "Country", "Variable", "Product", "Unit"], years_all, "country_energy_consumption", ["Country", "Product"])


# 7. Capacities and investments ######

d = E("NewCapacity")
d = d[d["TECHNOLOGY"].isin(capacity_technologies)]
sheets["total_capacity_investments"] = finish_sheet(
    to_wide(d.assign(value=d["VALUE"]).rename(columns={"TECHNOLOGY": "Technology"}),
            ["Technology"], "total_capacity_investments").assign(Sector="Steel", Region="Europe", Unit="Mt/a"),
    ["Sector", "Region", "Technology", "Unit"], years_investment, "total_capacity_investments", ["Technology"])

d = E("LocalCapitalInvestment")
d = d[d["TECHNOLOGY"].isin(investment_technologies)]
sheets["total_capital_investments"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"], "total_capital_investments")).rename(columns={"TECHNOLOGY": "Technology"}),
            ["Technology"], "total_capital_investments")
    .assign(Sector="Steel", Region="Europe", Variable="Capital Investment", Unit="M€"),
    ["Sector", "Region", "Variable", "Technology", "Unit"], years_investment, "total_capital_investments", ["Technology"])

# (the column is called "Region" in the R code although it contains countries)
d = C("LocalCapitalInvestment")
d = d[d["TECHNOLOGY"].isin(investment_technologies)]
sheets["country_capital_investment"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"], "country_capital_investment"))
            .rename(columns={"LOCATION": "Region", "TECHNOLOGY": "Technology"}),
            ["Region", "Technology"], "country_capital_investment")
    .assign(Sector="Steel", Variable="Capital Investment", Unit="M€"),
    ["Sector", "Region", "Variable", "Technology", "Unit"], years_investment, "country_capital_investment",
    ["Region", "Technology"])

d = E("TotalCapacity")
d = d[d["TECHNOLOGY"].isin(capacity_technologies)]
sheets["total_installed_capacities"] = finish_sheet(
    to_wide(d.assign(value=d["VALUE"]).rename(columns={"TECHNOLOGY": "Technology"}),
            ["Technology"], "total_installed_capacities").assign(Sector="Steel", Region="Europe", Unit="Mt/a"),
    ["Sector", "Region", "Technology", "Unit"], years_all, "total_installed_capacities", ["Technology"])

d = C("LocalTotalCapacity")
d = d[d["TECHNOLOGY"].isin(capacity_technologies)]
sheets["country_installed_capacities"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"], "country_installed_capacities"))
            .rename(columns={"LOCATION": "Country", "TECHNOLOGY": "Technology"}),
            ["Country", "Technology"], "country_installed_capacities").assign(Sector="Steel", Unit="Mt/a"),
    ["Sector", "Country", "Technology", "Unit"], years_all, "country_installed_capacities", ["Country", "Technology"])


# 8. Emissions ######

d = E("LocalTechnologyEmission")
d = d[d["LOCATION"] == "Europe"]
sheets["total_emissions"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"] / cap2act, "total_emissions")), [], "total_emissions")
    .assign(Sector="Steel", Region="Europe", Variable="Annual Emissions", Unit="Mt_CO2/a"),
    ["Sector", "Region", "Variable", "Unit"], years_all, "total_emissions")

d = E("LocalProductionByMode")
d = d[d["TECHNOLOGY"].isin(cfg["captured_emissions_technologies"]) & (d["LOCATION"] == "Europe")]
sheets["captured_emissions_by_route"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"] / cap2act, "captured_emissions_by_route"))
            .rename(columns={"MODE_OF_OPERATION": "Route"}), ["Route"], "captured_emissions_by_route")
    .assign(Sector="Steel", Region="Europe", Variable="Captured Emissions", Unit="Mt_CO2/a"),
    ["Sector", "Region", "Variable", "Route", "Unit"], years_ccs, "captured_emissions_by_route", ["Route"])

# CO2 storage is a region-level variable (no location)
co2_clusters = pd.DataFrame({"PRODUCT": list(cfg["co2_storage_clusters"]),
                             "Cluster": list(cfg["co2_storage_clusters"].values())})
d = E("Use").merge(co2_clusters, on="PRODUCT")
sheets["captured_emissions_by_cluster"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"] / cap2act, "captured_emissions_by_cluster")),
            ["Cluster"], "captured_emissions_by_cluster")
    .assign(Sector="Steel", Region="Europe", Variable="Stored Emissions", Unit="Mt_CO2/a"),
    ["Sector", "Region", "Variable", "Cluster", "Unit"], years_ccs, "captured_emissions_by_cluster", ["Cluster"])


# 9. Materials and scrap ######

d = E("LocalUse")
d = d[(d["LOCATION"] == "Europe") & d["PRODUCT"].isin(cfg["material_products"])]
sheets["total_material_consumption"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"] / cap2act, "total_material_consumption"))
            .rename(columns={"LOCATION": "Region", "PRODUCT": "Product"}),
            ["Region", "Product"], "total_material_consumption")
    .assign(Sector="Steel", Variable="Material consumption", Unit="Mt/a"),
    ["Sector", "Region", "Variable", "Product", "Unit"], years_all, "total_material_consumption", ["Product"])

d = C("LocalUse")
d = d[d["PRODUCT"].isin(cfg["material_products"])]
sheets["country_material_consumption"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"] / cap2act, "country_material_consumption"))
            .rename(columns={"LOCATION": "Country", "PRODUCT": "Product"}),
            ["Country", "Product"], "country_material_consumption")
    .assign(Sector="Steel", Variable="Material consumption", Unit="Mt/a"),
    ["Sector", "Country", "Variable", "Product", "Unit"], years_all, "country_material_consumption", ["Country", "Product"])


def country_use_aggregate(products, name):
    """Sum of several products per country and year, as one product"""
    d = C("LocalUse")
    return (d[d["PRODUCT"].isin(products)].groupby(["LOCATION", "YEAR"], as_index=False)["VALUE"].sum()
            .assign(PRODUCT=name))


# Scrap grades and intra-EU HBI summed to one product each
d = C("LocalUse")
d = d[d["PRODUCT"].isin([p for p in cfg["material_products"] if p not in cfg["scrap_products"]])]
d = pd.concat([d[["YEAR", "PRODUCT", "LOCATION", "VALUE"]],
               country_use_aggregate(cfg["scrap_products"], "scrap_aggregate"),
               country_use_aggregate(cfg["hbi_intra_eu_products"], "HBI_intra-EU")])
sheets["country_mat_consumption_agg"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"] / cap2act, "country_mat_consumption_agg"))
            .rename(columns={"LOCATION": "Country", "PRODUCT": "Product"}),
            ["Country", "Product"], "country_mat_consumption_agg")
    .assign(Sector="Steel", Variable="Material consumption", Unit="Mt/a"),
    ["Sector", "Country", "Variable", "Product", "Unit"], years_all, "country_mat_consumption_agg", ["Country", "Product"])

d = C("LocalProduction")
d = d[d["PRODUCT"].isin(cfg["slag_products"])]
sheets["country_slag_production"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"] / cap2act, "country_slag_production"))
            .rename(columns={"LOCATION": "Country"}), ["Country"], "country_slag_production")
    .assign(Sector="Steel", Variable="LocalProduction", Product="bf/saf-slag"),
    ["Sector", "Country", "Variable", "Product"], years_all, "country_slag_production", ["Country"])


# 10. Costs ######

d = E("TotalDiscountedCost")
sheets["total_discounted_cost"] = finish_sheet(
    to_wide(d.assign(value=rounded(d["VALUE"], "total_discounted_cost")), [], "total_discounted_cost")
    .assign(Sector="Steel", Region="Europe", Variable="TotalDiscountedCost", Unit="M€"),
    ["Sector", "Region", "Variable", "Unit"], years_all, "total_discounted_cost")

# Marginal production cost: undiscounted shadow price of the product balance (PB9_ProductBalance)
# of the final demand products, only where the discounted shadow price is positive.
# shadow_prices.csv is large (~0.5 GB), therefore it is read in chunks.
print("Reading shadow_prices (large file)")
chunks = pd.read_csv(model_file("shadow_prices"), chunksize=500_000,
                     usecols=["c.name", "index", "shadow_price", "undiscounted_shadow_price"])
shadow_prices = pd.concat(chunk[chunk["c.name"] == "PB9_ProductBalance"] for chunk in chunks)

pb9 = shadow_prices.join(shadow_prices["index"].str.split(";", n=2, expand=True)
                         .set_axis(["Region", "Product", "Year"], axis=1))
pb9 = pb9[pb9["Product"].str.contains("Final_", regex=False) & (pb9["shadow_price"] > 0)
          & pb9["Product"].isin(cfg["marginal_cost_products"])]
pb9 = pb9.assign(YEAR=pb9["Year"].astype(float).astype(int),
                 value=rounded(pb9["undiscounted_shadow_price"], "marginal_production_cost"))
sheets["marginal_production_cost"] = finish_sheet(
    to_wide(pb9, ["Region", "Product"], "marginal_production_cost").assign(Sector="Steel", Unit="€/t"),
    ["Sector", "Region", "Product", "Unit"], years_all, "marginal_production_cost", ["Region", "Product"])


# 11. IAMC format ######

def build_iamc():
    """Production (iron and steel in one table), final energy and material consumption per country in the IAMC-like layout:
    Model, Scenario, Region, Name, Unit, <years>, one table per variable.
    Name = Variable | Sector | detail, Region = country"""
    iamc = cfg["iamc"]
    sector = " | ".join(iamc["sector"])
    countries = set(region_grouping["Country"]) - {"Europe"}

    def country_rows(data, label):
        """Only real countries; flows at Europe_000, the energy center or the demand locations are left out"""
        left_out = data[~data["LOCATION"].isin(countries)]
        if abs(left_out["value"]).sum() > 1e-9:
            print(f"  NOTE ({label}): not a country, not in the IAMC table: "
                  + ", ".join(f"{loc} ({v:.2f})" for loc, v in left_out.groupby("LOCATION")["value"].sum().items()
                              if abs(v) > 1e-9) + " (sum over all years)")
        return data[data["LOCATION"].isin(countries)]

    def map_products(data, mapping, label):
        """detail = mapping[PRODUCT] or, for products mapped by mode of operation, mapping[PRODUCT|MODE]"""
        data = data[data["value"] != 0]
        by_mode = {key.split("|")[0] for key in mapping if "|" in key}
        key = np.where(data["PRODUCT"].isin(by_mode), data["PRODUCT"] + "|" + data.get("MODE_OF_OPERATION", ""),
                       data["PRODUCT"])
        if unmapped := sorted(set(key) - set(mapping)):
            sys.exit(f"iamc.{label}_mapping in config_steel.yaml has no entry for: {', '.join(unmapped)}")
        return data.assign(detail=pd.Series(key, index=data.index).map(mapping))

    # Production, model flows divided by cap2act. Steel: LocalProduction. Iron: LocalProductionByMode,
    # because the energy carrier of the DRI routes is only part of the mode of operation.
    production_mapping = {}
    for product, entry in iamc["production_mapping"].items():
        routes = entry.get("route_by_mode")
        if routes:
            production_mapping.update({f"{product}|{mode}": f"{entry['subsector']} | {route}"
                                       for mode, route in routes.items()})
        else:
            production_mapping[product] = f"{entry['subsector']} | {entry['route']}"

    steel = C("LocalProduction")
    steel = steel[steel["PRODUCT"].isin(LS_products)]
    iron = C("LocalProductionByMode")
    iron = iron[iron["PRODUCT"].isin(iron_products)]
    production = pd.concat([steel, iron]).assign(value=lambda x: x["VALUE"] / cap2act)
    production = map_products(country_rows(production, "production"), production_mapping, "production")

    # Final energy: the figures of the sheet country_energy_consumption, without rounding
    d = C("LocalUseByTechnology")
    d = d[d["PRODUCT"].isin(non_electric_primary_energy)]
    d = d.assign(value=d["VALUE"] / cap2act)[["YEAR", "LOCATION", "PRODUCT", "value"]]
    energy = pd.concat([d, coke_import_unrounded, electricity_from_grid_unrounded])
    energy = map_products(country_rows(energy, "final energy"), iamc["energy_carrier_mapping"], "energy_carrier")

    # Material consumption: the figures of the sheet country_mat_consumption_agg (scrap grades and
    # intra-EU HBI summed to one material each), without rounding
    material_of = {p: p for p in cfg["material_products"] if p not in cfg["scrap_products"]}
    material_of.update({p: "scrap_aggregate" for p in cfg["scrap_products"]})
    material_of.update({p: "HBI_intra-EU" for p in cfg["hbi_intra_eu_products"]})
    d = C("LocalUse")
    d = d[d["PRODUCT"].isin(material_of)]
    material = (d.assign(PRODUCT=d["PRODUCT"].map(material_of), value=d["VALUE"] / cap2act)
                .groupby(["YEAR", "LOCATION", "PRODUCT"], as_index=False)["value"].sum())
    material = map_products(country_rows(material, "material consumption"), iamc["material_mapping"], "material")

    tables = {}
    for variable, data, unit in [("Production", production, iamc["units"]["production"]),
                                 ("Final Energy", energy, iamc["units"]["final_energy"]),
                                 ("Material Consumption", material, iamc["units"]["material_consumption"])]:
        data = data.assign(Name=variable + " | " + sector + " | " + data["detail"]).rename(columns={"LOCATION": "Region"})
        wide = (data.groupby(["Region", "Name", "YEAR"])["value"].sum().unstack("YEAR")
                .reindex(columns=years_all).fillna(0.0).round(iamc["decimals"]))
        if iamc["drop_all_zero_rows"]:
            wide = wide[(wide != 0).any(axis=1)]
        table = wide.reset_index().assign(Model=iamc["model"], Scenario=scenario, Unit=unit)
        tables[variable] = table[["Model", "Scenario", "Region", "Name", "Unit"] + years_all]
    return tables


# 12. Writing the result ######

def format_sheet(ws):
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for column in ws.columns:
        width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column[:200])
        ws.column_dimensions[column[0].column_letter].width = min(width + 2, 60)


output_dir.mkdir(parents=True, exist_ok=True)

if "indicator_sheets" in output_formats:
    assert set(SHEET_ORDER) == set(sheets), "sheet list does not match"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for name in SHEET_ORDER:
            sheets[name].to_excel(writer, sheet_name=name, index=False)
            format_sheet(writer.sheets[name])
    print(f"Written: {output_file}")

if "iamc" in output_formats:
    print("Building IAMC table")
    iamc_tables = build_iamc()
    with pd.ExcelWriter(iamc_file, engine="openpyxl") as writer:
        for variable, table in iamc_tables.items():
            table.to_excel(writer, sheet_name=variable, index=False)
            format_sheet(writer.sheets[variable])
    print(f"Written: {iamc_file}")
