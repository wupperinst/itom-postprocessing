# ITOM-Steel postprocessing

Postprocessing of ITOM-Steel model output. The model writes long-format CSV files; this script extracts
indicators (production, energy demand, investments, emissions, materials, costs), aggregates them
spatially (location -> state -> country -> Europe) and writes them in a human-readable wide format
(one column per year) to Excel.


## Setup

Requires Python 3 with the packages in `requirements.txt` (pandas, openpyxl, pyyaml):

```
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

```
python postprocess_steel_model_output.py [path/to/config_steel.yaml] [--formats indicator_sheets iamc]
```

- Without an argument, `config_steel.yaml` next to the script is used.
- Relative paths in the config are resolved against the config file's folder.
- `--formats` overrides `output_formats` in the config. Runtime is a few seconds.

To process a different scenario, change `scenario:` in `config_steel.yaml`.

## Input

```
Model_data/
├── ITOM-Steel_region-grouping.xlsx          # sheets: Region_Grouping, Parameters
└── <scenario>/
    ├── <scenario>_LocalProduction.csv
    ├── <scenario>_LocalUse.csv
    └── ...                                  # one CSV per model variable
```

`Model_data/` is not part of the repository; place your model runs there (or point `paths:` in the
config to another location). The script reads only the files it needs: `LocalProduction`,
`LocalProductionByTechnology`, `LocalProductionByMode`, `LocalUse`, `LocalUseByTechnology`,
`LocalCapitalInvestment`, `LocalTotalCapacity`, `LocalTechnologyEmission`, `NewCapacity`,
`TotalCapacity`, `Use`, `TotalDiscountedCost` and `shadow_prices` (for the marginal production cost).

The region-grouping workbook maps states to countries and regions and gives the shares used to split
dummy locations (`<Region>_000`) to states. Its `Parameters` sheet holds the product and technology lists.

## Output

Written to `Output/` (`paths.output_dir`):

| Format | File | Content |
|---|---|---|
| `indicator_sheets` | `<scenario>_postprocessed_data.xlsx` | 24 sheets, one per indicator (`total_*`, `country_*`), in the order of the R code (plus `country_emissions` and `country_captured_emissions_by_technology`) |
| `iamc` | `<scenario>_IAMC.xlsx` | One sheet per variable: **Production**, **Final Energy**, **Material Consumption**, **Carbon Capture**, **CO2 Emissions** |

IAMC columns: `Model | Scenario | Region | Name | Unit | <years>`. Region is the country. Name has the form
`<Variable> | Industry | Iron and Steel | <detail>`, e.g. `Production | Industry | Iron and Steel | Iron | Blastfurnace`.
Values are rounded to `iamc.decimals`; all-zero rows are dropped.

Rounding in the indicator sheets is copied per sheet from the R code (`rounding_digits`) and is lossy in
places (e.g. steel production in whole Mt).

## Configuration

Everything configurable lives in `config_steel.yaml`: scenario and paths, output formats, years,
`cap2act`, technology and product lists, rounding, and the IAMC mappings (product/mode -> subsector and
route, energy product -> carrier, material -> name). If a product or mode with non-zero values has no
entry in an IAMC mapping, the script stops and names the missing entry.

## Notes on the method

- `TRANSPORT_HUB_*` locations are always dropped to avoid double counting. The energy center and the
  demand locations keep their own names.
- Steel sites are assigned to states by the first two letters of the location name; dummy locations
  are split to states with `region_share_dummy_CS_capacity`.
- Model flows (production, use, emissions) are divided by `cap2act` (5, as hard-coded in the R code).
- Europe totals for production, emissions and materials only include `LOCATION == "Europe"`
  (a deliberate deviation from the R code).

## Open points

- `cs_production_technologies` and `fs_production_technologies` are provisional and need confirmation.
- Some technologies with investment appear in no list and are missing from the capital investment sheet
  (e.g. `process_gas_heater`, blast furnace relining, `hot_briquetting`).
- The `cap2act = 5` value comes from the R code and has not been checked against the model run.

## Tests

There are no automated tests; results were checked against the R implementation and against the raw CSVs.
