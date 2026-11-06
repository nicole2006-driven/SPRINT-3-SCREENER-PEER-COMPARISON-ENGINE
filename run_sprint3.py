import os
import pandas as pd
from src.analytics.peer import (
    compute_peer_percentiles,
    generate_peer_comparison_report,
)
from src.analytics.radar import generate_company_radar_charts
from src.screener.engine import (
    compute_sector_relative_scores,
    export_screener_output,
    load_config,
    validate_universe_presets,
)

os.makedirs("output", exist_ok=True)
os.makedirs("reports/radar_charts", exist_ok=True)

print("1. Loading raw Excel files...")
# Load ratios and pick the most recent year per company
ratios_raw = pd.read_excel("financial_ratios.xlsx")
if "year" in ratios_raw.columns:
  ratios_df = (
      ratios_raw.sort_values(by=["company_id", "year"])
      .groupby("company_id")
      .last()
      .reset_index()
  )
else:
  ratios_df = ratios_raw.copy()

# Load companies (header starts on row 1)
comp_df = pd.read_excel("companies.xlsx", header=1)
comp_df = comp_df.rename(columns={"id": "company_id"})

# Merge company_name and roce_percentage into ratios
merge_comp_cols = [
    c
    for c in ["company_id", "company_name", "roce_percentage"]
    if c in comp_df.columns
]
ratios_df = ratios_df.merge(
    comp_df[merge_comp_cols], on="company_id", how="left"
)

# Load market_cap and pick latest year
if os.path.exists("market_cap.xlsx"):
  mcap_raw = pd.read_excel("market_cap.xlsx")
  if "year" in mcap_raw.columns:
    mcap_df = (
        mcap_raw.sort_values(by=["company_id", "year"])
        .groupby("company_id")
        .last()
        .reset_index()
    )
  else:
    mcap_df = mcap_raw.copy()
  mcap_cols = [
      c
      for c in [
          "company_id",
          "pe_ratio",
          "pb_ratio",
          "dividend_yield_pct",
          "market_cap_crore",
      ]
      if c in mcap_df.columns
  ]
  ratios_df = ratios_df.merge(mcap_df[mcap_cols], on="company_id", how="left")

# Merge sectors for broad_sector
if os.path.exists("sectors.xlsx"):
  sectors_df = pd.read_excel("sectors.xlsx")
  if "broad_sector" in sectors_df.columns:
    ratios_df = ratios_df.merge(
        sectors_df[["company_id", "broad_sector"]], on="company_id", how="left"
    )

peer_df = pd.read_excel("peer_groups.xlsx")
config = load_config("config/screener_config.yaml")

print("2. Computing composite quality scores...")
scored_df = compute_sector_relative_scores(ratios_df)

print("3. Running screener presets and saving output/screener_output.xlsx...")
presets = validate_universe_presets(scored_df, config)
export_screener_output(presets, config, "output/screener_output.xlsx")

print("4. Calculating peer percentiles and writing to financials.db...")
percentiles_df = compute_peer_percentiles(
    universe_df=scored_df,
    peer_groups_path="peer_groups.xlsx",
    db_path="financials.db",
    year=2026,
)

print("5. Generating output/peer_comparison.xlsx (11 sheets)...")
generate_peer_comparison_report(
    scored_df=scored_df,
    percentiles_df=percentiles_df,
    peer_groups_path="peer_groups.xlsx",
    output_path="output/peer_comparison.xlsx",
)

print("6. Generating radar charts in reports/radar_charts/...")
generate_company_radar_charts(scored_df, peer_df, "reports/radar_charts/")

print(
    "\nExecution Complete! All deliverables generated in output/ and"
    " reports/radar_charts/."
)