import numpy as np
import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd
import yaml

# ---------------------------------------------------------
# DAY 15: Config & Core Filter Engine
# ---------------------------------------------------------


def load_config(config_path: str = "config/screener_config.yaml") -> dict:
  """Load analyst-editable threshold definitions and metric mappings."""
  with open(config_path, "r") as f:
    return yaml.safe_load(f)


def apply_filters(
    df: pd.DataFrame, criteria: dict, config: dict
) -> pd.DataFrame:
  """Applies metric threshold rules to the financial universe DataFrame."""
  filtered = df.copy()

  # Treat "Debt Free" label or dashes as ICR = infinity
  if "interest_coverage" in filtered.columns:
    filtered["interest_coverage"] = (
        filtered["interest_coverage"]
        .replace(["Debt Free", "DEBT FREE", "debt free", "-"], np.inf)
        .astype(float)
    )

  for rule, threshold in criteria.items():
    # 1. Handle D/E max rules
    if rule == "de_max":
      col = config["metrics"]["de_max"]
      if col in filtered.columns:
        is_financial = (
            filtered["broad_sector"].astype(str).str.strip().str.lower()
            == "financials"
            if "broad_sector" in filtered.columns
            else pd.Series(False, index=filtered.index)
        )
        if threshold <= 0.15:
          # For debt-free presets: must NOT be Financials, and D/E <= threshold
          valid_de = (~is_financial) & (
              (filtered[col] <= threshold) | (filtered[col] == 0)
          )
        else:
          # Normal screens: skip D/E restriction for Financials
          valid_de = (
              is_financial
              | (filtered[col] <= threshold)
              | (filtered[col].isna())
          )
        filtered = filtered[valid_de]

    # 2. Minimum metric threshold
    elif rule.endswith("_min") and rule in config["metrics"]:
      col = config["metrics"][rule]
      if col in filtered.columns:
        filtered = filtered[filtered[col] >= threshold]

    # 3. Maximum metric threshold
    elif rule.endswith("_max") and rule in config["metrics"]:
      col = config["metrics"][rule]
      if col in filtered.columns:
        filtered = filtered[filtered[col] <= threshold]

    # 4. Exact match rules (e.g., debt_free_blue_chip D/E = 0)
    elif rule == "de_exact":
      col = config["metrics"]["de_max"]
      if col in filtered.columns:
        filtered = filtered[filtered[col] == threshold]

  return filtered


# ---------------------------------------------------------
# DAY 16: Preset Screeners
# ---------------------------------------------------------


def run_preset(
    preset_name: str, df: pd.DataFrame, config: dict
) -> pd.DataFrame:
  """Applies predefined preset filters across the universe."""
  preset_cfg = config["presets"][preset_name]

  if preset_name == "turnaround_watch":
    c1 = (
        (df["free_cash_flow_cr"] > 0)
        if "free_cash_flow_cr" in df.columns
        else pd.Series(True, index=df.index)
    )
    c2 = (
        (df["return_on_equity_pct"] >= 14.0)
        if "return_on_equity_pct" in df.columns
        else pd.Series(True, index=df.index)
    )
    c3 = (
        (df["operating_profit_margin_pct"] >= 12.0)
        if "operating_profit_margin_pct" in df.columns
        else pd.Series(True, index=df.index)
    )
    return df[c1 & c2 & c3].copy()

  elif preset_name == "growth_accelerator":
    res = df.copy()
    if "operating_profit_margin_pct" in res.columns:
      res = res[
          res["operating_profit_margin_pct"] >= preset_cfg.get("opm_min", 20.0)
      ]
    if "return_on_equity_pct" in res.columns:
      res = res[
          res["return_on_equity_pct"] >= preset_cfg.get("roe_min", 15.0)
      ]
    if "debt_to_equity" in res.columns:
      is_financial = (
          res["broad_sector"].astype(str).str.strip().str.lower()
          == "financials"
          if "broad_sector" in res.columns
          else pd.Series(False, index=res.index)
      )
      res = res[
          is_financial
          | (res["debt_to_equity"] <= preset_cfg.get("de_max", 2.0))
      ]
    return res

  elif preset_name == "dividend_champion":
    res = apply_filters(
        df,
        {
            k: v
            for k, v in preset_cfg.items()
            if k not in ["div_payout_max", "fcf_min"]
        },
        config,
    )
    if "dividend_payout_ratio_pct" in res.columns:
      res = res[res["dividend_payout_ratio_pct"] < preset_cfg["div_payout_max"]]
    if "free_cash_flow_cr" in res.columns:
      res = res[res["free_cash_flow_cr"] > preset_cfg["fcf_min"]]
    return res

  return apply_filters(df, preset_cfg, config)


def validate_universe_presets(
    universe_df: pd.DataFrame, config: dict
) -> dict[str, pd.DataFrame]:
  """Runs and validates return counts for each preset screener."""
  results = {}
  for preset_name in config["presets"]:
    filtered_df = run_preset(preset_name, universe_df, config)
    count = len(filtered_df)
    print(f"Preset '{preset_name}': {count} companies returned.")
    results[preset_name] = filtered_df
  return results


# ---------------------------------------------------------
# DAY 17: Composite Score & Excel Output
# ---------------------------------------------------------


def compute_sector_relative_scores(df: pd.DataFrame) -> pd.DataFrame:
  """Winsorises metrics at P10/P90, normalises within broad sectors, and computes

  a composite quality score (0 to 100).
  """
  metrics_weights = {
      "return_on_equity_pct": (0.15, False),
      "roce_percentage": (0.10, False),
      "net_profit_margin_pct": (0.10, False),
      "fcf_cagr_5yr": (0.15, False),
      "cfo_to_pat": (0.10, False),
      "fcf_positive_flag": (0.05, False),
      "revenue_cagr_5yr": (0.10, False),
      "pat_cagr_5yr": (0.10, False),
      "debt_to_equity": (0.10, True),  # Inverted: lower debt is better
      "interest_coverage": (0.05, False),
  }

  scored_df = df.copy()

  def winsorize_and_scale(group: pd.DataFrame, metric: str, invert: bool):
    vals = pd.to_numeric(group[metric], errors="coerce").replace(
        [np.inf, -np.inf], np.nan
    )
    p10, p90 = vals.quantile(0.10), vals.quantile(0.90)
    clamped = vals.clip(p10, p90)
    denom = clamped.max() - clamped.min()
    if denom == 0 or np.isnan(denom):
      scaled = pd.Series(50.0, index=group.index)
    else:
      scaled = (clamped - clamped.min()) / denom * 100.0
    return 100.0 - scaled if invert else scaled

  active_weights = {}
  for metric, (weight, invert) in metrics_weights.items():
    if metric == "fcf_positive_flag":
      if "free_cash_flow_cr" in scored_df.columns:
        scored_df[f"{metric}_pts"] = (
            scored_df["free_cash_flow_cr"] > 0
        ).astype(float) * 100.0
        active_weights[f"{metric}_pts"] = weight
    elif metric in scored_df.columns:
      sector_col = (
          "broad_sector" if "broad_sector" in scored_df.columns else None
      )
      if sector_col:
        scored_df[f"{metric}_pts"] = scored_df.groupby(
            sector_col, group_keys=False
        ).apply(lambda g: winsorize_and_scale(g, metric, invert))
      else:
        scored_df[f"{metric}_pts"] = winsorize_and_scale(
            scored_df, metric, invert
        )
      scored_df[f"{metric}_pts"] = scored_df[f"{metric}_pts"].fillna(50.0)
      active_weights[f"{metric}_pts"] = weight

  total_w = sum(active_weights.values())
  if total_w > 0:
    scored_df["composite_quality_score"] = sum(
        scored_df[col] * (w / total_w) for col, w in active_weights.items()
    ).round(2)
  else:
    scored_df["composite_quality_score"] = 50.0

  return scored_df


def export_screener_output(
    preset_dict: dict[str, pd.DataFrame],
    config: dict,
    output_path: str = "output/screener_output.xlsx",
):
  """Exports presets to Excel with conditional colour fills (Green = Pass, Red = Fail)."""
  with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for preset_name, data in preset_dict.items():
      sorted_data = data.sort_values(
          by="composite_quality_score", ascending=False
      )
      sorted_data.to_excel(writer, sheet_name=preset_name[:31], index=False)

  wb = openpyxl.load_workbook(output_path)
  green_fill = PatternFill(
      start_color="D4EDDA", end_color="D4EDDA", fill_type="solid"
  )
  red_fill = PatternFill(
      start_color="F8D7DA", end_color="F8D7DA", fill_type="solid"
  )

  for sheet in wb.sheetnames:
    ws = wb[sheet]
    headers = [cell.value for cell in ws[1]]
    criteria = config["presets"].get(sheet, {})

    for row in ws.iter_rows(min_row=2):
      for col_idx, cell in enumerate(row):
        col_name = headers[col_idx]
        for rule, threshold in criteria.items():
          target_col = config["metrics"].get(rule)
          if col_name == target_col and isinstance(cell.value, (int, float)):
            if rule.endswith("_min") and cell.value >= threshold:
              cell.fill = green_fill
            elif rule.endswith("_min") and cell.value < threshold:
              cell.fill = red_fill
            elif rule.endswith("_max") and cell.value <= threshold:
              cell.fill = green_fill
            elif rule.endswith("_max") and cell.value > threshold:
              cell.fill = red_fill
  wb.save(output_path)