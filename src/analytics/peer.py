import sqlite3
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill
import pandas as pd


def compute_peer_percentiles(
    universe_df: pd.DataFrame,
    peer_groups_path: str = "peer_groups.xlsx",
    db_path: str = "financials.db",
    year: int = 2026,
) -> pd.DataFrame:
  peer_df = pd.read_excel(peer_groups_path)
  merged = universe_df.merge(peer_df, on="company_id", how="left")

  metrics_to_rank = [
      ("return_on_equity_pct", False),
      ("roce_percentage", False),
      ("net_profit_margin_pct", False),
      ("debt_to_equity", True),  # Inverted: lower D/E = higher rank
      ("free_cash_flow_cr", False),
      ("pat_cagr_5yr", False),
      ("revenue_cagr_5yr", False),
      ("eps_cagr_5yr", False),
      ("interest_coverage", False),
      ("asset_turnover", False),
  ]

  records = []
  peer_col = "peer_group_name"

  for _, row in merged.iterrows():
    cid = row["company_id"]
    peer_group = row.get(peer_col)

    if pd.isna(peer_group) or not peer_group:
      records.append({
          "company_id": cid,
          "peer_group_name": "No peer group assigned",
          "metric": None,
          "value": None,
          "percentile_rank": None,
          "year": year,
      })
      continue

    peers = merged[merged[peer_col] == peer_group]

    for metric, invert in metrics_to_rank:
      if metric not in merged.columns:
        continue

      peer_vals = pd.to_numeric(peers[metric], errors="coerce").dropna().values
      val = row[metric]

      if pd.isna(val) or len(peer_vals) <= 1:
        rank = np.nan
      else:
        rank = (
            np.sum(peer_vals < val) + 0.5 * np.sum(peer_vals == val)
        ) / len(peer_vals)
        if invert:
          rank = 1.0 - rank

      records.append({
          "company_id": cid,
          "peer_group_name": peer_group,
          "metric": metric,
          "value": float(val) if pd.notna(val) else None,
          "percentile_rank": round(float(rank) * 100, 2)
          if pd.notna(rank)
          else None,
          "year": year,
      })

  res_df = pd.DataFrame(records)

  with sqlite3.connect(db_path) as conn:
    res_df.to_sql("peer_percentiles", conn, if_exists="replace", index=False)

  return res_df


def generate_peer_comparison_report(
    scored_df: pd.DataFrame,
    percentiles_df: pd.DataFrame,
    peer_groups_path: str = "peer_groups.xlsx",
    output_path: str = "output/peer_comparison.xlsx",
):
  peer_df = pd.read_excel(peer_groups_path)
  groups = peer_df["peer_group_name"].dropna().unique()

  green = PatternFill(
      start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
  )  # >= 75th
  yellow = PatternFill(
      start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"
  )  # 25th - 75th
  red = PatternFill(
      start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
  )  # <= 25th
  gold = PatternFill(
      start_color="FFE599", end_color="FFE599", fill_type="solid"
  )  # Benchmark row

  with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for group in groups:
      group_cids = peer_df[peer_df["peer_group_name"] == group][
          "company_id"
      ].tolist()
      group_data = scored_df[scored_df["company_id"].isin(group_cids)].copy()

      valid_pcts = percentiles_df[
          (percentiles_df["company_id"].isin(group_cids))
          & (percentiles_df["metric"].notna())
      ]
      if not valid_pcts.empty:
        piv = valid_pcts.pivot(
            index="company_id", columns="metric", values="percentile_rank"
        )
        piv.columns = [f"{c}_pct_rank" for c in piv.columns]
        group_data = group_data.merge(piv, on="company_id", how="left")

      sheet_name = str(group)[:31]
      group_data.to_excel(writer, sheet_name=sheet_name, index=False)

  wb = openpyxl.load_workbook(output_path)
  for group in groups:
    sheet_name = str(group)[:31]
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]

    benchmarks = peer_df[
        (peer_df["peer_group_name"] == group)
        & (peer_df.get("is_benchmark", 0) == 1)
    ]["company_id"].tolist()

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
      cid = row[headers.index("company_id")].value
      if cid in benchmarks:
        for cell in row:
          cell.fill = gold

      for idx, h in enumerate(headers):
        if (
            h
            and str(h).endswith("_pct_rank")
            and isinstance(row[idx].value, (int, float))
        ):
          v = row[idx].value
          if v >= 75:
            row[idx].fill = green
          elif v <= 25:
            row[idx].fill = red
          else:
            row[idx].fill = yellow

    m_row = ws.max_row + 1
    ws.cell(row=m_row, column=1, value="PEER MEDIAN").font = Font(bold=True)
    for col_idx in range(3, len(headers) + 1):
      col_letter = openpyxl.utils.get_column_letter(col_idx)
      ws.cell(
          row=m_row,
          column=col_idx,
          value=f"=MEDIAN({col_letter}2:{col_letter}{m_row - 1})",
      ).font = Font(bold=True)

  wb.save(output_path)