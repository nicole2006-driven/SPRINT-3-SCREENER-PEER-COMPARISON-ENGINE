import os
import sqlite3
import openpyxl
import pandas as pd
import pytest


def test_deliverable_files_exist():
  assert os.path.exists('output/screener_output.xlsx')
  assert os.path.exists('output/peer_comparison.xlsx')
  assert os.path.exists('financials.db')
  assert os.path.exists('reports/radar_charts')
  radar_files = [
      f for f in os.listdir('reports/radar_charts') if f.endswith('.png')
  ]
  assert len(radar_files) > 0


def test_screener_output_sheets_and_counts():
  wb = openpyxl.load_workbook('output/screener_output.xlsx')
  expected_sheets = {
      'quality_compounder',
      'value_pick',
      'growth_accelerator',
      'dividend_champion',
      'debt_free_blue_chip',
      'turnaround_watch',
  }
  assert set(wb.sheetnames) == expected_sheets

  for sheet in wb.sheetnames:
    ws = wb[sheet]
    row_count = ws.max_row - 1
    assert (
        5 <= row_count <= 50
    ), f"Sheet '{sheet}' returned {row_count} rows (expected 5 to 50)"


def test_quality_compounder_thresholds():
  df = pd.read_excel(
      'output/screener_output.xlsx', sheet_name='quality_compounder'
  )
  top_5 = df.head(5)
  if 'return_on_equity_pct' in top_5.columns:
    assert (top_5['return_on_equity_pct'] >= 15.0).all()


def test_peer_comparison_sheet_count():
  wb = openpyxl.load_workbook('output/peer_comparison.xlsx')
  assert len(wb.sheetnames) == 11


def test_sqlite_peer_percentiles_table():
  conn = sqlite3.connect('financials.db')
  df = pd.read_sql('SELECT * FROM peer_percentiles', conn)
  conn.close()

  assert not df.empty
  required_cols = {
      'company_id',
      'peer_group_name',
      'metric',
      'value',
      'percentile_rank',
      'year',
  }
  assert required_cols.issubset(df.columns)


def test_peer_percentile_it_services_roe():
  conn = sqlite3.connect('financials.db')
  df = pd.read_sql(
      "SELECT * FROM peer_percentiles WHERE peer_group_name = 'IT Services' AND"
      " metric = 'return_on_equity_pct'",
      conn,
  )
  conn.close()

  valid_df = df.dropna(subset=['value', 'percentile_rank'])
  assert not valid_df.empty, "No ROE data found for 'IT Services'"

  max_val_id = valid_df.loc[valid_df['value'].idxmax(), 'company_id']
  max_rank_id = valid_df.loc[valid_df['percentile_rank'].idxmax(), 'company_id']

  assert (
      max_val_id == max_rank_id
  ), f"Within IT Services, highest ROE ({max_val_id}) must have highest percentile rank ({max_rank_id})"