# Sprint 3: Screener + Peer Engine (Nifty 100 Universe)

Automated multi-factor financial screener, sector-adjusted percentile ranking engine, and comparative analytics suite for the Nifty 100 universe.

---

## 1. Directory Structure

```text
project3/
├── config/
│   └── screener_config.yaml    # Analyst thresholds & column mappings
├── src/
│   ├── screener/
│   │   ├── __init__.py
│   │   └── engine.py           # Core filtering, scoring & Excel styling
│   └── analytics/
│       ├── __init__.py
│       ├── peer.py             # Percentiles, SQLite export & report generation
│       └── radar.py            # Matplotlib 8-axis radar visualization
├── tests/
│   └── test_data_quality.py    # Automated DQ pytest unit tests
├── output/
│   ├── screener_output.xlsx    # 6 preset screener sheets
│   └── peer_comparison.xlsx    # 11 peer group comparison sheets
├── reports/
│   └── radar_charts/           # PNG radar plots for each company
├── financials.db               # SQLite database (peer_percentiles table)
├── run_sprint3.py              # End-to-end execution pipeline
├── requirements.txt            # Python dependencies
└── README.md                   # Documentation