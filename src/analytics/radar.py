import os
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


def generate_company_radar_charts(
    scored_df: pd.DataFrame,
    peer_df: pd.DataFrame,
    output_dir: str = "reports/radar_charts/",
):
  os.makedirs(output_dir, exist_ok=True)
  merged = scored_df.merge(peer_df, on="company_id", how="left")

  axes_cols = [
      "return_on_equity_pct",
      "roce_percentage",
      "net_profit_margin_pct",
      "debt_to_equity",
      "free_cash_flow_cr",
      "pat_cagr_5yr",
      "revenue_cagr_5yr",
      "composite_quality_score",
  ]
  available_axes = [c for c in axes_cols if c in merged.columns]
  if len(available_axes) < 3:
    return

  labels = [c.replace("_", " ").title() for c in available_axes]
  num_vars = len(labels)
  angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
  angles += angles[:1]

  for _, comp in merged.iterrows():
    cid = comp["company_id"]
    p_group = comp.get("peer_group_name")

    fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))

    if pd.isna(p_group) or not p_group:
      ref_data = (
          merged[available_axes]
          .apply(pd.to_numeric, errors="coerce")
          .mean(numeric_only=True)
      )
      ref_label = "Nifty 100 Avg"
    else:
      peer_subset = merged[merged["peer_group_name"] == p_group][
          available_axes
      ].apply(pd.to_numeric, errors="coerce")
      ref_data = peer_subset.mean(numeric_only=True)
      ref_label = f"{p_group} Avg"

    comp_vals = [
        min(max(float(comp[c]) if pd.notna(comp[c]) else 0.0, 0.0), 100.0)
        for c in available_axes
    ]
    ref_vals = [
        min(
            max(float(ref_data.get(c, 0.0)) if pd.notna(ref_data.get(c)) else 0.0, 0.0),
            100.0,
        )
        for c in available_axes
    ]

    comp_vals += comp_vals[:1]
    ref_vals += ref_vals[:1]

    name = comp.get("company_name", cid)
    ax.plot(angles, comp_vals, color="#1f77b4", linewidth=2, label=str(name))
    ax.fill(angles, comp_vals, color="#1f77b4", alpha=0.25)
    ax.plot(
        angles,
        ref_vals,
        color="#ff7f0e",
        linestyle="--",
        linewidth=1.5,
        label=ref_label,
    )

    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, size=8)
    ax.set_ylim(0, 100)

    plt.title(f"{name} ({cid})", size=11, y=1.08)
    plt.legend(loc="upper right", bbox_to_anchor=(1.35, 1.1), fontsize=8)
    plt.tight_layout()

    out_file = os.path.join(output_dir, f"{cid}_radar.png")
    plt.savefig(out_file, dpi=120, bbox_inches="tight")
    plt.close()