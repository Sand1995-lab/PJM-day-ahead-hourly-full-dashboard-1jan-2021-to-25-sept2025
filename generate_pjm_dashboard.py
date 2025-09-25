#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate a unified HTML dashboard for PJM Day-Ahead Zone Prices.

Usage:
  python generate_pjm_dashboard.py \
      --input "PJM_DayAhead_Prices_ZoneWise (1).xlsx" \
      --sheet ZoneWisePrices \
      --output PJM_Unified_Dashboard.html \
      --assets-dir assets \
      --include-plotlyjs cdn
"""
import argparse
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.offline import plot


# ---------- Helpers ----------
def load_data(xlsx_path: Path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    if "Date" not in df.columns:
        raise ValueError("Input sheet must contain a 'Date' column.")
    df["Date"] = pd.to_datetime(df["Date"])
    df = df.sort_values("Date").reset_index(drop=True)
    return df


def season_of_month(m: int) -> str:
    return {
        12: "Winter", 1: "Winter", 2: "Winter",
        3: "Spring", 4: "Spring", 5: "Spring",
        6: "Summer", 7: "Summer", 8: "Summer",
        9: "Autumn", 10: "Autumn", 11: "Autumn",
    }[m]


def fig_div(fig: go.Figure, include_js: str = "cdn") -> str:
    """
    include_js: 'cdn' (default) | True (embed JS) | False (no JS)
    """
    return plot(fig, include_plotlyjs=include_js, output_type="div")


# ---------- Core builder ----------
def build_dashboard(
    df: pd.DataFrame,
    output_html: Path,
    assets_dir: Path,
    include_plotlyjs: str = "cdn",
):
    assets_dir.mkdir(parents=True, exist_ok=True)

    zone_cols = [c for c in df.columns if c != "Date"]

    # Long form
    stacked = df.set_index("Date")[zone_cols].stack().reset_index()
    stacked.columns = ["Date", "Zone", "Price"]

    # Monthly long
    long_df = stacked.copy()
    long_df["Year"] = long_df["Date"].dt.year
    long_df["Month"] = long_df["Date"].dt.month
    monthly = long_df.groupby(["Zone", "Year", "Month"], as_index=False)["Price"].mean()
    monthly["YearMonth"] = pd.to_datetime(dict(year=monthly["Year"], month=monthly["Month"], day=1))
    monthly["YearMonthLabel"] = monthly["YearMonth"].dt.strftime("%Y-%m")

    # Save monthly CSV
    monthly_csv = assets_dir / "zone_year_month_avg_prices.csv"
    monthly.sort_values(["Zone", "Year", "Month"]).to_csv(monthly_csv, index=False)

    # KPIs & summaries
    overall = {
        "Start": df["Date"].min().date().isoformat(),
        "End": df["Date"].max().date().isoformat(),
        "Hours": len(df),
        "Zones": len(zone_cols),
        "Mean": stacked["Price"].mean(),
        "Median": stacked["Price"].median(),
        "Min": stacked["Price"].min(),
        "Max": stacked["Price"].max(),
        "P95": stacked["Price"].quantile(0.95),
        "Stdev": stacked["Price"].std(),
    }

    zone_summary = (
        stacked.groupby("Zone")
        .agg(
            Mean=("Price", "mean"),
            Median=("Price", "median"),
            Min=("Price", "min"),
            Max=("Price", "max"),
            P95=("Price", lambda s: s.quantile(0.95)),
            Stdev=("Price", "std"),
        )
        .sort_values("Mean", ascending=False)
        .reset_index()
    )
    zone_summary["Rank"] = zone_summary["Mean"].rank(method="first", ascending=False).astype(int)

    # ---- Figures: Core ----
    # Average across zones
    avg_series = df.set_index("Date")[zone_cols].mean(axis=1).rename("Average Price ($/MWh)").reset_index()
    fig_avg = px.line(avg_series, x="Date", y="Average Price ($/MWh)", title="PJM Day-Ahead — Average Across Zones")
    fig_avg.update_layout(xaxis=dict(rangeslider=dict(visible=True)), template="plotly_white",
                          margin=dict(l=10, r=10, t=60, b=10))

    # Top 10 zones by mean
    top10 = zone_summary.head(10)["Zone"].tolist()
    fig_top = go.Figure()
    for z in top10:
        fig_top.add_trace(go.Scatter(x=df["Date"], y=df[z], name=z, mode="lines"))
    fig_top.update_layout(
        title="Top 10 Zones by Mean Price — Trend",
        xaxis_title="Date", yaxis_title="$/MWh",
        xaxis=dict(rangeslider=dict(visible=True)),
        template="plotly_white", legend_title="Zone",
        margin=dict(l=10, r=10, t=60, b=10),
    )

    # Heatmap with dropdown (Hour x DOW)
    def heat_df_for(zone):
        t = df[["Date", zone]].copy()
        t["Hour"] = t["Date"].dt.hour
        t["DOW"] = t["Date"].dt.day_name()
        order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        t["DOW"] = pd.Categorical(t["DOW"], categories=order, ordered=True)
        g = t.groupby(["DOW", "Hour"])[zone].mean().reset_index().sort_values(["DOW", "Hour"])
        return g

    default_heat_zone = "PJM-RTO ZONE" if "PJM-RTO ZONE" in zone_cols else zone_cols[0]
    g0 = heat_df_for(default_heat_zone)
    fig_heat = px.density_heatmap(
        g0, x="Hour", y="DOW", z=default_heat_zone, histfunc="avg", nbinsx=24, nbinsy=7,
        color_continuous_scale="Viridis",
        title=f"Hourly × Day-of-Week Heatmap — {default_heat_zone}",
    )
    buttons = []
    for z in zone_cols:
        g = heat_df_for(z)
        buttons.append(dict(method="restyle", label=z, args=[{"z": [g[z]], "x": [g["Hour"]], "y": [g["DOW"]]}, [0]]))
    fig_heat.update_layout(
        template="plotly_white", margin=dict(l=10, r=10, t=60, b=10),
        updatemenus=[dict(buttons=buttons, direction="down", x=1.02, xanchor="left", y=1, yanchor="top")],
    )

    # Monthly distribution (box) for top zones
    df_m = df.copy()
    df_m["Month"] = df_m["Date"].dt.to_period("M").dt.to_timestamp()
    month_box_vals = df_m.groupby("Month")[top10].mean().reset_index()
    long_box = month_box_vals.melt(id_vars="Month", var_name="Zone", value_name="Price")
    fig_box = px.box(long_box, x="Zone", y="Price", title="Monthly Average Price Distribution — Top 10 Zones")
    fig_box.update_layout(template="plotly_white", margin=dict(l=10, r=10, t=60, b=10))

    # Sparklines (first 12 zones by mean)
    spark_zones = zone_summary.sort_values("Rank").head(12)["Zone"].tolist()
    spark_df = df[["Date"] + spark_zones].copy().melt(id_vars="Date", var_name="Zone", value_name="Price")
    fig_sparks = px.line(spark_df, x="Date", y="Price", facet_col="Zone", facet_col_wrap=4, height=700)
    fig_sparks.update_layout(title="Sparklines — First 12 Zones by Mean", showlegend=False,
                             template="plotly_white", margin=dict(l=10, r=10, t=60, b=10))
    for a in fig_sparks.layout.annotations:
        a.text = a.text.replace("Zone=", "")

    # ---- Figures: Advanced ----
    # Rolling mean & vol (30-day)
    avg2 = df.set_index("Date")[zone_cols].mean(axis=1).rename("Avg")
    roll = pd.DataFrame({"Avg": avg2})
    roll["RollMean30d"] = roll["Avg"].rolling(24 * 30, min_periods=24).mean()
    roll["RollVol30d"] = roll["Avg"].rolling(24 * 30, min_periods=24).std()
    fig_roll = go.Figure()
    fig_roll.add_trace(go.Scatter(x=roll.index, y=roll["Avg"], name="Hourly Avg", mode="lines"))
    fig_roll.add_trace(go.Scatter(x=roll.index, y=roll["RollMean30d"], name="30D Mean", mode="lines"))
    fig_roll.add_trace(go.Scatter(x=roll.index, y=roll["RollVol30d"], name="30D Volatility (σ)", mode="lines"))
    fig_roll.update_layout(
        title="Rolling 30-Day Mean & Volatility (Across All Zones)",
        xaxis_title="Date", yaxis_title="$/MWh",
        xaxis=dict(rangeslider=dict(visible=True)), template="plotly_white",
        margin=dict(l=10, r=10, t=60, b=10),
    )

    # Inter-zonal spread
    wide = df.set_index("Date")[zone_cols]
    spread = pd.DataFrame({"Date": wide.index, "Max": wide.max(axis=1), "Min": wide.min(axis=1)})
    spread["Spread"] = spread["Max"] - spread["Min"]
    fig_spread = go.Figure()
    fig_spread.add_trace(go.Scatter(x=spread["Date"], y=spread["Spread"], name="Spread", mode="lines"))
    fig_spread.update_layout(
        title="Inter-Zonal Price Spread (Max - Min by Hour)",
        xaxis_title="Date", yaxis_title="$/MWh",
        xaxis=dict(rangeslider=dict(visible=True)), template="plotly_white",
        margin=dict(l=10, r=10, t=60, b=10),
    )

    # Hour-of-day & Day-of-week
    hod = stacked.assign(Hour=lambda d: d["Date"].dt.hour).groupby("Hour")["Price"].mean().reset_index()
    dow = (
        stacked.assign(DOW=lambda d: d["Date"].dt.day_name())
        .groupby("DOW")["Price"].mean()
        .reindex(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"])
        .reset_index()
    )
    fig_hod = px.line(hod, x="Hour", y="Price", markers=True, title="Average by Hour of Day (All Zones)")
    fig_hod.update_layout(template="plotly_white", margin=dict(l=10, r=10, t=60, b=10))
    fig_dow = px.bar(dow, x="DOW", y="Price", title="Average by Day of Week (All Zones)")
    fig_dow.update_layout(template="plotly_white", margin=dict(l=10, r=10, t=60, b=10))

    # Seasonal
    seasonal = stacked.copy()
    seasonal["Season"] = seasonal["Date"].dt.month.map(season_of_month)
    seasonal = (
        seasonal.groupby("Season")["Price"]
        .agg(["mean", "median", "std", "min", "max"])
        .reindex(["Winter", "Spring", "Summer", "Autumn"])
        .reset_index()
    )
    fig_season = px.bar(seasonal, x="Season", y="mean", error_y=seasonal["std"], title="Seasonal Average Price (±σ)")
    fig_season.update_layout(yaxis_title="$/MWh", template="plotly_white", margin=dict(l=10, r=10, t=60, b=10))

    # Histogram with P95
    p95 = stacked["Price"].quantile(0.95)
    fig_hist = go.Figure()
    fig_hist.add_trace(go.Histogram(x=stacked["Price"], nbinsx=80, name="Prices"))
    fig_hist.add_shape(type="line", x0=p95, x1=p95, y0=0, y1=1, yref="paper")
    fig_hist.add_annotation(x=p95, y=1, yref="paper", text=f"P95 = {p95:.2f}", showarrow=True, arrowhead=2)
    fig_hist.update_layout(
        title="Price Distribution (All Zones, Hourly)",
        xaxis_title="$/MWh", template="plotly_white", margin=dict(l=10, r=10, t=60, b=10),
    )

    # Correlation heatmap
    corr = wide.corr().astype(float)
    fig_corr = go.Figure(
        data=go.Heatmap(z=corr.values, x=corr.columns, y=corr.index, zmin=-1, zmax=1, colorscale="RdBu")
    )
    fig_corr.update_layout(title="Correlation Between Zones", template="plotly_white", margin=dict(l=10, r=10, t=60, b=10))

    # Extremes tables (+ CSVs)
    idxmax = stacked.sort_values("Price", ascending=False).head(20)[["Date", "Zone", "Price"]]
    idxmin = stacked.sort_values("Price", ascending=True).head(20)[["Date", "Zone", "Price"]]

    idxmax_path = assets_dir / "top20_price_spikes.csv"
    idxmin_path = assets_dir / "top20_price_dips.csv"
    seasonal_path = assets_dir / "seasonal_summary.csv"
    corr_path = assets_dir / "zone_correlation.csv"
    hod_path = assets_dir / "hour_of_day_profile.csv"
    dow_path = assets_dir / "day_of_week_profile.csv"
    spread_path = assets_dir / "interzonal_spread_timeseries.csv"

    idxmax.to_csv(idxmax_path, index=False)
    idxmin.to_csv(idxmin_path, index=False)
    seasonal.to_csv(seasonal_path, index=False)
    corr.to_csv(corr_path)
    hod.to_csv(hod_path, index=False)
    dow.to_csv(dow_path, index=False)
    spread[["Date", "Spread"]].to_csv(spread_path, index=False)

    def df_to_html_table(d: pd.DataFrame) -> str:
        rows = ["<tr><th>Date</th><th>Zone</th><th>Price ($/MWh)</th></tr>"]
        for _, r in d.iterrows():
            rows.append(
                f"<tr><td>{r['Date']}</td><td>{r['Zone']}</td><td>${r['Price']:.2f}</td></tr>"
            )
        return "<table class='table'>" + "".join(rows) + "</table>"

    top_spikes_html = df_to_html_table(idxmax)
    top_dips_html = df_to_html_table(idxmin)

    # ---- Monthly Averages Section ----
    pivot = monthly.pivot_table(index="Zone", columns="YearMonthLabel", values="Price", aggfunc="mean")
    pivot = pivot.sort_index(axis=1)
    fig_heat_global = go.Figure(
        data=go.Heatmap(
            z=pivot.values, x=list(pivot.columns), y=list(pivot.index),
            colorscale="Viridis", colorbar_title="$/MWh"
        )
    )
    fig_heat_global.update_layout(
        title="Monthly Average Prices — Zone × Year-Month", xaxis_title="Year-Month", yaxis_title="Zone",
        template="plotly_white", margin=dict(l=10, r=10, t=60, b=10),
    )

    years_sorted = sorted(monthly["Year"].unique())
    months_ticks = list(range(1, 13))
    default_zone = "PJM-RTO ZONE" if "PJM-RTO ZONE" in zone_cols else zone_cols[0]

    def make_zone_season_plot(zone):
        dd = monthly[monthly["Zone"] == zone].copy()
        fig = go.Figure()
        for y in years_sorted:
            dyy = dd[dd["Year"] == y].sort_values("Month")
            fig.add_trace(go.Scatter(x=dyy["Month"], y=dyy["Price"], mode="lines+markers", name=str(y)))
        fig.update_layout(
            title=f"Monthly Average — {zone} (lines by Year)",
            xaxis_title="Month", yaxis_title="$/MWh", template="plotly_white",
            xaxis=dict(tickmode="array", tickvals=months_ticks),
        )
        return fig

    fig_zone = make_zone_season_plot(default_zone)
    zone_buttons = []
    for z in zone_cols:
        dd = monthly[monthly["Zone"] == z].copy()
        ys = []
        for y in years_sorted:
            dyy = dd[dd["Year"] == y].sort_values("Month")
            ys.append(dyy["Price"].tolist())
        traces_count = len(years_sorted)
        zone_buttons.append(
            dict(
                method="update",
                label=z,
                args=[{"y": ys, "x": [list(range(1, 13))] * traces_count}, {"title": f"Monthly Average — {z} (lines by Year)"}],
            )
        )
    fig_zone.update_layout(updatemenus=[dict(buttons=zone_buttons, direction="down", x=1.02, xanchor="left", y=1, yanchor="top")])

    def build_year_go(year: int) -> go.Figure:
        dd = monthly[monthly["Year"] == year].copy()
        zones = sorted(dd["Zone"].unique())
        fig = go.Figure()
        for z in zones:
            dz = dd[dd["Zone"] == z].sort_values("Month")
            fig.add_trace(go.Bar(name=z, x=dz["Month"], y=dz["Price"])))
        fig.update_layout(
            barmode="group",
            title=f"Year {year}: Zone-wise Monthly Average Prices",
            template="plotly_white", xaxis=dict(tickmode="array", tickvals=months_ticks),
            xaxis_title="Month", yaxis_title="$/MWh", margin=dict(l=10, r=10, t=60, b=10),
        )
        return fig

    default_year = years_sorted[-1]
    fig_year = build_year_go(default_year)
    year_menu_buttons = []
    for y in years_sorted:
        dd = monthly[monthly["Year"] == y].copy()
        zones = sorted(dd["Zone"].unique())
        xs, ys = [], []
        for z in zones:
            dz = dd[dd["Zone"] == z].sort_values("Month")
            xs.append(dz["Month"].tolist())
            ys.append(dz["Price"].tolist())
        year_menu_buttons.append(
            dict(method="update", label=str(y), args=[{"x": xs, "y": ys}, {"title": f"Year {y}: Zone-wise Monthly Average Prices"}])
        )
    fig_year.update_layout(updatemenus=[dict(buttons=year_menu_buttons, direction="down", x=1.02, xanchor="left", y=1, yanchor="top")])

    # ---------- HTML assembly ----------
    def kpi_html_block() -> str:
        return f"""
<div class='kpis'>
  <div class='kpi'><div class='label'>Data Span</div><div class='value'>{overall['Start']} → {overall['End']}</div></div>
  <div class='kpi'><div class='label'>Hours</div><div class='value'>{overall['Hours']:,}</div></div>
  <div class='kpi'><div class='label'>Zones</div><div class='value'>{overall['Zones']}</div></div>
  <div class='kpi'><div class='label'>Mean Price</div><div class='value'>${overall['Mean']:.2f}</div></div>
  <div class='kpi'><div class='label'>Volatility (σ)</div><div class='value'>{overall['Stdev']:.2f}</div></div>
</div>"""

    # Zone summary table
    rows = [
        "<tr><th>#</th><th>Zone</th><th>Mean</th><th>Median</th><th>P95</th><th>Min</th><th>Max</th><th>σ</th></tr>"
    ]
    for _, r in zone_summary.iterrows():
        rows.append(
            f"<tr>"
            f"<td>{r['Rank']}</td>"
            f"<td>{r['Zone']}</td>"
            f"<td>${r['Mean']:.2f}</td>"
            f"<td>${r['Median']:.2f}</td>"
            f"<td>${r['P95']:.2f}</td>"
            f"<td>${r['Min']:.2f}</td>"
            f"<td>${r['Max']:.2f}</td>"
            f"<td>{r['Stdev']:.2f}</td>"
            f"</tr>"
        )
    zone_table_html = "<table class='table'>" + "".join(rows) + "</table>"

    # CSS theme
    css = """
:root{
  --bg:#0b1220;
  --card:#121a2b;
  --ink:#e6eefc;
  --muted:#9fb2d8;
  --brand:#7aa2f7;
  --accent:#c099ff;
  --chip:#1f2a44;
}
*{box-sizing:border-box}
body{margin:24px;background:var(--bg);color:var(--ink);font:14px/1.5 system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif}
h1,h2,h3{color:var(--ink);margin:0 0 12px}
.section{margin:28px 0}
.grid{display:grid;gap:16px}
.grid.cols-2{grid-template-columns:1fr 1fr}
.grid.cols-3{grid-template-columns:1fr 1fr 1fr}
.card{background:var(--card);border:1px solid #19233a;border-radius:16px;padding:16px;box-shadow:0 6px 24px rgba(0,0,0,.25)}
.kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:12px}
.kpi{background:linear-gradient(180deg,#16213a,#101827);border:1px solid #1b2642;border-radius:14px;padding:14px}
.kpi .label{color:var(--muted);font-size:12px}
.kpi .value{font-size:22px;font-weight:700;margin-top:6px}
.table{width:100%;border-collapse:separate;border-spacing:0 8px}
.table th{color:#cbd5e1;text-align:left;font-weight:600;padding:8px}
.table td{padding:10px 8px;background:#0f172a;border-top:1px solid #1e293b;border-bottom:1px solid #1e293b}
.badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px;background:var(--chip);color:var(--ink)}
a{color:var(--brand);text-decoration:none}
a:hover{text-decoration:underline}
.footer{color:var(--muted);margin-top:24px;font-size:12px}
"""

    # Divs
    div_avg = fig_div(fig_avg, include_plotlyjs)
    # All the rest can omit JS because the first div includes it (when include_plotlyjs='cdn' or True)
    div_top = fig_div(fig_top, False)
    div_heat = fig_div(fig_heat, False)
    div_box = fig_div(fig_box, False)
    div_sparks = fig_div(fig_sparks, False)
    div_roll = fig_div(fig_roll, False)
    div_spread = fig_div(fig_spread, False)
    div_hod = fig_div(fig_hod, False)
    div_dow = fig_div(fig_dow, False)
    div_season = fig_div(fig_season, False)
    div_hist = fig_div(fig_hist, False)
    div_corr = fig_div(fig_corr, False)
    div_heat_global = fig_div(fig_heat_global, False)
    div_zone = fig_div(fig_zone, False)
    div_year = fig_div(fig_year, False)

    html = f"""
<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>PJM — Unified HTML Dashboard</title>
<style>{css}</style>
</head>
<body>
  <h1>PJM Day-Ahead Prices — Unified HTML Dashboard</h1>

  <div class="section card">{kpi_html_block()}</div>

  <div class="section grid cols-2">
    <div class="card">{div_avg}</div>
    <div class="card">{div_top}</div>
  </div>

  <div class="section grid cols-2">
    <div class="card">{div_heat}</div>
    <div class="card">{div_box}</div>
  </div>

  <div class="section card">{div_sparks}</div>

  <div class="section grid cols-2">
    <div class="card">{div_roll}</div>
    <div class="card">{div_spread}</div>
  </div>

  <div class="section grid cols-2">
    <div class="card">{div_hod}</div>
    <div class="card">{div_dow}</div>
  </div>

  <div class="section grid cols-2">
    <div class="card">{div_season}</div>
    <div class="card">{div_hist}</div>
  </div>

  <div class="section card">{div_corr}</div>

  <div class="section grid cols-2">
    <div class="card">
      <h2>Top Spikes (Top 20)</h2>
      <p class="badge"><a href="{idxmax_path.name}">Download CSV</a></p>
      {top_spikes_html}
    </div>
    <div class="card">
      <h2>Top Dips (Bottom 20)</h2>
      <p class="badge"><a href="{idxmin_path.name}">Download CSV</a></p>
      {top_dips_html}
    </div>
  </div>

  <div class="section card">
    <h2>Monthly Averages — All Zones, Years & Months</h2>
    <p class="badge">Download full table (CSV): <a href="{monthly_csv.name}">{monthly_csv.name}</a></p>
  </div>

  <div class="section card">
    {div_heat_global}
  </div>

  <div class="section grid cols-2">
    <div class="card">{div_zone}</div>
    <div class="card">{div_year}</div>
  </div>

  <div class="section card">
    <h2>Extra Downloads</h2>
    <ul>
      <li><a href="{seasonal_path.name}">Seasonal summary (CSV)</a></li>
      <li><a href="{corr_path.name}">Zone correlation matrix (CSV)</a></li>
      <li><a href="{hod_path.name}">Hour-of-day profile (CSV)</a></li>
      <li><a href="{dow_path.name}">Day-of-week profile (CSV)</a></li>
      <li><a href="{spread_path.name}">Inter-zonal spread timeseries (CSV)</a></li>
    </ul>
  </div>

  <div class="footer">Use the dropdowns on the heatmaps and charts to switch Zone/Year. Range sliders help focus on specific periods.</div>
</body>
</html>
"""

    output_html.write_text(html, encoding="utf-8")


def parse_args():
    p = argparse.ArgumentParser(description="Generate unified PJM HTML dashboard.")
    p.add_argument("--input", required=True, help="Path to Excel file")
    p.add_argument("--sheet", default="ZoneWisePrices", help="Sheet name (default: ZoneWisePrices)")
    p.add_argument("--output", default="PJM_Unified_Dashboard.html", help="Output HTML file")
    p.add_argument("--assets-dir", default="assets", help="Folder for CSV downloads (default: assets)")
    p.add_argument(
        "--include-plotlyjs",
        choices=["cdn", "true", "false"],
        default="cdn",
        help="Include Plotly JS: 'cdn' (default), 'true' (embed), or 'false' (no JS)",
    )
    return p.parse_args()


def main():
    args = parse_args()
    xlsx_path = Path(args.input)
    sheet_name = args.sheet
    output_html = Path(args.output)
    assets_dir = Path(args.assets_dir)

    df = load_data(xlsx_path, sheet_name)

    include_plotlyjs = {"cdn": "cdn", "true": True, "false": False}[args.include_plotlyjs]

    build_dashboard(
        df=df,
        output_html=output_html,
        assets_dir=assets_dir,
        include_plotlyjs=include_plotlyjs,
    )
    print(f"✅ Dashboard written to: {output_html}")
    print(f"📁 Downloadable CSV assets in: {assets_dir}")


if __name__ == "__main__":
    main()
