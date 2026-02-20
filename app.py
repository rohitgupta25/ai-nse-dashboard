from flask import Flask, render_template_string
import pandas as pd
from pathlib import Path
import importlib
import re
import os
import html as html_lib

app = Flask(__name__)

HTML = """
<html>
<head>
<title>NSE Dashboard</title>
<style>
:root {
    --bg:#071022;
    --bg-soft:#0f1e39;
    --panel:#0f1a30cc;
    --panel-strong:#132446;
    --line:#29406d;
    --text:#e6edf8;
    --muted:#9fb3d4;
    --accent:#24c8ff;
    --pos:#22c55e;
    --neg:#ef4444;
    --warn:#f59e0b;
}
* { box-sizing:border-box; }
body {
    margin:0;
    font-family:"Space Grotesk","Avenir Next","Segoe UI",sans-serif;
    color:var(--text);
    background:
        radial-gradient(1000px 500px at 10% -5%, #1f4aa544, transparent 60%),
        radial-gradient(900px 400px at 100% 0%, #0094c433, transparent 55%),
        linear-gradient(160deg, #050b18 0%, #08152b 45%, #071022 100%);
    padding:22px;
}
main { max-width: 1280px; margin: 0 auto; }
.hero {
    background:linear-gradient(135deg, #132847d9, #0e1b31d9);
    border:1px solid #2f4f84;
    border-radius:18px;
    padding:20px 24px;
    margin-bottom:18px;
    box-shadow:0 18px 40px #02081666;
}
h1 { margin:0; font-size: clamp(26px, 4vw, 40px); letter-spacing:0.2px; }
.subtitle { color:var(--muted); margin:8px 0 0; font-size:14px; }
h2 { color:var(--accent); font-size:21px; margin:0 0 14px; }
.panel {
    background:var(--panel);
    border:1px solid var(--line);
    border-radius:16px;
    padding:18px;
    margin-bottom:18px;
    backdrop-filter: blur(8px);
}
table { border-collapse: collapse; width: 100%; margin:0; font-size:13px; }
th, td { border: 1px solid #22375d; padding: 10px 8px; text-align: center; }
th { background:var(--panel-strong); color:#d9e8ff; }
tr:nth-child(even) td { background:#0d1a33aa; }
tr:nth-child(odd) td { background:#0a152baa; }

.buy { color:#22c55e; font-weight:bold; }
.sell { color:#ef4444; font-weight:bold; }
.hold { color:#eab308; font-weight:bold; }

.positive { color:#22c55e; }
.negative { color:#ef4444; }

.cards {
    display:grid;
    grid-template-columns:repeat(4, minmax(0,1fr));
    gap:12px;
    margin-bottom:18px;
}
.card {
    background:linear-gradient(155deg, #12284acd, #0d1c37cd);
    border:1px solid #35598d;
    padding:14px;
    border-radius:14px;
    min-height:90px;
}
.card .k { color:var(--muted); font-size:12px; text-transform:uppercase; letter-spacing:0.8px; }
.card .v { margin-top:8px; font-size:24px; font-weight:700; color:#f3f8ff; }
.chart-box { margin: 0; }
.chart-row { display:flex; align-items:center; gap:10px; margin:9px 0; }
.chart-label { width:170px; font-size:12px; text-align:right; color:#cbd5e1; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.chart-track { flex:1; height:14px; background:#0b1730; border:1px solid #24406f; border-radius:999px; overflow:hidden; }
.chart-bar { height:100%; border-radius:999px; }
.bar-pos { background:#22c55e; }
.bar-neg { background:#ef4444; }
.bar-neutral { background:#38bdf8; }
.chart-value { width:95px; text-align:left; font-size:12px; color:#e2e8f0; }
@media (max-width: 980px) {
    .cards { grid-template-columns:repeat(2, minmax(0,1fr)); }
}
@media (max-width: 640px) {
    body { padding:12px; }
    .cards { grid-template-columns:1fr; }
    .chart-label { width:108px; font-size:11px; }
    .chart-value { width:68px; font-size:11px; }
    table { font-size:11px; }
    th, td { padding:8px 6px; }
}
</style>
</head>
<body>
<main>

<section class="hero">
<h1>NSE Stock Dashboard</h1>
<p class="subtitle">Portfolio health, momentum signals, and fundamental overlap in one view</p>
</section>

<div class="cards">
<div class="card"><div class="k">Total Portfolio Value</div><div class="v">{{ total_value }}</div></div>
<div class="card"><div class="k">Total PnL</div><div class="v">{{ total_pnl }}</div></div>
<div class="card"><div class="k">Win Rate</div><div class="v">{{ win_rate }}%</div></div>
<div class="card"><div class="k">Top Stock</div><div class="v">{{ best_stock }}</div></div>
</div>

<section class="panel">
<h2>Portfolio PnL Chart</h2>
{{ portfolio_chart|safe }}
</section>

<section class="panel">
<h2>Common Stocks Fundamental Score Chart</h2>
{{ common_chart|safe }}
</section>

<section class="panel">
<h2>Top Gainers</h2>
{{ gainers|safe }}
</section>

<section class="panel">
<h2>Top Losers</h2>
{{ losers|safe }}
</section>

<section class="panel">
<h2>Top Momentum Stocks</h2>
{{ momentum|safe }}
</section>

<section class="panel">
<h2>Portfolio Performance</h2>
{{ portfolio|safe }}
</section>

<section class="panel">
<h2>Common Stocks (Portfolio vs Fundamentals)</h2>
{{ common|safe }}
</section>

</main>
</body>
</html>
"""

OUTPUT_DIR = Path("outputs")
OUTPUT_FILES = {
    "gainers": OUTPUT_DIR / "top_gainers.xlsx",
    "losers": OUTPUT_DIR / "top_losers.xlsx",
    "momentum": OUTPUT_DIR / "potential_stocks.xlsx",
    "portfolio": OUTPUT_DIR / "portfolio_performance.xlsx",
}


def _first_present(columns: list[str], candidates: list[str]) -> str | None:
    for candidate in candidates:
        if candidate in columns:
            return candidate
    return None


def _normalize_col(name: str) -> str:
    s = str(name).replace("\xa0", " ").strip().lower()
    s = re.sub(r"[^a-z0-9%]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _to_num(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace(r"[^0-9.\-]", "", regex=True)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def _outputs_exist() -> bool:
    return all(path.exists() for path in OUTPUT_FILES.values())


def _add_signal_column(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "signal" in out.columns:
        out["signal"] = out["signal"].astype(str).str.upper()
        return out

    score_col = None
    for col in ["final_score", "tech_score", "score"]:
        if col in out.columns:
            score_col = col
            break

    if score_col is None or out.empty:
        out["signal"] = "HOLD"
        return out

    scores = pd.to_numeric(out[score_col], errors="coerce").fillna(0.0)
    low = scores.quantile(0.33)
    high = scores.quantile(0.66)

    def classify(v: float) -> str:
        if v >= high:
            return "BUY"
        if v <= low:
            return "SELL"
        return "HOLD"

    out["signal"] = scores.apply(classify)
    return out


def _build_bar_chart(df: pd.DataFrame, label_col: str, value_col: str, signed: bool = False, max_rows: int = 12) -> str:
    if label_col not in df.columns or value_col not in df.columns or df.empty:
        return "<p>No data available for chart.</p>"

    chart_df = df[[label_col, value_col]].copy()
    chart_df[value_col] = pd.to_numeric(chart_df[value_col], errors="coerce").fillna(0.0)
    chart_df = chart_df.sort_values(value_col, key=lambda s: s.abs(), ascending=False).head(max_rows)

    if chart_df.empty:
        return "<p>No data available for chart.</p>"

    max_abs = max(float(chart_df[value_col].abs().max()), 1.0)
    rows = []
    for _, row in chart_df.iterrows():
        label = html_lib.escape(str(row[label_col]))
        value = float(row[value_col])
        width = max((abs(value) / max_abs) * 100.0, 2.0)
        if signed:
            klass = "bar-pos" if value >= 0 else "bar-neg"
        else:
            klass = "bar-neutral"
        value_txt = f"{value:.2f}%"
        rows.append(
            f"<div class='chart-row'><div class='chart-label'>{label}</div>"
            f"<div class='chart-track'><div class='chart-bar {klass}' style='width:{width:.1f}%'></div></div>"
            f"<div class='chart-value'>{value_txt}</div></div>"
        )
    return "<div class='chart-box'>" + "".join(rows) + "</div>"


def _generate_offline_outputs() -> tuple[bool, str]:
    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        fund_df = pd.read_excel("fundamentals.xlsx")
        fund_df.columns = [_normalize_col(c) for c in fund_df.columns]

        cols = fund_df.columns.tolist()
        name_col = _first_present(cols, ["name", "company", "symbol"])
        roe_col = _first_present(cols, ["roe %", "roe"])
        debt_col = _first_present(cols, ["debt eq", "debt equity", "debt"])
        sales_col = _first_present(cols, ["sales var 3yrs %", "sales growth 3yrs %", "sales growth"])
        profit_col = _first_present(cols, ["qtr profit var %", "profit growth 3years", "profit growth"])

        if name_col is None:
            return False, "missing name/company column in fundamentals.xlsx"

        work = pd.DataFrame()
        work["symbol"] = fund_df[name_col].astype(str).str.upper().str.replace(" ", "", regex=False)
        work["roe"] = _to_num(fund_df[roe_col]) if roe_col else 0.0
        work["debt"] = _to_num(fund_df[debt_col]) if debt_col else 0.0
        work["sales_growth"] = _to_num(fund_df[sales_col]) if sales_col else 0.0
        work["profit_growth"] = _to_num(fund_df[profit_col]) if profit_col else 0.0
        work = work.fillna(0.0)
        work = work[work["symbol"].str.lower() != "name"]

        work["score"] = (
            work["roe"] * 0.35
            + work["sales_growth"] * 0.25
            + work["profit_growth"] * 0.25
            - work["debt"] * 0.15
        )

        gainers = work.sort_values("score", ascending=False).head(50).rename(columns={"score": "pct_change_est"})
        losers = work.sort_values("score", ascending=True).head(50).rename(columns={"score": "pct_change_est"})

        momentum = work.copy()
        momentum["tech_score"] = (
            momentum["roe"] * 0.4
            + momentum["sales_growth"] * 0.3
            + momentum["profit_growth"] * 0.3
        )
        momentum["final_score"] = momentum["tech_score"] - momentum["debt"] * 0.2
        momentum = momentum.sort_values("final_score", ascending=False).head(50)
        momentum = _add_signal_column(momentum)

        portfolio = pd.read_excel("portfolio.xlsx")
        portfolio.columns = [_normalize_col(c) for c in portfolio.columns]
        pcols = portfolio.columns.tolist()
        symbol_col = _first_present(pcols, ["symbol", "name", "stock", "instrument"])
        entry_col = _first_present(pcols, ["entry price", "entry", "buy price", "avg cost"])
        qty_col = _first_present(pcols, ["quantity", "qty", "qty."])
        ltp_col = _first_present(pcols, ["ltp", "current", "cur val"])
        pnl_col = _first_present(pcols, ["p l", "p&l", "pnl"])

        if symbol_col is None or entry_col is None or qty_col is None:
            portfolio_df = pd.DataFrame([{
                "symbol": "N/A",
                "entry": 0.0,
                "current": 0.0,
                "pnl": 0.0,
                "pnl_pct": 0.0,
            }])
        else:
            portfolio_df = pd.DataFrame()
            portfolio_df["symbol"] = portfolio[symbol_col].astype(str).str.upper()
            portfolio_df["entry"] = _to_num(portfolio[entry_col]).fillna(0.0)
            portfolio_df["quantity"] = _to_num(portfolio[qty_col]).fillna(0.0)
            if ltp_col:
                portfolio_df["current"] = _to_num(portfolio[ltp_col]).fillna(portfolio_df["entry"])
            else:
                portfolio_df["current"] = portfolio_df["entry"]
            if pnl_col:
                portfolio_df["pnl"] = _to_num(portfolio[pnl_col]).fillna(0.0)
            else:
                portfolio_df["pnl"] = (portfolio_df["current"] - portfolio_df["entry"]) * portfolio_df["quantity"]
            portfolio_df["pnl_pct"] = (
                (portfolio_df["current"] - portfolio_df["entry"])
                .div(portfolio_df["entry"].replace(0, pd.NA))
                .fillna(0.0)
                * 100
            )

        gainers.to_excel(OUTPUT_FILES["gainers"], index=False)
        losers.to_excel(OUTPUT_FILES["losers"], index=False)
        momentum.to_excel(OUTPUT_FILES["momentum"], index=False)
        portfolio_df.to_excel(OUTPUT_FILES["portfolio"], index=False)

        return True, "offline outputs generated from local Excel files"
    except Exception as exc:
        return False, f"{exc.__class__.__name__}: {exc}"


def _load_fundamentals_scored() -> pd.DataFrame:
    fund_df = pd.read_excel("fundamentals.xlsx")
    fund_df.columns = [_normalize_col(c) for c in fund_df.columns]

    cols = fund_df.columns.tolist()
    name_col = _first_present(cols, ["name", "company", "symbol"])
    roe_col = _first_present(cols, ["roe %", "roe"])
    debt_col = _first_present(cols, ["debt eq", "debt equity", "debt"])
    sales_col = _first_present(cols, ["sales var 3yrs %", "sales growth 3yrs %", "sales growth"])
    profit_col = _first_present(cols, ["qtr profit var %", "profit growth 3years", "profit growth"])

    if name_col is None:
        return pd.DataFrame(columns=["symbol", "roe", "debt", "sales_growth", "profit_growth", "fund_score"])

    out = pd.DataFrame()
    out["symbol"] = fund_df[name_col].astype(str).str.upper().str.replace(" ", "", regex=False)
    out["roe"] = _to_num(fund_df[roe_col]) if roe_col else 0.0
    out["debt"] = _to_num(fund_df[debt_col]) if debt_col else 0.0
    out["sales_growth"] = _to_num(fund_df[sales_col]) if sales_col else 0.0
    out["profit_growth"] = _to_num(fund_df[profit_col]) if profit_col else 0.0
    out = out.fillna(0.0)
    out = out[out["symbol"].str.lower() != "name"]
    out["fund_score"] = (
        out["roe"] * 0.35
        + out["sales_growth"] * 0.25
        + out["profit_growth"] * 0.25
        - out["debt"] * 0.15
    )
    return out


@app.route("/")
def dashboard():
    refresh_ok = False
    try:
        importlib.import_module("agent_core")
        refresh_ok = True
    except Exception:
        refresh_ok = False

    if (not refresh_ok) or (not _outputs_exist()):
        _generate_offline_outputs()

    gainers_df = pd.read_excel(OUTPUT_FILES["gainers"]).head(20)
    losers_df = pd.read_excel(OUTPUT_FILES["losers"]).head(20)
    momentum_df = pd.read_excel(OUTPUT_FILES["momentum"]).head(20)
    portfolio_df = pd.read_excel(OUTPUT_FILES["portfolio"])
    fundamentals_df = _load_fundamentals_scored()

    if "current" in portfolio_df.columns:
        total_value = pd.to_numeric(portfolio_df["current"], errors="coerce").fillna(0).sum()
    else:
        total_value = 0.0
    if "pnl" in portfolio_df.columns:
        total_pnl = pd.to_numeric(portfolio_df["pnl"], errors="coerce").fillna(0).sum()
        wins = portfolio_df[pd.to_numeric(portfolio_df["pnl"], errors="coerce").fillna(0) > 0].shape[0]
    else:
        total_pnl = 0.0
        wins = 0
    total = portfolio_df.shape[0]
    win_rate = round((wins / total) * 100, 2) if total > 0 else 0.0
    best_stock = "-"
    if "symbol" in portfolio_df.columns and "pnl" in portfolio_df.columns and not portfolio_df.empty:
        best_stock = portfolio_df.sort_values("pnl", ascending=False).iloc[0]["symbol"]

    gain_col = "pct_change" if "pct_change" in gainers_df.columns else ("pct_change_est" if "pct_change_est" in gainers_df.columns else None)
    lose_col = "pct_change" if "pct_change" in losers_df.columns else ("pct_change_est" if "pct_change_est" in losers_df.columns else None)
    if gain_col:
        gainers_df[gain_col] = pd.to_numeric(gainers_df[gain_col], errors="coerce").fillna(0).apply(
            lambda x: f'<span class="positive">{round(x, 2)}%</span>'
        )
    if lose_col:
        losers_df[lose_col] = pd.to_numeric(losers_df[lose_col], errors="coerce").fillna(0).apply(
            lambda x: f'<span class="negative">{round(x, 2)}%</span>'
        )
    momentum_df = _add_signal_column(momentum_df)
    momentum_df["signal"] = momentum_df["signal"].astype(str).str.upper().apply(
        lambda x: f'<span class="{x.lower()}">{x}</span>'
    )

    # Portfolio styling + hold/review suggestion.
    portfolio_fmt = portfolio_df.copy()
    portfolio_fmt["pnl"] = pd.to_numeric(portfolio_fmt.get("pnl", 0), errors="coerce").fillna(0.0)
    portfolio_fmt["pnl_pct"] = pd.to_numeric(portfolio_fmt.get("pnl_pct", 0), errors="coerce").fillna(0.0)

    # Join fundamentals for better hold/review signal.
    common_df = pd.merge(
        portfolio_fmt,
        fundamentals_df[["symbol", "roe", "debt", "sales_growth", "profit_growth", "fund_score"]],
        on="symbol",
        how="inner",
    )
    low_fund_cutoff = common_df["fund_score"].quantile(0.4) if not common_df.empty else 0.0

    def suggest_action(row: pd.Series) -> str:
        pnl_pct = float(row.get("pnl_pct", 0.0))
        fund_score = float(row.get("fund_score", 0.0))
        if pnl_pct < -10 and fund_score <= low_fund_cutoff:
            return "REVIEW"
        return "HOLD"

    merged_for_signal = pd.merge(
        portfolio_fmt,
        fundamentals_df[["symbol", "fund_score"]],
        on="symbol",
        how="left",
    )
    merged_for_signal["suggestion"] = merged_for_signal.apply(suggest_action, axis=1)
    merged_for_signal["performance"] = merged_for_signal["pnl"].apply(
        lambda x: "GAIN" if x >= 0 else "LOSS"
    )
    portfolio_chart = _build_bar_chart(merged_for_signal, "symbol", "pnl", signed=True)

    common_numeric = common_df.copy()
    common_chart = _build_bar_chart(common_numeric, "symbol", "fund_score", signed=False)

    merged_for_signal["pnl"] = merged_for_signal["pnl"].apply(
        lambda x: f'<span class="positive">{round(x, 2)}</span>' if x >= 0 else f'<span class="negative">{round(x, 2)}</span>'
    )
    merged_for_signal["pnl_pct"] = merged_for_signal["pnl_pct"].apply(
        lambda x: f'<span class="positive">{round(x, 2)}%</span>' if x >= 0 else f'<span class="negative">{round(x, 2)}%</span>'
    )
    merged_for_signal["performance"] = merged_for_signal["performance"].apply(
        lambda x: f'<span class="positive">{x}</span>' if x == "GAIN" else f'<span class="negative">{x}</span>'
    )
    merged_for_signal["suggestion"] = merged_for_signal["suggestion"].apply(
        lambda x: f'<span class="hold">{x}</span>' if x == "HOLD" else f'<span class="sell">{x}</span>'
    )

    if not common_df.empty:
        common_df["pnl"] = pd.to_numeric(common_df.get("pnl", 0), errors="coerce").fillna(0.0)
        common_df["pnl_pct"] = pd.to_numeric(common_df.get("pnl_pct", 0), errors="coerce").fillna(0.0)
        common_df["suggestion"] = common_df.apply(suggest_action, axis=1)
        common_df = common_df[
            ["symbol", "pnl", "pnl_pct", "roe", "debt", "sales_growth", "profit_growth", "fund_score", "suggestion"]
        ].sort_values("fund_score", ascending=False)
        common_df["pnl"] = common_df["pnl"].apply(
            lambda x: f'<span class="positive">{round(x, 2)}</span>' if x >= 0 else f'<span class="negative">{round(x, 2)}</span>'
        )
        common_df["pnl_pct"] = common_df["pnl_pct"].apply(
            lambda x: f'<span class="positive">{round(x, 2)}%</span>' if x >= 0 else f'<span class="negative">{round(x, 2)}%</span>'
        )
        common_df["suggestion"] = common_df["suggestion"].apply(
            lambda x: f'<span class="hold">{x}</span>' if x == "HOLD" else f'<span class="sell">{x}</span>'
        )
    else:
        common_df = pd.DataFrame([{"symbol": "No common stocks found", "suggestion": "-"}])

    gainers = gainers_df.to_html(index=False, escape=False)
    losers = losers_df.to_html(index=False, escape=False)
    momentum = momentum_df.to_html(index=False, escape=False)
    portfolio = merged_for_signal.to_html(index=False, escape=False)
    common = common_df.to_html(index=False, escape=False)

    return render_template_string(
        HTML,
        gainers=gainers,
        losers=losers,
        momentum=momentum,
        portfolio=portfolio,
        common=common,
        portfolio_chart=portfolio_chart,
        common_chart=common_chart,
        total_value=round(total_value, 2),
        total_pnl=round(total_pnl, 2),
        win_rate=win_rate,
        best_stock=best_stock
    )

if __name__ == "__main__":
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "5001"))
    app.run(host=host, port=port, debug=False, use_reloader=False)
