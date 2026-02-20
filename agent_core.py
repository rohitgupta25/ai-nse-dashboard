import pandas as pd
from nsepython import *
from ta.momentum import RSIIndicator
from ta.trend import SMAIndicator

# ==============================
# LOAD FUNDAMENTALS (EXCEL)
# ==============================
fund_df = pd.read_excel("fundamentals.xlsx")

fund_df.columns = fund_df.columns.str.strip().str.lower()

fund_df = fund_df.rename(columns={
    "name": "name",
    "roe %": "roe",
    "debt / eq": "debt",
    "sales var 3yrs %": "sales_growth",
    "profit growth 3years": "profit_growth",
    "market cap": "market_cap"
})

# Convert Name → NSE symbol
nse_symbols = nse_eq_symbols()

def name_to_symbol(name):
    name = str(name).lower().replace(" ", "")
    for sym in nse_symbols:
        if name in sym.lower():
            return sym
    return None

fund_df["symbol"] = fund_df["name"].apply(name_to_symbol)
fund_df = fund_df.dropna(subset=["symbol"])

# Fundamental scoring
fund_df["fund_score"] = (
    fund_df["roe"] * 0.4 +
    fund_df["sales_growth"] * 0.2 +
    fund_df["profit_growth"] * 0.2 +
    (1 / (fund_df["debt"] + 0.1)) * 0.2
)

fund_symbols = set(fund_df["symbol"])

# ==============================
# FETCH MARKET DATA
# ==============================
stocks = list(fund_symbols)  # faster: only scan fundamental stocks

data = []

for symbol in stocks:
    try:
        q = nse_eq_quote(symbol)
        price = q['priceInfo']['lastPrice']
        prev_close = q['priceInfo']['previousClose']
        volume = q['securityWiseDP']['quantityTraded']

        pct_change = ((price - prev_close) / prev_close) * 100

        data.append({
            "symbol": symbol,
            "price": price,
            "pct_change": pct_change,
            "volume": volume
        })
    except:
        continue

df = pd.DataFrame(data)

# ==============================
# TOP GAINERS / LOSERS
# ==============================
top_gainers = df.sort_values(by="pct_change", ascending=False).head(50)
top_losers = df.sort_values(by="pct_change").head(50)

# ==============================
# MOMENTUM + FUNDAMENTAL SCAN
# ==============================
momentum_list = []

for symbol in df["symbol"]:
    try:
        hist = equity_history(symbol, "6month")
        hist_df = pd.DataFrame(hist)

        hist_df['close'] = hist_df['CH_CLOSING_PRICE']
        hist_df['volume'] = hist_df['CH_TOT_TRADED_QTY']

        hist_df['rsi'] = RSIIndicator(hist_df['close'], window=14).rsi()
        hist_df['sma50'] = SMAIndicator(hist_df['close'], window=50).sma_indicator()
        hist_df['sma200'] = SMAIndicator(hist_df['close'], window=200).sma_indicator()
        hist_df['avg_volume'] = hist_df['volume'].rolling(20).mean()

        latest = hist_df.iloc[-1]
        signal = "HOLD"

        if (
            latest['close'] > latest['sma50'] and
            latest['sma50'] > latest['sma200'] and
            55 <= latest['rsi'] <= 70
        ):
            signal = "BUY"

        elif latest['rsi'] > 75 or latest['close'] < latest['sma50']:
            signal = "SELL"

            tech_score = latest['rsi'] + (latest['volume'] / latest['avg_volume'])

            momentum_list.append({
            "symbol": symbol,
            "price": latest['close'],
            "rsi": round(latest['rsi'], 2),
            "tech_score": tech_score,
            "signal": signal
        })

    except:
        continue

momentum_df = pd.DataFrame(momentum_list)

momentum_df = momentum_df.merge(
    fund_df[["symbol", "fund_score"]],
    on="symbol",
    how="left"
)

momentum_df["final_score"] = (
    momentum_df["tech_score"] * 0.6 +
    momentum_df["fund_score"] * 0.4
)

momentum_df = momentum_df.sort_values(by="final_score", ascending=False).head(50)

# ==============================
# PORTFOLIO TRACKING
# ==============================
portfolio = pd.read_excel("portfolio.xlsx")

portfolio_data = []

for _, row in portfolio.iterrows():
    symbol = row["symbol"]
    entry = row["entry_price"]
    qty = row["quantity"]

    try:
        q = nse_eq_quote(symbol)
        current = q['priceInfo']['lastPrice']

        pnl = (current - entry) * qty
        pnl_pct = ((current - entry) / entry) * 100

        portfolio_data.append({
            "symbol": symbol,
            "entry": entry,
            "current": current,
            "pnl": pnl,
            "pnl_pct": pnl_pct
        })
    except:
        continue

portfolio_df = pd.DataFrame(portfolio_data)

# ==============================
# SAVE OUTPUTS (EXCEL)
# ==============================
top_gainers.to_excel("outputs/top_gainers.xlsx", index=False)
top_losers.to_excel("outputs/top_losers.xlsx", index=False)
momentum_df.to_excel("outputs/potential_stocks.xlsx", index=False)
portfolio_df.to_excel("outputs/portfolio_performance.xlsx", index=False)