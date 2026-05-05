"""
Daily baseline builder for Liquid ETF Screener (Scanner 6).
Computes returns over standard windows + 52W high/low + 30D avg volume
for the 160 ETFs from the curated liquidity list.
Output: data/liquid_etf_baseline.csv

Designed to be run by .github/workflows/update_liquid_etf.yml
"""
import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta
import os, sys

# ── Master ETF list (160 ETFs across 21 categories) ──────────────────────────
ETF_DATA = [
    # US Equity – Broad Market
    {"Category":"US Equity – Broad Market","Segment":"S&P 500","Ticker":"SPY","Name":"SPDR S&P 500 ETF Trust","AMC":"SPDR"},
    {"Category":"US Equity – Broad Market","Segment":"S&P 500 Low Cost","Ticker":"VOO","Name":"Vanguard S&P 500 ETF","AMC":"Vanguard"},
    {"Category":"US Equity – Broad Market","Segment":"Nasdaq 100","Ticker":"QQQ","Name":"Invesco QQQ Trust","AMC":"Invesco"},
    {"Category":"US Equity – Broad Market","Segment":"Total US Market","Ticker":"VTI","Name":"Vanguard Total Stock Market ETF","AMC":"Vanguard"},
    {"Category":"US Equity – Broad Market","Segment":"Dow Jones 30","Ticker":"DIA","Name":"SPDR Dow Jones Industrial Average","AMC":"SPDR"},
    {"Category":"US Equity – Broad Market","Segment":"Russell 2000","Ticker":"IWM","Name":"iShares Russell 2000 ETF","AMC":"iShares"},
    {"Category":"US Equity – Broad Market","Segment":"S&P 400 Mid Cap","Ticker":"IJH","Name":"iShares Core S&P Mid-Cap ETF","AMC":"iShares"},
    {"Category":"US Equity – Broad Market","Segment":"Equal Weight S&P 500","Ticker":"RSP","Name":"Invesco S&P 500 Equal Weight ETF","AMC":"Invesco"},
    # US Equity – Sector
    {"Category":"US Equity – Sector","Segment":"Technology","Ticker":"XLK","Name":"Technology Select Sector SPDR","AMC":"SPDR"},
    {"Category":"US Equity – Sector","Segment":"Financials","Ticker":"XLF","Name":"Financial Select Sector SPDR","AMC":"SPDR"},
    {"Category":"US Equity – Sector","Segment":"Healthcare","Ticker":"XLV","Name":"Health Care Select Sector SPDR","AMC":"SPDR"},
    {"Category":"US Equity – Sector","Segment":"Energy","Ticker":"XLE","Name":"Energy Select Sector SPDR","AMC":"SPDR"},
    {"Category":"US Equity – Sector","Segment":"Industrials","Ticker":"XLI","Name":"Industrial Select Sector SPDR","AMC":"SPDR"},
    {"Category":"US Equity – Sector","Segment":"Consumer Disc","Ticker":"XLY","Name":"Consumer Discretionary SPDR","AMC":"SPDR"},
    {"Category":"US Equity – Sector","Segment":"Consumer Staples","Ticker":"XLP","Name":"Consumer Staples SPDR","AMC":"SPDR"},
    {"Category":"US Equity – Sector","Segment":"Materials","Ticker":"XLB","Name":"Materials Select Sector SPDR","AMC":"SPDR"},
    {"Category":"US Equity – Sector","Segment":"Utilities","Ticker":"XLU","Name":"Utilities Select Sector SPDR","AMC":"SPDR"},
    {"Category":"US Equity – Sector","Segment":"Real Estate","Ticker":"XLRE","Name":"Real Estate Select Sector SPDR","AMC":"SPDR"},
    {"Category":"US Equity – Sector","Segment":"Comm Services","Ticker":"XLC","Name":"Communication Services SPDR","AMC":"SPDR"},
    # US Equity – Market Cap
    {"Category":"US Equity – Market Cap","Segment":"Mega Cap","Ticker":"MGC","Name":"Vanguard Mega Cap ETF","AMC":"Vanguard"},
    {"Category":"US Equity – Market Cap","Segment":"Large Cap","Ticker":"IVV","Name":"iShares Core S&P 500 ETF","AMC":"iShares"},
    {"Category":"US Equity – Market Cap","Segment":"Mid Cap","Ticker":"VO","Name":"Vanguard Mid-Cap ETF","AMC":"Vanguard"},
    {"Category":"US Equity – Market Cap","Segment":"Micro Cap","Ticker":"IWC","Name":"iShares Micro-Cap ETF","AMC":"iShares"},
    # US Equity – Style
    {"Category":"US Equity – Style","Segment":"Large Cap Growth","Ticker":"IVW","Name":"iShares S&P 500 Growth ETF","AMC":"iShares"},
    {"Category":"US Equity – Style","Segment":"Large Cap Value","Ticker":"IVE","Name":"iShares S&P 500 Value ETF","AMC":"iShares"},
    {"Category":"US Equity – Style","Segment":"Mid Cap Growth","Ticker":"IJK","Name":"iShares S&P Mid-Cap 400 Growth","AMC":"iShares"},
    {"Category":"US Equity – Style","Segment":"Mid Cap Value","Ticker":"IJJ","Name":"iShares S&P Mid-Cap 400 Value","AMC":"iShares"},
    {"Category":"US Equity – Style","Segment":"Small Cap Growth","Ticker":"IJT","Name":"iShares S&P Small-Cap 600 Growth","AMC":"iShares"},
    {"Category":"US Equity – Style","Segment":"Small Cap Value","Ticker":"IJS","Name":"iShares S&P Small-Cap 600 Value","AMC":"iShares"},
    # Factor / Smart Beta
    {"Category":"Factor / Smart Beta","Segment":"Quality","Ticker":"QUAL","Name":"iShares MSCI USA Quality Factor ETF","AMC":"iShares"},
    {"Category":"Factor / Smart Beta","Segment":"Momentum","Ticker":"MTUM","Name":"iShares MSCI USA Momentum Factor ETF","AMC":"iShares"},
    {"Category":"Factor / Smart Beta","Segment":"Value","Ticker":"VLUE","Name":"iShares MSCI USA Value Factor ETF","AMC":"iShares"},
    {"Category":"Factor / Smart Beta","Segment":"Min Volatility","Ticker":"USMV","Name":"iShares MSCI USA Min Vol Factor ETF","AMC":"iShares"},
    {"Category":"Factor / Smart Beta","Segment":"Low Volatility","Ticker":"SPLV","Name":"Invesco S&P 500 Low Volatility ETF","AMC":"Invesco"},
    {"Category":"Factor / Smart Beta","Segment":"High Dividend","Ticker":"HDV","Name":"iShares Core High Dividend ETF","AMC":"iShares"},
    {"Category":"Factor / Smart Beta","Segment":"Dividend Growth","Ticker":"VIG","Name":"Vanguard Dividend Appreciation ETF","AMC":"Vanguard"},
    {"Category":"Factor / Smart Beta","Segment":"Large Cap Multifactor","Ticker":"LRGF","Name":"iShares US Equity Factor ETF","AMC":"iShares"},
    {"Category":"Factor / Smart Beta","Segment":"Buyback","Ticker":"PKW","Name":"Invesco Buyback Achievers ETF","AMC":"Invesco"},
    {"Category":"Factor / Smart Beta","Segment":"Profitability","Ticker":"VPU","Name":"Vanguard Utilities ETF","AMC":"Vanguard"},
    # Fixed Income – US
    {"Category":"Fixed Income – US","Segment":"Aggregate Bond","Ticker":"AGG","Name":"iShares Core US Aggregate Bond ETF","AMC":"iShares"},
    {"Category":"Fixed Income – US","Segment":"Short Term Treasury","Ticker":"SHY","Name":"iShares 1-3 Year Treasury Bond ETF","AMC":"iShares"},
    {"Category":"Fixed Income – US","Segment":"Intermediate Treasury","Ticker":"IEF","Name":"iShares 7-10 Year Treasury Bond ETF","AMC":"iShares"},
    {"Category":"Fixed Income – US","Segment":"Long Term Treasury","Ticker":"TLT","Name":"iShares 20+ Year Treasury Bond ETF","AMC":"iShares"},
    {"Category":"Fixed Income – US","Segment":"TIPS","Ticker":"TIP","Name":"iShares TIPS Bond ETF","AMC":"iShares"},
    {"Category":"Fixed Income – US","Segment":"Muni Bond","Ticker":"MUB","Name":"iShares National Muni Bond ETF","AMC":"iShares"},
    {"Category":"Fixed Income – US","Segment":"Floating Rate","Ticker":"FLOT","Name":"iShares Floating Rate Bond ETF","AMC":"iShares"},
    # Fixed Income – Credit
    {"Category":"Fixed Income – Credit","Segment":"Investment Grade Corp","Ticker":"LQD","Name":"iShares iBoxx Investment Grade Corp Bond","AMC":"iShares"},
    {"Category":"Fixed Income – Credit","Segment":"High Yield","Ticker":"HYG","Name":"iShares iBoxx High Yield Corp Bond ETF","AMC":"iShares"},
    {"Category":"Fixed Income – Credit","Segment":"Short High Yield","Ticker":"SJNK","Name":"SPDR Bloomberg Short Term High Yield","AMC":"SPDR"},
    {"Category":"Fixed Income – Credit","Segment":"Preferred Stock","Ticker":"PFF","Name":"iShares Preferred & Income Securities ETF","AMC":"iShares"},
    # Fixed Income – Global
    {"Category":"Fixed Income – Global","Segment":"Emerging Market Bond","Ticker":"EMB","Name":"iShares JP Morgan USD Emerging Markets Bond","AMC":"iShares"},
    {"Category":"Fixed Income – Global","Segment":"Global Bond","Ticker":"BNDW","Name":"Vanguard Total World Bond ETF","AMC":"Vanguard"},
    {"Category":"Fixed Income – Global","Segment":"International Govt","Ticker":"IGOV","Name":"iShares International Treasury Bond ETF","AMC":"iShares"},
    # Commodities
    {"Category":"Commodities","Segment":"Gold (Large)","Ticker":"GLD","Name":"SPDR Gold Shares","AMC":"SPDR"},
    {"Category":"Commodities","Segment":"Gold (Low Cost)","Ticker":"IAU","Name":"iShares Gold Trust","AMC":"iShares"},
    {"Category":"Commodities","Segment":"Silver","Ticker":"SLV","Name":"iShares Silver Trust","AMC":"iShares"},
    {"Category":"Commodities","Segment":"Broad Commodities","Ticker":"PDBC","Name":"Invesco Optimum Yield Diversified Commodity","AMC":"Invesco"},
    {"Category":"Commodities","Segment":"Crude Oil","Ticker":"USO","Name":"United States Oil Fund LP","AMC":"USCF"},
    {"Category":"Commodities","Segment":"Energy Commodities","Ticker":"DBE","Name":"Invesco DB Energy Fund","AMC":"Invesco"},
    {"Category":"Commodities","Segment":"Natural Gas","Ticker":"UNG","Name":"United States Natural Gas Fund LP","AMC":"USCF"},
    {"Category":"Commodities","Segment":"Copper","Ticker":"CPER","Name":"United States Copper Index Fund","AMC":"USCF"},
    {"Category":"Commodities","Segment":"Wheat","Ticker":"WEAT","Name":"Teucrium Wheat Fund","AMC":"Teucrium"},
    {"Category":"Commodities","Segment":"Gold Miners","Ticker":"GDX","Name":"VanEck Gold Miners ETF","AMC":"VanEck"},
    # Alternatives
    {"Category":"Alternatives","Segment":"Global REIT","Ticker":"REET","Name":"iShares Global REIT ETF","AMC":"iShares"},
    {"Category":"Alternatives","Segment":"Managed Futures","Ticker":"DBMF","Name":"iMGP DBi Managed Futures Strategy ETF","AMC":"iMGP"},
    {"Category":"Alternatives","Segment":"Merger Arb","Ticker":"MNA","Name":"IQ Merger Arbitrage ETF","AMC":"IQ"},
    {"Category":"Alternatives","Segment":"Market Neutral","Ticker":"BTAL","Name":"AGF US Market Neutral Anti-Beta Fund","AMC":"AGF"},
    {"Category":"Alternatives","Segment":"VIX Short Term","Ticker":"UVXY","Name":"ProShares Ultra VIX Short-Term Futures ETF","AMC":"ProShares"},
    {"Category":"Alternatives","Segment":"S&P 500 Inverse","Ticker":"SH","Name":"ProShares Short S&P500","AMC":"ProShares"},
    {"Category":"Alternatives","Segment":"Balanced Allocation","Ticker":"AOM","Name":"iShares Core Moderate Allocation ETF","AMC":"iShares"},
    {"Category":"Alternatives","Segment":"Global Infrastructure","Ticker":"IGF","Name":"iShares Global Infrastructure ETF","AMC":"iShares"},
    # Thematic – Technology
    {"Category":"Thematic – Technology","Segment":"Semiconductors","Ticker":"SOXX","Name":"iShares Semiconductor ETF","AMC":"iShares"},
    {"Category":"Thematic – Technology","Segment":"Robotics & AI","Ticker":"BOTZ","Name":"Global X Robotics & AI ETF","AMC":"Global X"},
    {"Category":"Thematic – Technology","Segment":"Cloud Computing","Ticker":"SKYY","Name":"First Trust Cloud Computing ETF","AMC":"First Trust"},
    {"Category":"Thematic – Technology","Segment":"Cybersecurity","Ticker":"CIBR","Name":"First Trust Nasdaq Cybersecurity ETF","AMC":"First Trust"},
    {"Category":"Thematic – Technology","Segment":"Fintech","Ticker":"FINX","Name":"Global X FinTech ETF","AMC":"Global X"},
    {"Category":"Thematic – Technology","Segment":"IoT","Ticker":"SNSR","Name":"Global X Internet of Things ETF","AMC":"Global X"},
    {"Category":"Thematic – Technology","Segment":"Blockchain","Ticker":"BLOK","Name":"Amplify Transformational Data Sharing ETF","AMC":"Amplify"},
    {"Category":"Thematic – Technology","Segment":"eSports & Gaming","Ticker":"ESPO","Name":"VanEck Video Gaming and eSports ETF","AMC":"VanEck"},
    {"Category":"Thematic – Technology","Segment":"E-Commerce","Ticker":"IBUY","Name":"Amplify Online Retail ETF","AMC":"Amplify"},
    {"Category":"Thematic – Technology","Segment":"Disruptive Tech","Ticker":"ARKK","Name":"ARK Innovation ETF","AMC":"ARK"},
    # Thematic – Energy & Climate
    {"Category":"Thematic – Energy & Climate","Segment":"Clean Energy","Ticker":"ICLN","Name":"iShares Global Clean Energy ETF","AMC":"iShares"},
    {"Category":"Thematic – Energy & Climate","Segment":"Solar","Ticker":"TAN","Name":"Invesco Solar ETF","AMC":"Invesco"},
    {"Category":"Thematic – Energy & Climate","Segment":"EVs & Future Mobility","Ticker":"DRIV","Name":"Global X Autonomous & Electric Vehicles ETF","AMC":"Global X"},
    {"Category":"Thematic – Energy & Climate","Segment":"Lithium & Battery","Ticker":"LIT","Name":"Global X Lithium & Battery Tech ETF","AMC":"Global X"},
    {"Category":"Thematic – Energy & Climate","Segment":"Uranium","Ticker":"URNM","Name":"Sprott Uranium Miners ETF","AMC":"Sprott"},
    {"Category":"Thematic – Energy & Climate","Segment":"Water","Ticker":"PHO","Name":"Invesco Water Resources ETF","AMC":"Invesco"},
    # Thematic – Healthcare
    {"Category":"Thematic – Healthcare","Segment":"Biotech","Ticker":"IBB","Name":"iShares Biotechnology ETF","AMC":"iShares"},
    {"Category":"Thematic – Healthcare","Segment":"Genomics","Ticker":"ARKG","Name":"ARK Genomic Revolution ETF","AMC":"ARK"},
    {"Category":"Thematic – Healthcare","Segment":"Digital Health","Ticker":"EDOC","Name":"Global X Telemedicine & Digital Health ETF","AMC":"Global X"},
    # Thematic – Other
    {"Category":"Thematic – Other","Segment":"Defense & Aerospace","Ticker":"ITA","Name":"iShares US Aerospace & Defense ETF","AMC":"iShares"},
    {"Category":"Thematic – Other","Segment":"Agriculture","Ticker":"MOO","Name":"VanEck Agribusiness ETF","AMC":"VanEck"},
    {"Category":"Thematic – Other","Segment":"Luxury","Ticker":"LUXE","Name":"Emles Luxury Goods ETF","AMC":"Emles"},
    {"Category":"Thematic – Other","Segment":"Cannabis","Ticker":"MSOS","Name":"AdvisorShares Pure US Cannabis ETF","AMC":"AdvisorShares"},
    {"Category":"Thematic – Other","Segment":"Space","Ticker":"ARKX","Name":"ARK Space Exploration & Innovation ETF","AMC":"ARK"},
    {"Category":"Thematic – Other","Segment":"3D Printing","Ticker":"PRNT","Name":"3D Printing ETF","AMC":"ARK"},
    # Geography – Americas
    {"Category":"Geography – Americas","Segment":"Canada","Ticker":"EWC","Name":"iShares MSCI Canada ETF","AMC":"iShares"},
    {"Category":"Geography – Americas","Segment":"Brazil","Ticker":"EWZ","Name":"iShares MSCI Brazil ETF","AMC":"iShares"},
    {"Category":"Geography – Americas","Segment":"Mexico","Ticker":"EWW","Name":"iShares MSCI Mexico ETF","AMC":"iShares"},
    {"Category":"Geography – Americas","Segment":"Latin America","Ticker":"ILF","Name":"iShares Latin America 40 ETF","AMC":"iShares"},
    {"Category":"Geography – Americas","Segment":"Chile","Ticker":"ECH","Name":"iShares MSCI Chile ETF","AMC":"iShares"},
    {"Category":"Geography – Americas","Segment":"Colombia","Ticker":"GXG","Name":"Global X MSCI Colombia ETF","AMC":"Global X"},
    {"Category":"Geography – Americas","Segment":"Peru","Ticker":"EPU","Name":"iShares MSCI All Peru Capped ETF","AMC":"iShares"},
    {"Category":"Geography – Americas","Segment":"Argentina","Ticker":"ARGT","Name":"Global X MSCI Argentina ETF","AMC":"Global X"},
    # Geography – Europe
    {"Category":"Geography – Europe","Segment":"Europe Broad","Ticker":"VGK","Name":"Vanguard FTSE Europe ETF","AMC":"Vanguard"},
    {"Category":"Geography – Europe","Segment":"UK","Ticker":"EWU","Name":"iShares MSCI United Kingdom ETF","AMC":"iShares"},
    {"Category":"Geography – Europe","Segment":"Germany","Ticker":"EWG","Name":"iShares MSCI Germany ETF","AMC":"iShares"},
    {"Category":"Geography – Europe","Segment":"France","Ticker":"EWQ","Name":"iShares MSCI France ETF","AMC":"iShares"},
    {"Category":"Geography – Europe","Segment":"Switzerland","Ticker":"EWL","Name":"iShares MSCI Switzerland ETF","AMC":"iShares"},
    {"Category":"Geography – Europe","Segment":"Italy","Ticker":"EWI","Name":"iShares MSCI Italy ETF","AMC":"iShares"},
    {"Category":"Geography – Europe","Segment":"Spain","Ticker":"EWP","Name":"iShares MSCI Spain ETF","AMC":"iShares"},
    {"Category":"Geography – Europe","Segment":"Netherlands","Ticker":"EWN","Name":"iShares MSCI Netherlands ETF","AMC":"iShares"},
    {"Category":"Geography – Europe","Segment":"Sweden","Ticker":"EWD","Name":"iShares MSCI Sweden ETF","AMC":"iShares"},
    {"Category":"Geography – Europe","Segment":"Poland","Ticker":"EPOL","Name":"iShares MSCI Poland ETF","AMC":"iShares"},
    {"Category":"Geography – Europe","Segment":"Turkey","Ticker":"TUR","Name":"iShares MSCI Turkey ETF","AMC":"iShares"},
    {"Category":"Geography – Europe","Segment":"Greece","Ticker":"GREK","Name":"Global X MSCI Greece ETF","AMC":"Global X"},
    {"Category":"Geography – Europe","Segment":"Norway","Ticker":"ENOR","Name":"iShares MSCI Norway ETF","AMC":"iShares"},
    {"Category":"Geography – Europe","Segment":"Denmark","Ticker":"EDEN","Name":"iShares MSCI Denmark ETF","AMC":"iShares"},
    # Geography – Asia-Pacific
    {"Category":"Geography – Asia-Pacific","Segment":"Asia ex-Japan","Ticker":"AAXJ","Name":"iShares MSCI All Country Asia ex Japan ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"SE Asia","Ticker":"ASEA","Name":"Global X FTSE Southeast Asia ETF","AMC":"Global X"},
    {"Category":"Geography – Asia-Pacific","Segment":"Japan","Ticker":"EWJ","Name":"iShares MSCI Japan ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"China","Ticker":"MCHI","Name":"iShares MSCI China ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"Hong Kong","Ticker":"EWH","Name":"iShares MSCI Hong Kong ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"Taiwan","Ticker":"EWT","Name":"iShares MSCI Taiwan ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"South Korea","Ticker":"EWY","Name":"iShares MSCI South Korea ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"India","Ticker":"INDA","Name":"iShares MSCI India ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"Australia","Ticker":"EWA","Name":"iShares MSCI Australia ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"Singapore","Ticker":"EWS","Name":"iShares MSCI Singapore ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"Indonesia","Ticker":"EIDO","Name":"iShares MSCI Indonesia ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"Malaysia","Ticker":"EWM","Name":"iShares MSCI Malaysia ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"Thailand","Ticker":"THD","Name":"iShares MSCI Thailand ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"Philippines","Ticker":"EPHE","Name":"iShares MSCI Philippines ETF","AMC":"iShares"},
    {"Category":"Geography – Asia-Pacific","Segment":"Vietnam","Ticker":"VNM","Name":"VanEck Vietnam ETF","AMC":"VanEck"},
    {"Category":"Geography – Asia-Pacific","Segment":"Pakistan","Ticker":"PAK","Name":"Global X MSCI Pakistan ETF","AMC":"Global X"},
    {"Category":"Geography – Asia-Pacific","Segment":"New Zealand","Ticker":"ENZL","Name":"iShares MSCI New Zealand ETF","AMC":"iShares"},
    # Geography – Middle East & Africa
    {"Category":"Geography – Middle East & Africa","Segment":"Israel","Ticker":"EIS","Name":"iShares MSCI Israel ETF","AMC":"iShares"},
    {"Category":"Geography – Middle East & Africa","Segment":"Saudi Arabia","Ticker":"KSA","Name":"iShares MSCI Saudi Arabia ETF","AMC":"iShares"},
    {"Category":"Geography – Middle East & Africa","Segment":"UAE","Ticker":"UAE","Name":"iShares MSCI UAE ETF","AMC":"iShares"},
    {"Category":"Geography – Middle East & Africa","Segment":"Qatar","Ticker":"QAT","Name":"iShares MSCI Qatar ETF","AMC":"iShares"},
    {"Category":"Geography – Middle East & Africa","Segment":"Kuwait","Ticker":"KWT","Name":"iShares MSCI Kuwait ETF","AMC":"iShares"},
    {"Category":"Geography – Middle East & Africa","Segment":"Egypt","Ticker":"EGPT","Name":"VanEck Egypt Index ETF","AMC":"VanEck"},
    {"Category":"Geography – Middle East & Africa","Segment":"South Africa","Ticker":"EZA","Name":"iShares MSCI South Africa ETF","AMC":"iShares"},
    {"Category":"Geography – Middle East & Africa","Segment":"Nigeria","Ticker":"NGE","Name":"Global X MSCI Nigeria ETF","AMC":"Global X"},
    # Geography – Global / Multi-Region
    {"Category":"Geography – Global / Multi-Region","Segment":"Global All Country","Ticker":"VT","Name":"Vanguard Total World Stock ETF","AMC":"Vanguard"},
    {"Category":"Geography – Global / Multi-Region","Segment":"World Index","Ticker":"URTH","Name":"iShares MSCI World ETF","AMC":"iShares"},
    {"Category":"Geography – Global / Multi-Region","Segment":"Developed ex-US","Ticker":"EFA","Name":"iShares MSCI EAFE ETF","AMC":"iShares"},
    {"Category":"Geography – Global / Multi-Region","Segment":"All World ex-US","Ticker":"VEU","Name":"Vanguard FTSE All-World ex-US ETF","AMC":"Vanguard"},
    {"Category":"Geography – Global / Multi-Region","Segment":"Japan Hedged","Ticker":"DXJH","Name":"WisdomTree Japan Hedged Quality Dividend","AMC":"WisdomTree"},
    {"Category":"Geography – Global / Multi-Region","Segment":"Small Cap World","Ticker":"VSS","Name":"Vanguard FTSE All-World ex-US Small-Cap ETF","AMC":"Vanguard"},
    # Geography – Emerging Markets
    {"Category":"Geography – Emerging Markets","Segment":"EM Broad","Ticker":"VWO","Name":"Vanguard FTSE Emerging Markets ETF","AMC":"Vanguard"},
    {"Category":"Geography – Emerging Markets","Segment":"EM ex-China","Ticker":"EMXC","Name":"iShares MSCI Emerging Markets ex China ETF","AMC":"iShares"},
    {"Category":"Geography – Emerging Markets","Segment":"EM Small Cap","Ticker":"EWX","Name":"SPDR S&P Emerging Markets Small Cap ETF","AMC":"SPDR"},
    {"Category":"Geography – Emerging Markets","Segment":"EM Min Vol","Ticker":"EVAL","Name":"iShares MSCI EM Min Volatility Factor ETF","AMC":"iShares"},
    # Geography – Frontier Markets
    {"Category":"Geography – Frontier Markets","Segment":"Frontier Markets","Ticker":"FM","Name":"iShares MSCI Frontier and Select EM ETF","AMC":"iShares"},
    {"Category":"Geography – Frontier Markets","Segment":"Africa","Ticker":"AFK","Name":"VanEck Africa Index ETF","AMC":"VanEck"},
    {"Category":"Geography – Frontier Markets","Segment":"Kazakhstan","Ticker":"KAZ","Name":"iShares MSCI Kazakhstan ETF","AMC":"iShares"},
]


def build_baseline():
    master = pd.DataFrame(ETF_DATA)
    # Dedupe — same ticker may appear in multiple categories (like SPY, IWM, RSP).
    # Keep first occurrence so each ticker appears once.
    master = master.drop_duplicates(subset="Ticker", keep="first").reset_index(drop=True)

    tickers = master["Ticker"].tolist()
    print(f"[baseline] downloading {len(tickers)} ETFs from yfinance…")

    end_dt = (datetime.today() + timedelta(days=1)).strftime("%Y-%m-%d")
    start_dt = (datetime.today() - timedelta(days=400)).strftime("%Y-%m-%d")

    raw = yf.download(tickers, start=start_dt, end=end_dt,
                      auto_adjust=True, progress=False, threads=True, group_by="ticker")

    rows = []
    for tkr in tickers:
        try:
            if isinstance(raw.columns, pd.MultiIndex):
                if tkr not in raw.columns.get_level_values(0):
                    continue
                df = raw[tkr]
            else:
                df = raw
            closes = df["Close"].dropna() if "Close" in df.columns else pd.Series(dtype=float)
            vols   = df["Volume"].dropna() if "Volume" in df.columns else pd.Series(dtype=float)
            if len(closes) < 5:
                continue

            latest = float(closes.iloc[-1])
            def pct(n):
                return round((closes.iloc[-1]/closes.iloc[-n]-1)*100, 2) if len(closes)>=n else np.nan

            rows.append({
                "Ticker": tkr,
                "CMP_Baseline": round(latest, 2),
                "1W (%)":  pct(6),
                "1M (%)":  pct(22),
                "3M (%)":  pct(66),
                "6M (%)":  pct(132),
                "1Y (%)":  pct(252) if len(closes)>=252 else round((closes.iloc[-1]/closes.iloc[0]-1)*100, 2),
                "52W High": round(float(closes.tail(252).max()), 2),
                "52W Low":  round(float(closes.tail(252).min()), 2),
                "30D_Volume_Baseline": round(float(vols.tail(30).mean()), 0) if len(vols)>=5 else np.nan,
            })
        except Exception as e:
            print(f"[baseline] skip {tkr}: {e}")
            continue

    perf = pd.DataFrame(rows)
    out = master.merge(perf, on="Ticker", how="left")
    print(f"[baseline] {len(out)} rows · {perf['CMP_Baseline'].notna().sum()} with prices")
    return out


if __name__ == "__main__":
    df = build_baseline()
    os.makedirs("data", exist_ok=True)
    out_path = "data/liquid_etf_baseline.csv"
    df.to_csv(out_path, index=False)
    print(f"[baseline] wrote {out_path} · {len(df)} rows · {datetime.now()}")
