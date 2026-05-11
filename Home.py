"""
Right Horizons Terminal — Scanner Hub
Landing page with all four scanners as a clean 2x2 grid.
"""
import streamlit as st
from theme import (
    apply_theme, render_header,
    RH_GOLD, RH_GOLD_LIGHT, RH_GOLD_DIM, RH_RED, RH_GREEN,
    RH_BG, RH_SURFACE, RH_TEXT, RH_MUTED, RH_BORDER
)

st.set_page_config(
    page_title="Right Horizons | Scanner Hub",
    layout="wide",
    initial_sidebar_state="expanded",
    page_icon="◆",
)
apply_theme()
render_header("Quantitative Scanner Suite · Hub")


# ─────────────────────────────────────────────────────
#  SCANNER REGISTRY
#  All four scanners are real and wired. Edit metric values
#  here once your live data comes in, or pull them dynamically
#  from each scanner's data file (see comments at the bottom).
# ─────────────────────────────────────────────────────
SCANNERS = [
    {
        "n": "01",
        "title": "Nifty 500 Breadth",
        "tagline": "Market internals",
        "desc": "Synchronized price + breadth action across the Nifty 500. Tracks 52-week highs and lows, net breadth, and the share of stocks trading above the 200-day moving average.",
        "metrics": [
            ("Stocks > 200 SMA", "62.4%", "+1.8%", "up"),
            ("Net breadth", "+47", "+12", "up"),
        ],
        "tags": ["Breadth", "Daily", "Nifty 500"],
        "page": "1_Breadth_Scanner",
        "status": "LIVE",
    },
    {
        "n": "02",
        "title": "Global ETF Screener",
        "tagline": "Capital rotation",
        "desc": "Finviz-style treemap covering 290+ global ETFs across regions, sectors, and themes. Box size shows 30-day volume, color encodes performance over your chosen window.",
        "metrics": [
            ("ETFs tracked", "290", "Active universe", "neutral"),
            ("Avg 1D return", "+0.42%", "+12 gainers", "up"),
        ],
        "tags": ["ETF", "Global", "Rotation"],
        "page": "2_ETF_Screener",
        "status": "LIVE",
    },
    {
        "n": "03",
        "title": "NSE Swing Breakout",
        "tagline": "Momentum signals",
        "desc": "End-of-day breakout scanner across 80+ NSE large/mid-caps. Detects close above the most recent swing high with 1.2× average volume confirmation.",
        "metrics": [
            ("Breakouts today", "7", "+3 vs yesterday", "up"),
            ("Vol confirmation", "1.2×", "Avg threshold", "neutral"),
        ],
        "tags": ["Swing", "Momentum", "EOD"],
        "page": "3_Swing_Breakout",
        "status": "LIVE",
    },
    {
        "n": "04",
        "title": "India Custom Research",
        "tagline": "Fundamentals deep-dive",
        "desc": "Full fundamentals screener across 700+ NSE tickers. Save custom filter sets, track ratios over time, and pull fresh news. Refreshes monthly via GitHub Actions.",
        "metrics": [
            ("Universe", "700+", "NSE tickers", "neutral"),
            ("Last refresh", "Monthly", "Auto-pipeline", "neutral"),
        ],
        "tags": ["Fundamentals", "Custom", "NSE"],
        "page": "4_India_Research",
        "status": "LIVE",
    },
    {
    "n": "05",
    "title": "Liquid ETF Screener",
    "tagline": "160 tickers · 21 categories",
    "desc": "Capital rotation treemap across the most liquid global ETFs — equity, fixed income, commodities, thematic and geography.",
    "metrics": [
        ("Universe", "156", "ETFs", "neutral"),
        ("Categories", "21", "curated", "neutral"),
    ],
    "tags": ["ETF", "Global", "Treemap", "Rotation"],
    "page": "6_Liquid_ETF_Screener",
    "status": "LIVE",
    },
    {                # ← paste HERE, INSIDE the [ ]
        "n": "06",
        "title": "Valuation Tracker",
        "tagline": "Price vs EPS · 25 indices",
        "desc": "Tracks cumulative price CAGR vs EPS CAGR across all major NSE indices + Nasdaq 100. Flags where price has re-rated faster than earnings (overvalued) or de-rated against fundamentals (undervalued).",
        "metrics": [
            ("Indices", "25", "tracked", "neutral"),
            ("Updates", "Daily", "4:30 PM IST", "neutral"),
        ],
        "tags": ["Valuation", "Macro", "Sector"],
        "page": "7_Valuation_Tracker",
        "status": "LIVE",
    },
]
# ─────────────────────────────────────────────────────
#  TILE STYLES — refined card design
# ─────────────────────────────────────────────────────
st.markdown(f"""
<style>
.scanner-tile {{
    background: {RH_SURFACE};
    border: 1px solid {RH_BORDER};
    padding: 22px 24px 20px;
    min-height: 280px;
    display: flex;
    flex-direction: column;
    position: relative;
    transition: all 0.2s ease;
    margin-bottom: 8px;
}}

.scanner-tile:hover {{
    border-color: {RH_GOLD_DIM};
    background: #1A1A1A;
}}

.scanner-tile::before {{
    content: '';
    position: absolute;
    top: 0; left: 0;
    width: 3px;
    height: 100%;
    background: linear-gradient(to bottom, {RH_GOLD} 0%, transparent 60%);
    opacity: 0.6;
}}

.tile-head {{
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    margin-bottom: 14px;
}}

.tile-num-block {{
    display: flex;
    align-items: baseline;
    gap: 12px;
}}

.tile-num {{
    font-family: 'Fraunces', serif;
    font-size: 38px;
    font-weight: 900;
    color: {RH_GOLD_DIM};
    line-height: 1;
    letter-spacing: -0.02em;
}}

.tile-tagline {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 9px;
    color: {RH_GOLD_LIGHT};
    letter-spacing: 0.18em;
    text-transform: uppercase;
}}

.tile-status {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 9px;
    letter-spacing: 0.15em;
    padding: 3px 8px;
    border: 1px solid;
    color: {RH_GREEN};
    border-color: rgba(46,204,113,0.4);
    background: rgba(46,204,113,0.05);
}}
.tile-status.beta {{ color: {RH_GOLD_LIGHT}; border-color: rgba(212,168,48,0.4); background: rgba(212,168,48,0.05); }}
.tile-status.soon {{ color: {RH_MUTED}; border-color: rgba(122,112,96,0.4); background: transparent; }}

.tile-title {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 16px;
    font-weight: 500;
    color: {RH_TEXT};
    letter-spacing: 0.05em;
    text-transform: uppercase;
    margin: 0 0 8px;
    line-height: 1.3;
}}

.tile-desc {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11.5px;
    color: {RH_MUTED};
    line-height: 1.65;
    margin-bottom: 16px;
    letter-spacing: 0.02em;
    flex-grow: 1;
}}

.tile-tags {{
    display: flex;
    gap: 6px;
    flex-wrap: wrap;
    margin-bottom: 14px;
}}

.tile-tag {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 9px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: {RH_GOLD_DIM};
    border: 1px dashed {RH_BORDER};
    padding: 2px 8px;
}}

.tile-metrics {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
    border-top: 1px dashed {RH_BORDER};
    padding-top: 14px;
}}

.tile-metric-block {{}}

.tile-metric-label {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 8px;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: {RH_MUTED};
    margin-bottom: 3px;
}}

.tile-metric-value {{
    font-family: 'Fraunces', serif;
    font-size: 20px;
    font-weight: 700;
    color: {RH_TEXT};
    line-height: 1;
    margin-bottom: 4px;
}}

.tile-metric-delta {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    letter-spacing: 0.05em;
}}
.delta-up {{ color: {RH_GREEN}; }}
.delta-dn {{ color: {RH_RED}; }}
.delta-nu {{ color: {RH_MUTED}; }}

/* Sub-strip below header */
.rh-substrip {{
    display: flex;
    gap: 28px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    color: {RH_MUTED};
    letter-spacing: 0.1em;
    text-transform: uppercase;
    padding: 4px 0 18px;
    border-bottom: 1px solid {RH_BORDER};
    margin-bottom: 22px;
    flex-wrap: wrap;
}}
.rh-substrip strong {{
    color: {RH_GOLD_LIGHT};
    font-weight: 500;
    margin-right: 6px;
}}
.rh-substrip .right {{
    margin-left: auto;
}}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────
#  STATUS STRIP
# ─────────────────────────────────────────────────────
live_count = sum(1 for s in SCANNERS if s["status"] == "LIVE")
total_count = len(SCANNERS)

st.markdown(f"""
<div class="rh-substrip">
    <span><strong>{live_count}</strong>Active scanners</span>
    <span><strong>{total_count}</strong>Total tools</span>
    <span><strong>NSE · Global ETF</strong>Coverage</span>
    <span class="right">Internal Quantitative Tool · v1.0</span>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────
#  RENDER SCANNERS — 2x2 grid
# ─────────────────────────────────────────────────────
def render_metric(label, value, delta, direction):
    delta_class = {"up": "delta-up", "down": "delta-dn", "neutral": "delta-nu"}[direction]
    return f"""
    <div class="tile-metric-block">
        <div class="tile-metric-label">{label}</div>
        <div class="tile-metric-value">{value}</div>
        <div class="tile-metric-delta {delta_class}">{delta}</div>
    </div>
    """


def render_tile(s):
    status_class = s["status"].lower()
    metrics_html = "".join(
        render_metric(label, val, delta, direction)
        for label, val, delta, direction in s["metrics"]
    )
    tags_html = "".join(f'<span class="tile-tag">{t}</span>' for t in s["tags"])

    return f"""
    <div class="scanner-tile">
        <div class="tile-head">
            <div class="tile-num-block">
                <div class="tile-num">{s['n']}</div>
                <div class="tile-tagline">{s['tagline']}</div>
            </div>
            <div class="tile-status {status_class}">{s['status']}</div>
        </div>
        <div class="tile-title">{s['title']}</div>
        <div class="tile-desc">{s['desc']}</div>
        <div class="tile-tags">{tags_html}</div>
        <div class="tile-metrics">{metrics_html}</div>
    </div>
    """


# Render in 2-column rows so the 2×2 grid feels balanced
rows = [SCANNERS[i:i+2] for i in range(0, len(SCANNERS), 2)]

for row in rows:
    cols = st.columns(2, gap="medium")
    for col, scanner in zip(cols, row):
        with col:
            st.markdown(render_tile(scanner), unsafe_allow_html=True)
            if st.button(
                f"OPEN {scanner['title'].upper()}  →",
                key=f"open_{scanner['n']}",
                use_container_width=True,
            ):
                st.switch_page(f"pages/{scanner['page']}.py")
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────
#  FOOTER
# ─────────────────────────────────────────────────────
st.markdown(f"""
<div style='margin-top:32px; padding-top:14px; border-top:1px solid {RH_BORDER};
            font-family:IBM Plex Mono; font-size:9px; color:{RH_MUTED};
            letter-spacing:0.1em; text-transform:uppercase;'>
    ◇ Right Horizons Wealth Management · Internal Tool · For research and informational use only
</div>
""", unsafe_allow_html=True)
