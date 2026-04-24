"""
app.py
Frontend Streamlit — Micro SaaS de Análise Financeira Institucional.
Orquestra: data_fetcher → claude_agent → pdf_generator + pptx_generator
"""

import streamlit as st
import json
import traceback
from datetime import datetime

from data_fetcher import fetch_all_data
from claude_agent import analyze_with_claude
from pdf_generator import generate_pdf
from pptx_generator import generate_pptx

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DeepResearch — Institutional Equity Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    .main-header {
        background: linear-gradient(135deg, #0D1B2A 0%, #1B2A4A 100%);
        padding: 2rem 2.5rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        border-left: 5px solid #1F5EFF;
    }
    .main-header h1 {
        color: #FFFFFF;
        font-size: 2.2rem;
        font-weight: 700;
        margin: 0 0 0.4rem 0;
    }
    .main-header p {
        color: #8A97A8;
        font-size: 0.95rem;
        margin: 0;
    }

    .verdict-card {
        padding: 1.5rem;
        border-radius: 10px;
        text-align: center;
        border: 1px solid #D0D8E8;
    }
    .verdict-buy  { background: #1A7A4A; color: white; }
    .verdict-hold { background: #D4A017; color: white; }
    .verdict-sell { background: #C0392B; color: white; }

    .metric-card {
        background: #F4F6FA;
        border-radius: 8px;
        padding: 1rem;
        border: 1px solid #E0E6F0;
        text-align: center;
    }
    .metric-label { color: #8A97A8; font-size: 0.8rem; font-weight: 600; text-transform: uppercase; }
    .metric-value { color: #0D1B2A; font-size: 1.6rem; font-weight: 700; margin-top: 4px; }

    .section-header {
        color: #1B2A4A;
        font-size: 1.1rem;
        font-weight: 600;
        padding: 0.5rem 0;
        border-bottom: 2px solid #1F5EFF;
        margin-bottom: 1rem;
    }

    .risk-card {
        background: #FFF5F5;
        border-left: 4px solid #C0392B;
        padding: 1rem;
        border-radius: 6px;
        margin-bottom: 0.8rem;
    }
    .risk-title { color: #C0392B; font-weight: 700; font-size: 0.95rem; }

    .stDownloadButton > button {
        background: linear-gradient(135deg, #1B2A4A, #1F5EFF);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.7rem 1.5rem;
        font-weight: 600;
        width: 100%;
        font-size: 1rem;
        cursor: pointer;
    }
    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #1F5EFF, #1B2A4A);
        transform: translateY(-1px);
    }

    div[data-testid="stStatusWidget"] { display: none; }
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📊 DeepResearch</h1>
    <p>Institutional-Grade Equity Analysis · Powered by Gemini AI · DCF + Comparables + Bear Case</p>
</div>
""", unsafe_allow_html=True)

# ── Session State ────────────────────────────────────────────────────────────
if "raw_data" not in st.session_state:
    st.session_state.raw_data = None
if "analysis" not in st.session_state:
    st.session_state.analysis = None
if "pdf_bytes" not in st.session_state:
    st.session_state.pdf_bytes = None
if "pptx_bytes" not in st.session_state:
    st.session_state.pptx_bytes = None
if "ticker" not in st.session_state:
    st.session_state.ticker = ""

# ── Input Row ─────────────────────────────────────────────────────────────────
col_input, col_btn, col_spacer = st.columns([3, 1.2, 4])

with col_input:
    ticker_input = st.text_input(
        "Enter Stock Ticker",
        placeholder="e.g. AAPL, NVDA, MSFT, ASML",
        value=st.session_state.ticker,
        label_visibility="collapsed",
    ).strip().upper()

with col_btn:
    run_button = st.button("🔍 Run Analysis", use_container_width=True, type="primary")

st.markdown("---")

# ── Validation ───────────────────────────────────────────────────────────────
def _validate_ticker(t: str) -> bool:
    if not t:
        st.warning("⚠️ Please enter a ticker symbol.")
        return False
    if len(t) > 10 or not t.replace(".", "").replace("-", "").isalnum():
        st.error("❌ Invalid ticker format. Use standard exchange symbols (e.g. AAPL, BRK.B).")
        return False
    return True

# ── Main Pipeline ─────────────────────────────────────────────────────────────
if run_button and _validate_ticker(ticker_input):
    st.session_state.ticker = ticker_input
    st.session_state.raw_data   = None
    st.session_state.analysis   = None
    st.session_state.pdf_bytes  = None
    st.session_state.pptx_bytes = None

    progress = st.progress(0, text="Initialising pipeline...")
    status   = st.empty()

    try:
        status.info(f"📡 **Phase 1 / 3** — Fetching market data for **{ticker_input}**...")
        progress.progress(15, text="Fetching financial data from market sources...")

        raw_data = fetch_all_data(ticker_input)
        st.session_state.raw_data = raw_data
        progress.progress(35, text="Data fetched successfully.")

        company_name = raw_data.get("company", {}).get("name", ticker_input)
        status.success(f"✅ Data collected for **{company_name}** ({ticker_input})")

        status.info(f"🤖 **Phase 2 / 3** — Running deep analysis with Gemini AI...")
        progress.progress(45, text="Sending data to AI — this may take 15-30 seconds...")

        analysis = analyze_with_claude(raw_data)
        st.session_state.analysis = analysis
        progress.progress(72, text="Analysis complete. Compiling documents...")

        status.info("📄 **Phase 3 / 3** — Generating PDF report...")
        progress.progress(80, text="Generating PDF report...")
        pdf_bytes = generate_pdf(analysis, ticker_input)
        st.session_state.pdf_bytes = pdf_bytes

        status.info("📊 **Phase 3 / 3** — Generating PPTX presentation...")
        progress.progress(93, text="Generating PowerPoint presentation...")
        pptx_bytes = generate_pptx(analysis, ticker_input)
        st.session_state.pptx_bytes = pptx_bytes

        progress.progress(100, text="Pipeline complete.")
        status.success(f"🎉 Analysis complete for **{company_name}**. Documents ready for download.")

    except Exception as e:
        progress.empty()
        st.error(f"❌ **Unexpected error:** {e}")
        with st.expander("🔧 Debug Traceback"):
            st.code(traceback.format_exc())
        st.stop()


# ── Results Display ───────────────────────────────────────────────────────────
if st.session_state.analysis:
    analysis = st.session_state.analysis
    ticker   = st.session_state.ticker

    # ==========================================
    # O DETETIVE DE JSON (Resolve o problema dos N/A)
    # Se o Gemini guardou os dados dentro de uma "caixa extra", nós tiramos de lá!
    # ==========================================
    if "company_summary" not in analysis and "verdict" not in analysis:
        # Procurar dentro das chaves se existe lá a informação
        for key, value in analysis.items():
            if isinstance(value, dict) and ("company_summary" in value or "verdict" in value):
                analysis = value
                break
    
    # Se for uma lista em vez de um dicionário (erro raro, mas previne crash)
    if isinstance(analysis, list) and len(analysis) > 0:
        analysis = analysis[0]
    # ==========================================

    verdict  = analysis.get("verdict", {})
    company  = analysis.get("company_summary", {})
    dcf      = analysis.get("dcf_model", {})
    tech     = analysis.get("technical_analysis", {})
    mult     = analysis.get("multiples_analysis", {})

    rating = str(verdict.get("rating", "N/A")).upper()
    
    # Fix para garantir que a cor muda mesmo que ele escreva "STRONG BUY"
    if "BUY" in rating:
        rating_class = "verdict-buy"
    elif "SELL" in rating:
        rating_class = "verdict-sell"
    else:
        rating_class = "verdict-hold"

    def _f(val, prefix="", suffix="", dec=2):
        if val is None or val == "" or str(val).upper() == "N/A":
            return "N/A"
        try:
            # Tira caracteres como $ ou % se o Gemini os enviou por engano
            clean_val = str(val).replace("$", "").replace("%", "").replace(",", "")
            return f"{prefix}{float(clean_val):.{dec}f}{suffix}"
        except:
            return str(val)

    # ── Summary KPIs ──────────────────────────────────────────────────────────
    st.markdown(f"### {company.get('name', ticker)} — Research Summary")

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Current Price</div>
            <div class="metric-value">{_f(verdict.get("current_price"), prefix="$")}</div>
        </div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">DCF Target</div>
            <div class="metric-value">{_f(verdict.get("dcf_target_price"), prefix="$")}</div>
        </div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Blended Target</div>
            <div class="metric-value">{_f(verdict.get("blended_target_price"), prefix="$")}</div>
        </div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Upside / Downside</div>
            <div class="metric-value">{_f(verdict.get("upside_pct"), suffix="%")}</div>
        </div>""", unsafe_allow_html=True)
    with k5:
        st.markdown(f"""
        <div class="verdict-card {rating_class}">
            <div style="font-size:0.8rem; opacity:0.85;">RATING</div>
            <div style="font-size:2rem; font-weight:700;">{rating}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Download Buttons ──────────────────────────────────────────────────────
    st.markdown('<div class="section-header">📥 Download Documents</div>', unsafe_allow_html=True)
    dl1, dl2, dl3 = st.columns([1, 1, 2])
    with dl1:
        if st.session_state.pdf_bytes:
            filename_pdf = f"{ticker}_DeepResearch_{datetime.utcnow().strftime('%Y%m%d')}.pdf"
            st.download_button(
                label="📄 Download PDF Report",
                data=st.session_state.pdf_bytes,
                file_name=filename_pdf,
                mime="application/pdf",
                use_container_width=True,
            )
    with dl2:
        if st.session_state.pptx_bytes:
            filename_pptx = f"{ticker}_Presentation_{datetime.utcnow().strftime('%Y%m%d')}.pptx"
            st.download_button(
                label="📊 Download Presentation",
                data=st.session_state.pptx_bytes,
                file_name=filename_pptx,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

    st.markdown("---")

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🏢 Business Model",
        "📈 Technical & Insiders",
        "💰 DCF Model",
        "📊 Comparables",
        "⚠️ Bear Case",
    ])

    # ── Tab 1: Business Model ─────────────────────────────────────────────────
    with tab1:
        biz = analysis.get("company_summary", {})
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Business Model**")
            st.write(biz.get("business_model", "N/A"))
            st.markdown("**Revenue Streams**")
            
            # Formatação segura de listas
            streams = biz.get("revenue_streams", [])
            if isinstance(streams, list):
                for rs in streams:
                    st.markdown(f"- {rs}")
            else:
                st.write(streams)
                
        with c2:
            st.markdown("**Competitive Moat Factors**")
            moats = biz.get("moat_factors", [])
            if isinstance(moats, list):
                for mf in moats:
                    st.markdown(f"- {mf}")
            else:
                st.write(moats)
                
            st.markdown("**Moat Durability Assessment**")
            st.write(biz.get("moat_assessment", "N/A"))

    # ── Tab 2: Technical ──────────────────────────────────────────────────────
    with tab2:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Price & Moving Averages**")
            st.table({
                "Metric": ["Current Price", "MA 50", "MA 200", "vs MA50 %", "vs MA200 %"],
                "Value":  [
                    _f(tech.get("current_price"), prefix="$"),
                    _f(tech.get("ma50"), prefix="$"),
                    _f(tech.get("ma200"), prefix="$"),
                    _f(tech.get("price_vs_ma50_pct"), suffix="%"),
                    _f(tech.get("price_vs_ma200_pct"), suffix="%"),
                ],
            })
        with c2:
            st.markdown("**Momentum Signals**")
            st.table({
                "Indicator": ["RSI (14)", "RSI Signal", "MACD Signal", "Trend Bias"],
                "Reading":   [
                    _f(tech.get("rsi_14")),
                    tech.get("rsi_signal", "N/A"),
                    tech.get("macd_signal", "N/A"),
                    tech.get("trend_bias", "N/A"),
                ],
            })
        st.markdown("**Insider Activity**")
        st.info(f"Pattern: **{tech.get('insider_pattern', 'N/A')}** — {tech.get('insider_summary', 'N/A')}")

    # ── Tab 3: DCF ────────────────────────────────────────────────────────────
    with tab3:
        g = dcf.get("growth_assumptions", {})
        if not isinstance(g, dict): g = {}
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Assumptions**")
            st.table({
                "Parameter": ["Base FCF", "WACC", "Terminal Growth", "Conservative Growth", "Base Growth", "Bull Growth"],
                "Value":     [
                    _f(dcf.get("base_fcf_bn"), suffix="B"),
                    _f(dcf.get("wacc_pct"), suffix="%"),
                    _f(dcf.get("terminal_growth_rate_pct"), suffix="%"),
                    _f(g.get("conservative"), suffix="%"),
                    _f(g.get("base"), suffix="%"),
                    _f(g.get("bull"), suffix="%"),
                ],
            })
        with c2:
            st.markdown("**Valuation Bridge**")
            st.table({
                "Component": ["Sum PV FCFs", "PV Terminal Value", "Enterprise Value", "Net Debt", "Equity Value", "Intrinsic Value/Share"],
                "Value ($B)": [
                    _f(dcf.get("sum_pv_fcfs_bn"), suffix="B"),
                    _f(dcf.get("pv_terminal_value_bn"), suffix="B"),
                    _f(dcf.get("enterprise_value_bn"), suffix="B"),
                    _f(dcf.get("net_debt_bn"), suffix="B"),
                    _f(dcf.get("equity_value_bn"), suffix="B"),
                    _f(dcf.get("dcf_intrinsic_value"), prefix="$"),
                ],
            })
        if dcf.get("dcf_notes"):
            st.caption(f"Notes: {dcf.get('dcf_notes')}")

    # ── Tab 4: Comparables ───────────────────────────────────────────────────
    with tab4:
        subj = mult.get("subject", {})
        p1   = mult.get("peer_1", {})
        p2   = mult.get("peer_2", {})

        import pandas as pd
        comp_df = pd.DataFrame({
            "Metric":  ["EV/EBITDA", "P/E (TTM)", "P/S (TTM)", "Rev Growth %", "Value/Growth Score"],
            ticker:    [
                _f(subj.get("ev_ebitda")), _f(subj.get("pe_ttm")),
                _f(subj.get("ps_ttm")),
                _f(subj.get("revenue_growth_yoy_pct"), suffix="%"),
                _f(subj.get("value_growth_score"), dec=3),
            ],
            f"{p1.get('name','Peer 1')} ({p1.get('ticker','')})": [
                _f(p1.get("ev_ebitda")), _f(p1.get("pe_ttm")), _f(p1.get("ps_ttm")), "—", "—",
            ],
            f"{p2.get('name','Peer 2')} ({p2.get('ticker','')})": [
                _f(p2.get("ev_ebitda")), _f(p2.get("pe_ttm")), _f(p2.get("ps_ttm")), "—", "—",
            ],
        }).set_index("Metric")
        st.dataframe(comp_df, use_container_width=True)
        st.success(f"**Multiples-Implied Price:** {_f(mult.get('multiples_implied_price'), prefix='$')}")
        if mult.get("multiples_methodology"):
            st.caption(mult.get("multiples_methodology"))

    # ── Tab 5: Bear Case ─────────────────────────────────────────────────────
    with tab5:
        bear = analysis.get("bear_case", {})
        for rk in ["risk_1", "risk_2", "risk_3"]:
            risk = bear.get(rk, {})
            if not risk:
                continue
            st.markdown(f"""
            <div class="risk-card">
                <div class="risk-title">⚠️ {risk.get('category', 'Risk')}</div>
                <p style="margin: 0.5rem 0;">{risk.get('description', 'N/A')}</p>
                <small>
                    <b>Probability:</b> {risk.get('probability', 'N/A')} &nbsp;|&nbsp;
                    <b>Impact:</b> {risk.get('impact', 'N/A')} &nbsp;|&nbsp;
                    <b>Mitigant:</b> {risk.get('mitigant', 'N/A')}
                </small>
            </div>
            """, unsafe_allow_html=True)

    # ── Investment Thesis Footer ──────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**Investment Thesis**")
    st.info(verdict.get("investment_thesis", "N/A"))

    # ── Raw JSON Expander ─────────────────────────────────────────────────────
    with st.expander("🔬 Raw Analysis JSON (Clica aqui se vires falta de dados)"):
        st.json(analysis)

# ── Empty State ───────────────────────────────────────────────────────────────
elif not run_button:
    st.markdown("""
    <div style="text-align:center; padding: 3rem; color: #8A97A8;">
        <div style="font-size: 3rem;">🔬</div>
        <h3 style="color:#1B2A4A;">Institutional Deep Research Engine</h3>
        <p>Enter a stock ticker above and click <b>Run Analysis</b> to generate:</p>
        <p>📄 Full PDF Report &nbsp;·&nbsp; 📊 7-Slide Presentation &nbsp;·&nbsp;
            💰 DCF Model &nbsp;·&nbsp; 📈 Technical Analysis &nbsp;·&nbsp; ⚠️ Bear Case</p>
        <br>
        <p style="font-size:0.8rem;">Powered by yfinance + Gemini 2.5 Flash · For institutional use only</p>
    </div>
    """, unsafe_allow_html=True)