"""
pdf_generator.py
Gera o relatório PDF institucional com base nos dados do Claude.
"""

import io
import logging
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY

logger = logging.getLogger(__name__)

# ── Paleta de Cores ──────────────────────────────────────────────────────────
DARK_NAVY    = colors.HexColor("#0D1B2A")
MID_NAVY     = colors.HexColor("#1B2A4A")
ACCENT_BLUE  = colors.HexColor("#1F5EFF")
LIGHT_GREY   = colors.HexColor("#F4F6FA")
MID_GREY     = colors.HexColor("#8A97A8")
WHITE        = colors.white
RED_RISK     = colors.HexColor("#C0392B")
GREEN_OK     = colors.HexColor("#1A7A4A")
AMBER        = colors.HexColor("#D4A017")
TABLE_HEADER = colors.HexColor("#1B2A4A")
TABLE_ALT    = colors.HexColor("#EEF1F7")


def _build_styles():
    base = getSampleStyleSheet()
    styles = {
        "title": ParagraphStyle(
            "ReportTitle",
            parent=base["Title"],
            fontName="Helvetica-Bold",
            fontSize=26,
            textColor=WHITE,
            alignment=TA_LEFT,
            spaceAfter=4,
        ),
        "subtitle": ParagraphStyle(
            "Subtitle",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=12,
            textColor=colors.HexColor("#A0B4CC"),
            alignment=TA_LEFT,
            spaceAfter=2,
        ),
        "h1": ParagraphStyle(
            "H1",
            parent=base["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=14,
            textColor=DARK_NAVY,
            spaceBefore=14,
            spaceAfter=4,
            borderPad=0,
        ),
        "h2": ParagraphStyle(
            "H2",
            parent=base["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=11,
            textColor=MID_NAVY,
            spaceBefore=10,
            spaceAfter=3,
        ),
        "body": ParagraphStyle(
            "BodyText",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=9.5,
            textColor=DARK_NAVY,
            leading=15,
            alignment=TA_JUSTIFY,
            spaceAfter=6,
        ),
        "body_grey": ParagraphStyle(
            "BodyGrey",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=9,
            textColor=MID_GREY,
            leading=14,
            spaceAfter=4,
        ),
        "label": ParagraphStyle(
            "Label",
            parent=base["Normal"],
            fontName="Helvetica-Bold",
            fontSize=9,
            textColor=MID_NAVY,
        ),
        "risk_title": ParagraphStyle(
            "RiskTitle",
            parent=base["Normal"],
            fontName="Helvetica-Bold",
            fontSize=10,
            textColor=RED_RISK,
            spaceBefore=6,
            spaceAfter=2,
        ),
        "verdict_box": ParagraphStyle(
            "VerdictBox",
            parent=base["Normal"],
            fontName="Helvetica-Bold",
            fontSize=13,
            textColor=WHITE,
            alignment=TA_CENTER,
        ),
        "footer": ParagraphStyle(
            "Footer",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=7.5,
            textColor=MID_GREY,
            alignment=TA_CENTER,
        ),
    }
    return styles


def _header_footer(canvas, doc):
    """Desenha header e footer em cada página."""
    canvas.saveState()
    W, H = A4

    # Header bar (apenas páginas 2+)
    if doc.page > 1:
        canvas.setFillColor(DARK_NAVY)
        canvas.rect(0, H - 1.2 * cm, W, 1.2 * cm, fill=1, stroke=0)
        canvas.setFillColor(WHITE)
        canvas.setFont("Helvetica-Bold", 9)
        canvas.drawString(1.5 * cm, H - 0.8 * cm, doc._report_title)
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(W - 1.5 * cm, H - 0.8 * cm, doc._report_date)

    # Footer
    canvas.setFillColor(MID_GREY)
    canvas.setFont("Helvetica", 7.5)
    canvas.drawString(1.5 * cm, 0.7 * cm,
        "CONFIDENTIAL — For Institutional Use Only. "
        "This report does not constitute investment advice.")
    canvas.drawRightString(
        W - 1.5 * cm, 0.7 * cm, f"Page {doc.page}"
    )
    canvas.restoreState()


def _make_table(data, col_widths, header_row=True):
    """Cria tabela formatada."""
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), TABLE_HEADER),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("ALIGN", (0, 1), (0, -1), "LEFT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, TABLE_ALT]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D0D8E8")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle(style))
    return t


def _fmt(val, suffix="", prefix="", decimals=2, na="N/A"):
    """Formata valores numéricos com fallback para N/A."""
    if val is None:
        return na
    try:
        return f"{prefix}{float(val):.{decimals}f}{suffix}"
    except (TypeError, ValueError):
        return str(val) if val else na


def _rating_color(rating: str):
    r = (rating or "").upper()
    if r == "BUY":
        return GREEN_OK
    elif r == "SELL":
        return RED_RISK
    return AMBER


def generate_pdf(analysis: dict, ticker: str) -> bytes:
    """
    Gera o PDF completo e retorna os bytes prontos para download.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=2.0 * cm,
        bottomMargin=1.5 * cm,
    )
    doc._report_title = f"{ticker} — Institutional Research Report"
    doc._report_date = datetime.utcnow().strftime("%d %b %Y")

    S = _build_styles()
    W_page = A4[0] - 3 * cm  # usable width
    story = []

    # ── 1. CAPA ──────────────────────────────────────────────────────────────
    company = analysis.get("company_summary", {})
    verdict = analysis.get("verdict", {})
    rating = verdict.get("rating", "N/A")
    rating_color = _rating_color(rating)

    # Capa background
    cover_data = [[
        Paragraph(f"<b>{company.get('name', ticker)}</b>", S["title"]),
    ]]
    cover_table = Table(
        [[
            Paragraph(f"<font color='white'><b>{company.get('name', ticker)}</b></font>", S["title"]),
        ]],
        colWidths=[W_page]
    )
    cover_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), DARK_NAVY),
        ("LEFTPADDING", (0, 0), (-1, -1), 20),
        ("RIGHTPADDING", (0, 0), (-1, -1), 20),
        ("TOPPADDING", (0, 0), (-1, -1), 28),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
    ]))
    story.append(cover_table)

    # Metadata row
    meta_data = [[
        Paragraph(f"<font color='#A0B4CC'>Ticker</font>", S["body_grey"]),
        Paragraph(f"<font color='#A0B4CC'>Sector</font>", S["body_grey"]),
        Paragraph(f"<font color='#A0B4CC'>Report Date</font>", S["body_grey"]),
        Paragraph(f"<font color='#A0B4CC'>Rating</font>", S["body_grey"]),
    ],[
        Paragraph(f"<b>{ticker}</b>", S["h2"]),
        Paragraph(f"{company.get('sector', 'N/A')}", S["h2"]),
        Paragraph(f"{doc._report_date}", S["h2"]),
        Paragraph(f"<font color='{rating_color.hexval()}'><b>{rating}</b></font>", S["h2"]),
    ]]
    mt = Table(meta_data, colWidths=[W_page/4]*4)
    mt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LIGHT_GREY),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D0D8E8")),
    ]))
    story.append(mt)
    story.append(Spacer(1, 12))

    # Target price highlight
    tp = _fmt(verdict.get("blended_target_price"), prefix="$")
    cp = _fmt(verdict.get("current_price"), prefix="$")
    up = _fmt(verdict.get("upside_pct"), suffix="%")
    kpi_data = [["Current Price", "Target Price (Blended)", "Upside / Downside", "Rating"],
                [cp, tp, up, rating]]
    kpi_t = Table(kpi_data, colWidths=[W_page/4]*4)
    kpi_t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), TABLE_HEADER),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("BACKGROUND", (3, 1), (3, 1), rating_color),
        ("TEXTCOLOR", (3, 1), (3, 1), WHITE),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D0D8E8")),
    ]))
    story.append(kpi_t)
    story.append(Spacer(1, 8))

    # Investment Thesis
    thesis = verdict.get("investment_thesis", "")
    if thesis:
        story.append(Paragraph("Investment Thesis", S["h2"]))
        story.append(Paragraph(thesis, S["body"]))

    story.append(HRFlowable(width="100%", thickness=0.5, color=MID_GREY, spaceAfter=6))
    story.append(PageBreak())

    # ── 2. BUSINESS MODEL & MOAT ─────────────────────────────────────────────
    story.append(Paragraph("1. Business Model & Competitive Moat", S["h1"]))
    biz = analysis.get("company_summary", {})

    story.append(Paragraph("Business Description", S["h2"]))
    story.append(Paragraph(biz.get("business_model", "N/A"), S["body"]))

    rev_streams = biz.get("revenue_streams", [])
    if rev_streams:
        story.append(Paragraph("Revenue Streams", S["h2"]))
        for rs in rev_streams:
            story.append(Paragraph(f"• {rs}", S["body"]))

    moat_factors = biz.get("moat_factors", [])
    if moat_factors:
        story.append(Paragraph("Identified Moat Factors", S["h2"]))
        for mf in moat_factors:
            story.append(Paragraph(f"• {mf}", S["body"]))

    moat_assessment = biz.get("moat_assessment", "")
    if moat_assessment:
        story.append(Paragraph("Moat Durability Assessment", S["h2"]))
        story.append(Paragraph(moat_assessment, S["body"]))

    story.append(PageBreak())

    # ── 3. ANÁLISE TÉCNICA & INSIDERS ────────────────────────────────────────
    story.append(Paragraph("2. Technical Analysis & Insider Activity", S["h1"]))
    tech = analysis.get("technical_analysis", {})

    tech_data = [
        ["Metric", "Value", "Signal"],
        ["Current Price", _fmt(tech.get("current_price"), prefix="$"), "—"],
        ["MA 50", _fmt(tech.get("ma50"), prefix="$"),
         f"{_fmt(tech.get('price_vs_ma50_pct'), suffix='%')} vs price"],
        ["MA 200", _fmt(tech.get("ma200"), prefix="$"),
         f"{_fmt(tech.get('price_vs_ma200_pct'), suffix='%')} vs price"],
        ["RSI (14)", _fmt(tech.get("rsi_14")), tech.get("rsi_signal", "N/A")],
        ["MACD Signal", "—", tech.get("macd_signal", "N/A")],
        ["Trend Bias", "—", tech.get("trend_bias", "N/A")],
        ["Insider Pattern", "—", tech.get("insider_pattern", "N/A")],
    ]
    story.append(_make_table(tech_data, [W_page * 0.35, W_page * 0.3, W_page * 0.35]))
    story.append(Spacer(1, 8))

    insider_summary = tech.get("insider_summary", "")
    if insider_summary:
        story.append(Paragraph("Insider Transaction Commentary", S["h2"]))
        story.append(Paragraph(insider_summary, S["body"]))

    story.append(PageBreak())

    # ── 4. MODELO DCF ────────────────────────────────────────────────────────
    story.append(Paragraph("3. Discounted Cash Flow Valuation (Mid-Year Convention)", S["h1"]))
    dcf = analysis.get("dcf_model", {})

    # Assumptions
    assump_data = [
        ["Parameter", "Value"],
        ["Base FCF (Year 0)", _fmt(dcf.get("base_fcf_bn"), suffix="B")],
        ["WACC", _fmt(dcf.get("wacc_pct"), suffix="%")],
        ["Terminal Growth Rate", _fmt(dcf.get("terminal_growth_rate_pct"), suffix="%")],
        ["Projection Period", f"{dcf.get('projection_years', 5)} years"],
        ["Growth — Conservative", _fmt(dcf.get("growth_assumptions", {}).get("conservative"), suffix="%")],
        ["Growth — Base Case", _fmt(dcf.get("growth_assumptions", {}).get("base"), suffix="%")],
        ["Growth — Bull Case", _fmt(dcf.get("growth_assumptions", {}).get("bull"), suffix="%")],
    ]
    story.append(Paragraph("Model Assumptions", S["h2"]))
    story.append(_make_table(assump_data, [W_page * 0.55, W_page * 0.45]))
    story.append(Spacer(1, 10))

    # FCF Projections
    proj = dcf.get("fcf_projections_bn", [])
    pvs  = dcf.get("pv_fcfs_bn", [])
    if proj:
        story.append(Paragraph("FCF Projections & Present Values (Base Case, $B)", S["h2"]))
        n = len(proj)
        headers = ["Year"] + [f"Y{i+1}" for i in range(n)]
        fcf_row = ["Projected FCF ($B)"] + [_fmt(v) for v in proj]
        pv_row  = ["PV of FCF ($B)"]   + [_fmt(v) for v in pvs]
        proj_data = [headers, fcf_row, pv_row]
        col_w = [W_page * 0.28] + [(W_page * 0.72) / n] * n
        story.append(_make_table(proj_data, col_w))
        story.append(Spacer(1, 10))

    # Valuation Bridge
    bridge_data = [
        ["Valuation Component", "Value ($B)"],
        ["Sum PV of FCFs", _fmt(dcf.get("sum_pv_fcfs_bn"), suffix="B")],
        ["Terminal Value (undiscounted)", _fmt(dcf.get("terminal_value_bn"), suffix="B")],
        ["PV of Terminal Value", _fmt(dcf.get("pv_terminal_value_bn"), suffix="B")],
        ["Enterprise Value", _fmt(dcf.get("enterprise_value_bn"), suffix="B")],
        ["Less: Net Debt", _fmt(dcf.get("net_debt_bn"), suffix="B")],
        ["Equity Value", _fmt(dcf.get("equity_value_bn"), suffix="B")],
        ["Shares Outstanding (B)", _fmt(dcf.get("shares_outstanding_bn"), decimals=3)],
        ["DCF Intrinsic Value per Share", _fmt(dcf.get("dcf_intrinsic_value"), prefix="$", decimals=2)],
    ]
    story.append(Paragraph("DCF Valuation Bridge", S["h2"]))
    bridge_t = _make_table(bridge_data, [W_page * 0.6, W_page * 0.4])
    story.append(bridge_t)

    notes = dcf.get("dcf_notes", "")
    if notes:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<i>Notes: {notes}</i>", S["body_grey"]))

    story.append(PageBreak())

    # ── 5. MÚLTIPLOS COMPARATIVOS ────────────────────────────────────────────
    story.append(Paragraph("4. Comparable Company Analysis", S["h1"]))
    mult = analysis.get("multiples_analysis", {})
    subj = mult.get("subject", {})
    p1   = mult.get("peer_1", {})
    p2   = mult.get("peer_2", {})

    comp_data = [
        ["Metric", ticker, f"{p1.get('name','Peer 1')}\n({p1.get('ticker','')})",
         f"{p2.get('name','Peer 2')}\n({p2.get('ticker','')})"],
        ["EV/EBITDA", _fmt(subj.get("ev_ebitda")), _fmt(p1.get("ev_ebitda")), _fmt(p2.get("ev_ebitda"))],
        ["P/E (TTM)",  _fmt(subj.get("pe_ttm")),   _fmt(p1.get("pe_ttm")),   _fmt(p2.get("pe_ttm"))],
        ["P/S (TTM)",  _fmt(subj.get("ps_ttm")),   _fmt(p1.get("ps_ttm")),   _fmt(p2.get("ps_ttm"))],
        ["YoY Revenue Growth", _fmt(subj.get("revenue_growth_yoy_pct"), suffix="%"), "—", "—"],
        ["Value/Growth Score", _fmt(subj.get("value_growth_score"), decimals=2), "—", "—"],
    ]
    story.append(_make_table(comp_data, [W_page*0.28, W_page*0.24, W_page*0.24, W_page*0.24]))
    story.append(Spacer(1, 8))

    story.append(Paragraph(
        f"Multiples-Implied Price: <b>{_fmt(mult.get('multiples_implied_price'), prefix='$')}</b>",
        S["h2"]
    ))
    methodology = mult.get("multiples_methodology", "")
    if methodology:
        story.append(Paragraph(methodology, S["body"]))

    story.append(PageBreak())

    # ── 6. BEAR CASE ─────────────────────────────────────────────────────────
    story.append(Paragraph("5. Bear Case — Key Risk Factors", S["h1"]))
    bear = analysis.get("bear_case", {})

    for risk_key in ["risk_1", "risk_2", "risk_3"]:
        risk = bear.get(risk_key, {})
        if not risk:
            continue
        story.append(KeepTogether([
            Paragraph(f"Risk: {risk.get('category', 'N/A')}", S["risk_title"]),
            Paragraph(risk.get("description", "N/A"), S["body"]),
        ]))

        risk_meta = [
            ["Probability", "Impact", "Mitigant"],
            [risk.get("probability", "N/A"),
             risk.get("impact", "N/A"),
             risk.get("mitigant", "N/A")],
        ]
        story.append(_make_table(risk_meta, [W_page*0.2, W_page*0.2, W_page*0.6]))
        story.append(Spacer(1, 10))

    story.append(PageBreak())

    # ── 7. VEREDITO FINAL ────────────────────────────────────────────────────
    story.append(Paragraph("6. Verdict & Price Target", S["h1"]))

    verdict_data = [
        ["Component", "Target Price", "Weight"],
        ["DCF Intrinsic Value", _fmt(verdict.get("dcf_target_price"), prefix="$"), "50%"],
        ["Multiples-Implied Price", _fmt(verdict.get("multiples_target_price"), prefix="$"), "50%"],
        ["Blended Target Price", _fmt(verdict.get("blended_target_price"), prefix="$"), "100%"],
        ["Current Price", _fmt(verdict.get("current_price"), prefix="$"), "—"],
        ["Implied Upside / Downside", _fmt(verdict.get("upside_pct"), suffix="%"), "—"],
    ]
    verdict_t = _make_table(verdict_data, [W_page*0.45, W_page*0.3, W_page*0.25])
    story.append(verdict_t)
    story.append(Spacer(1, 14))

    # Rating box
    r_color = _rating_color(rating)
    rating_data = [[Paragraph(f"<b>RATING: {rating}</b>",
                              ParagraphStyle("R", fontName="Helvetica-Bold",
                                             fontSize=16, textColor=WHITE,
                                             alignment=TA_CENTER))]]
    rating_box = Table(rating_data, colWidths=[W_page])
    rating_box.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), r_color),
        ("TOPPADDING", (0, 0), (-1, -1), 14),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    story.append(rating_box)
    story.append(Spacer(1, 12))

    if thesis:
        story.append(Paragraph("Investment Thesis", S["h2"]))
        story.append(Paragraph(thesis, S["body"]))

    story.append(Spacer(1, 20))
    story.append(HRFlowable(width="100%", thickness=0.5, color=MID_GREY))
    story.append(Spacer(1, 6))
    disclaimer = (
        "DISCLAIMER: This report is generated for informational and research purposes only. "
        "It does not constitute investment advice, a solicitation, or a recommendation to buy or sell any security. "
        "Past performance is not indicative of future results. All projections are inherently uncertain. "
        "This document is confidential and intended solely for institutional use."
    )
    story.append(Paragraph(disclaimer, S["footer"]))

    # Build
    doc.build(story, onFirstPage=_header_footer, onLaterPages=_header_footer)
    return buf.getvalue()
