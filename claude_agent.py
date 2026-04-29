import google.generativeai as genai
import json
import os
from dotenv import load_dotenv

load_dotenv()

def analyze_with_claude(raw_data: dict) -> dict:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return {"verdict": {"rating": "ERRO: CHAVE NÃO CONFIGURADA"}}

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")

    prompt = f"""
    Analise estes dados financeiros: {json.dumps(raw_data)}
    Responda APENAS um JSON com estas chaves exatas:
    {{
      "company_summary": {{"name": "", "business_model": "", "revenue_streams": [], "moat_factors": [], "moat_assessment": ""}},
      "technical_analysis": {{"current_price": 0.0, "ma50": 0.0, "ma200": 0.0, "price_vs_ma50_pct": 0.0, "price_vs_ma200_pct": 0.0, "rsi_14": 0.0, "rsi_signal": "", "macd_signal": "", "trend_bias": "", "insider_pattern": "", "insider_summary": ""}},
      "dcf_model": {{"base_fcf_bn": 0.0, "wacc_pct": 0.0, "terminal_growth_rate_pct": 0.0, "growth_assumptions": {{"conservative": 0.0, "base": 0.0, "bull": 0.0}}, "sum_pv_fcfs_bn": 0.0, "pv_terminal_value_bn": 0.0, "enterprise_value_bn": 0.0, "net_debt_bn": 0.0, "equity_value_bn": 0.0, "dcf_intrinsic_value": 0.0, "dcf_notes": ""}},
      "multiples_analysis": {{"subject": {{"ev_ebitda": 0.0, "pe_ttm": 0.0, "ps_ttm": 0.0, "revenue_growth_yoy_pct": 0.0, "value_growth_score": 0.0}}, "peer_1": {{"name": "", "ticker": "", "ev_ebitda": 0.0, "pe_ttm": 0.0, "ps_ttm": 0.0}}, "peer_2": {{"name": "", "ticker": "", "ev_ebitda": 0.0, "pe_ttm": 0.0, "ps_ttm": 0.0}}, "multiples_implied_price": 0.0, "multiples_methodology": ""}},
      "bear_case": {{"risk_1": {{"category": "", "description": "", "probability": "", "impact": "", "mitigant": ""}}, "risk_2": {{"category": "", "description": "", "probability": "", "impact": "", "mitigant": ""}}, "risk_3": {{"category": "", "description": "", "probability": "", "impact": "", "mitigant": ""}}}},
      "verdict": {{"current_price": 0.0, "dcf_target_price": 0.0, "blended_target_price": 0.0, "upside_pct": 0.0, "rating": "BUY/HOLD/SELL", "investment_thesis": ""}}
    }}
    """

    try:
        response = model.generate_content(prompt)
        text = response.text.strip()
        if "```json" in text:
            text = text.split("```json")[1].split("```")[0].strip()
        return json.loads(text)
    except Exception as e:
        return { "verdict": { "rating": f"ERRO API: {str(e)[:20]}" } }
