import google.generativeai as genai
import json
import logging
import os
from dotenv import load_dotenv

load_dotenv()
logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """You are an institutional-grade financial analyst at a quantitative hedge fund.
Your mandate is strictly analytical. You present data, calculate metrics, and identify risks with surgical precision.
You will receive a JSON payload with raw financial data for a publicly listed company.
Your task is to produce a deep-research analysis in a specific JSON structure.

OUTPUT FORMAT:
Return ONLY a valid JSON object with the following exact schema. No markdown, no preamble, no extra keys. NUNCA alteres o nome destas chaves.

{
  "company_summary": {
    "name": "",
    "ticker": "",
    "business_model": "",
    "revenue_streams": [],
    "moat_factors": [],
    "moat_assessment": ""
  },
  "technical_analysis": {
    "current_price": 0.0,
    "ma50": 0.0,
    "ma200": 0.0,
    "price_vs_ma50_pct": 0.0,
    "price_vs_ma200_pct": 0.0,
    "rsi_14": 0.0,
    "rsi_signal": "",
    "macd_signal": "",
    "trend_bias": "",
    "insider_pattern": "",
    "insider_summary": ""
  },
  "dcf_model": {
    "base_fcf_bn": 0.0,
    "wacc_pct": 0.0,
    "terminal_growth_rate_pct": 0.0,
    "growth_assumptions": {
      "conservative": 0.0,
      "base": 0.0,
      "bull": 0.0
    },
    "sum_pv_fcfs_bn": 0.0,
    "pv_terminal_value_bn": 0.0,
    "enterprise_value_bn": 0.0,
    "net_debt_bn": 0.0,
    "equity_value_bn": 0.0,
    "dcf_intrinsic_value": 0.0,
    "dcf_notes": ""
  },
  "multiples_analysis": {
    "subject": {
      "ev_ebitda": 0.0,
      "pe_ttm": 0.0,
      "ps_ttm": 0.0,
      "revenue_growth_yoy_pct": 0.0,
      "value_growth_score": 0.0
    },
    "peer_1": { "name": "", "ticker": "", "ev_ebitda": 0.0, "pe_ttm": 0.0, "ps_ttm": 0.0 },
    "peer_2": { "name": "", "ticker": "", "ev_ebitda": 0.0, "pe_ttm": 0.0, "ps_ttm": 0.0 },
    "multiples_implied_price": 0.0,
    "multiples_methodology": ""
  },
  "bear_case": {
    "risk_1": { "category": "Fraud / Accounting", "description": "", "probability": "", "impact": "", "mitigant": "" },
    "risk_2": { "category": "Customer Concentration", "description": "", "probability": "", "impact": "", "mitigant": "" },
    "risk_3": { "category": "Competitive Displacement", "description": "", "probability": "", "impact": "", "mitigant": "" }
  },
  "verdict": {
    "current_price": 0.0,
    "dcf_target_price": 0.0,
    "blended_target_price": 0.0,
    "upside_pct": 0.0,
    "rating": "BUY/HOLD/SELL",
    "investment_thesis": ""
  }
}"""

def _build_user_prompt(raw_data: dict) -> str:
    ticker = raw_data.get("meta", {}).get("ticker", "UNKNOWN")
    return (
        f"Analyze these financial data for {ticker}.\n"
        f"DATA (JSON):\n{json.dumps(raw_data, indent=2)}\n\n"
        "Return ONLY the structured JSON exactly as requested in the SYSTEM PROMPT."
    )

def analyze_with_claude(raw_data: dict) -> dict:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        api_key = "AIzaSyBBCbu5coCK10XvBHhYU8knjWNsoyovAs4"

    genai.configure(api_key=api_key)
    
    # Motor super rápido 2.5 com a garantia de cuspir sempre o JSON no formato exato!
    model = genai.GenerativeModel(
        model_name="gemini-2.5-flash",
        system_instruction=SYSTEM_PROMPT,
        generation_config={"response_mime_type": "application/json"}
    )

    user_message = _build_user_prompt(raw_data)
    logger.info("A enviar dados para o Gemini 2.5 Flash...")

    try:
        response = model.generate_content(user_message)
        raw_text = response.text.strip()
        
        if "```json" in raw_text:
            raw_text = raw_text.split("```json")[1].split("```")[0].strip()
        elif "```" in raw_text:
            raw_text = raw_text.split("```")[1].split("```")[0].strip()

        return json.loads(raw_text)
        
    except Exception as e:
        logger.error(f"Erro na API: {str(e)}")
        raise RuntimeError(f"Falha na API do Gemini: {str(e)}")