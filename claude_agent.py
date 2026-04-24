import google.generativeai as genai
import json
import os
from dotenv import load_dotenv

load_dotenv()

# Este é o segredo para os números aparecerem: as chaves exatas que o app.py procura
SCHEMA_INSTRUCTION = """
Retorna APENAS um objeto JSON com esta estrutura exata:
{
  "company_summary": {"name": "", "business_model": "", "revenue_streams": [], "moat_factors": [], "moat_assessment": ""},
  "technical_analysis": {"current_price": 0.0, "ma50": 0.0, "ma200": 0.0, "price_vs_ma50_pct": 0.0, "price_vs_ma200_pct": 0.0, "rsi_14": 0.0, "rsi_signal": "", "macd_signal": "", "trend_bias": "", "insider_pattern": "", "insider_summary": ""},
  "dcf_model": {"base_fcf_bn": 0.0, "wacc_pct": 0.0, "terminal_growth_rate_pct": 0.0, "growth_assumptions": {"conservative": 0.0, "base": 0.0, "bull": 0.0}, "sum_pv_fcfs_bn": 0.0, "pv_terminal_value_bn": 0.0, "enterprise_value_bn": 0.0, "net_debt_bn": 0.0, "equity_value_bn": 0.0, "dcf_intrinsic_value": 0.0, "dcf_notes": ""},
  "multiples_analysis": {"subject": {"ev_ebitda": 0.0, "pe_ttm": 0.0, "ps_ttm": 0.0, "revenue_growth_yoy_pct": 0.0, "value_growth_score": 0.0}, "peer_1": {"name": "", "ticker": "", "ev_ebitda": 0.0, "pe_ttm": 0.0, "ps_ttm": 0.0}, "peer_2": {"name": "", "ticker": "", "ev_ebitda": 0.0, "pe_ttm": 0.0, "ps_ttm": 0.0}, "multiples_implied_price": 0.0, "multiples_methodology": ""},
  "bear_case": {"risk_1": {"category": "", "description": "", "probability": "", "impact": "", "mitigant": ""}, "risk_2": {"category": "", "description": "", "probability": "", "impact": "", "mitigant": ""}, "risk_3": {"category": "", "description": "", "probability": "", "impact": "", "mitigant": ""}},
  "verdict": {"current_price": 0.0, "dcf_target_price": 0.0, "blended_target_price": 0.0, "upside_pct": 0.0, "rating": "", "investment_thesis": ""}
}
"""

def analyze_with_claude(raw_data: dict) -> dict:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("Configuração de segurança: GEMINI_API_KEY não encontrada.")

    genai.configure(api_key=api_key)
    
    # Configuramos o modelo para ser estritamente JSON
    model = genai.GenerativeModel(
        model_name="gemini-2.5-flash",
        generation_config={"response_mime_type": "application/json"}
    )

    ticker = raw_data.get("meta", {}).get("ticker", "UNKNOWN")
    prompt = f"Analise os dados financeiros para {ticker}: {json.dumps(raw_data)}\n\n{SCHEMA_INSTRUCTION}"

    try:
        response = model.generate_content(prompt)
        res_text = response.text.strip()
        
        # Limpeza de segurança para garantir que o JSON é puro
        if "```json" in res_text:
            res_text = res_text.split("```json")[1].split("```")[0].strip()
        elif "```" in res_text:
            res_text = res_text.split("```")[1].split("```")[0].strip()
            
        return json.loads(res_text)
    except Exception as e:
        # Se falhar, tentamos uma última vez sem a restrição de MIME type (plano C)
        model_fallback = genai.GenerativeModel(model_name="gemini-2.5-flash")
        response = model_fallback.generate_content(prompt)
        text = response.text.strip()
        if "
