import google.generativeai as genai
import json
import os
from dotenv import load_dotenv

load_dotenv()

def analyze_with_claude(raw_data: dict) -> dict:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return {"error": "Chave API não configurada"}

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")

    # Este prompt obriga a IA a usar as chaves que o teu ecrã (a imagem) precisa
    prompt = f"""
    Analise estes dados financeiros: {json.dumps(raw_data)}
    Responda APENAS um JSON com estas chaves exatas para preencher o ecrã:
    {{
      "company_summary": {{"name": "NVIDIA", "business_model": "", "revenue_streams": [], "moat_factors": [], "moat_assessment": ""}},
      "technical_analysis": {{"current_price": 0.0, "rsi_14": 0.0, "rsi_signal": ""}},
      "dcf_model": {{"dcf_intrinsic_value": 0.0}},
      "multiples_analysis": {{"multiples_implied_price": 0.0}},
      "verdict": {{
        "current_price": 120.0, 
        "dcf_target_price": 150.0, 
        "blended_target_price": 145.0, 
        "upside_pct": 25.0, 
        "rating": "BUY"
      }}
    }}
    """

    try:
        response = model.generate_content(prompt)
        text = response.text.strip()
        # Limpeza de segurança se a IA colocar ```json
        if "```json" in text:
            text = text.split("```json")[1].split("```")[0].strip()
        elif "```" in text:
            text = text.split("```")[1].split("```")[0].strip()
        return json.loads(text)
    except Exception:
        # Se falhar, retorna uma estrutura mínima para não aparecer N/A
        return {
            "verdict": {"current_price": 0.0, "dcf_target_price": 0.0, "blended_target_price": 0.0, "upside_pct": 0.0, "rating": "Erro na API"}
        }
