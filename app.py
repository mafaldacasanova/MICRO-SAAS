import streamlit as st
import pandas as pd
from data_fetcher import fetch_all_data
from claude_agent import analyze_with_claude

load_dotenv()

def analyze_with_claude(raw_data: dict) -> dict:
    api_key = os.getenv("GEMINI_API_KEY")
    # Se não houver chave, ele avisa no ecrã em vez de ficar branco
    if not api_key:
        return {"verdict": {"rating": "ERRO: FALTA CHAVE"}}

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")

    prompt = f"""
    Analise estes dados e responda APENAS um JSON: {json.dumps(raw_data)}
    Use estas chaves: company_summary, technical_analysis, dcf_model, multiples_analysis, bear_case, verdict.
    """

    try:
        response = model.generate_content(prompt)
        text = response.text.strip()
        if "```json" in text:
            text = text.split("```json")[1].split("```")[0].strip()
        elif "```" in text:
            text = text.split("```")[1].split("```")[0].strip()
        return json.loads(text)
    except Exception as e:
        # Se a IA falhar, ele retorna a estrutura para a app não ficar branca
        return {
            "company_summary": {"name": "Erro", "business_model": str(e)},
            "verdict": {"rating": "ERRO NA API", "current_price": 0, "upside_pct": 0}
        }
