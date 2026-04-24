import google.generativeai as genai
import json
import logging
import os
from dotenv import load_dotenv

# Isto carrega o .env localmente, mas no Streamlit Cloud ele vai ler os "Secrets"
load_dotenv()
logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """You are an institutional-grade financial analyst. 
Return ONLY a valid JSON object with the exact schema requested. 
NUNCA alteres o nome das chaves (company_summary, technical_analysis, dcf_model, etc)."""

def analyze_with_claude(raw_data: dict) -> dict:
    # AQUI ESTÁ O SEGREDO: 
    # Ele tenta ler a chave das configurações do Streamlit primeiro.
    api_key = os.getenv("GEMINI_API_KEY")
    
    if not api_key:
        logger.error("ERRO: GEMINI_API_KEY não encontrada!")
        raise RuntimeError("A chave API não foi configurada nos Secrets do Streamlit.")

    genai.configure(api_key=api_key)
    
    model = genai.GenerativeModel(
        model_name="gemini-1.5-flash", # Versão estável e rápida
        system_instruction=SYSTEM_PROMPT,
        generation_config={"response_mime_type": "application/json"}
    )

    ticker = raw_data.get("meta", {}).get("ticker", "UNKNOWN")
    user_message = f"Analyze these financial data for {ticker}:\n{json.dumps(raw_data)}"

    try:
        response = model.generate_content(user_message)
        return json.loads(response.text)
    except Exception as e:
        logger.error(f"Erro na API: {str(e)}")
        raise RuntimeError(f"Falha na API do Gemini: {str(e)}")
