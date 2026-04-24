import google.generativeai as genai
import json
import os
from dotenv import load_dotenv

load_dotenv()

def analyze_with_claude(raw_data: dict) -> dict:
    # Procura a chave guardada nos Secrets do Streamlit
    api_key = os.getenv("GEMINI_API_KEY")
    
    if not api_key:
        raise RuntimeError("Erro de Configuração: GEMINI_API_KEY não encontrada.")

    genai.configure(api_key=api_key)
    
    # Usamos o modelo que o teu teste confirmou: gemini-2.5-flash
    model = genai.GenerativeModel(model_name="gemini-2.5-flash")

    # Prompt simplificado para evitar erros de aspas
    user_message = f"Analise os seguintes dados financeiros e retorne apenas um JSON estruturado: {json.dumps(raw_data)}"

    try:
        response = model.generate_content(user_message)
        res_text = response.text.strip()
        
        # Limpa blocos de código se a IA os incluir
        if "```json" in res_text:
            res_text = res_text.split("```json")[1].split("```")[0].strip()
        elif "```" in res_text:
            res_text = res_text.split("```")[1].split("```")[0].strip()
            
        return json.loads(res_text)
    except Exception as e:
        raise RuntimeError(f"Erro na análise: {str(e)}")
