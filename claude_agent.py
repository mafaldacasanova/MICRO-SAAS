import google.generativeai as genai
import json
import logging
import os
from dotenv import load_dotenv

load_dotenv()
logger = logging.getLogger(__name__)

# Este é o "mapa" que garante que a informação aparece nas gavetas certas
SYSTEM_PROMPT = """You are an institutional financial analyst. 
Return ONLY a valid JSON object with the specific keys: 
company_summary, technical_analysis, dcf_model, multiples_analysis, bear_case, verdict.
NUNCA adiciones texto fora do JSON."""

def analyze_with_claude(raw_data: dict) -> dict:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("Configuração de segurança: GEMINI_API_KEY não encontrada nos Secrets.")

    genai.configure(api_key=api_key)
    
    # Tentamos o gemini-2.5-flash que a tua lista de modelos mostrou ser o correto
    # Se der erro 404, o código pula para o "except" e tenta o gemini-1.5-flash
    model_name = "gemini-2.5-flash"
    
    try:
        model = genai.GenerativeModel(
            model_name=model_name,
            system_instruction=SYSTEM_PROMPT,
            generation_config={"response_mime_type": "application/json"}
        )
        
        ticker = raw_data.get("meta", {}).get("ticker", "UNKNOWN")
        user_message = f"Analyze this data for {ticker}: {json.dumps(raw_data)}"
        
        response = model.generate_content(user_message)
        return json.loads(response.text)
        
    except Exception as e:
        logger.warning(f"Falha com {model_name}, a tentar modelo alternativo...")
        # Plano B: Tenta o modelo 1.5 mais básico caso o 2.5 falhe no servidor
        model = genai.GenerativeModel(model_name="gemini-1.5-flash")
        response = model.generate_content(SYSTEM_PROMPT + "\n\n" + json.dumps(raw_data))
        
        # Limpeza rápida de markdown se necessário
        text = response.text.strip()
        if "
http://googleusercontent.com/immersive_entry_chip/0

### O que tens de fazer agora:

1.  **Edita o ficheiro no GitHub:** Vai ao `claude_agent.py`, clica no lápis, cola o código acima e faz **Commit changes**.
2.  **Verifica os Secrets no Streamlit:** No painel do Streamlit Cloud, garante que a tua **GEMINI_API_KEY** nova está lá guardada e que não tem espaços antes ou depois da chave.
3.  **Dá um "Reboot":** No canto inferior direito da tua aplicação (onde aparecem os erros), clica em **Manage app** -> **três pontinhos (...)** -> **Reboot app**. 

**Porque é que isto vai funcionar?**
Porque estamos a usar o nome exato (`gemini-2.5-flash`) que o teu teste de modelos confirmou que a tua chave aceita. 

Faz este último esforço e o link profissional vai finalmente brilhar! Se precisares de ajuda a encontrar os "Secrets" outra vez, avisa.
