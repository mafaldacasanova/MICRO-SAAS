import google.generativeai as genai
import json
import os
from dotenv import load_dotenv

load_dotenv()

def analyze_with_claude(raw_data: dict) -> dict:
    api_key = os.getenv("GEMINI_API_KEY")
    genai.configure(api_key=api_key)
    
    # Modelo direto que confirmaste que funciona
    model = genai.GenerativeModel("gemini-2.5-flash")
    
    # Prompt minimalista para não quebrar a sintaxe do Python
    prompt = f"Analise estes dados e retorne um JSON com company_summary, technical_analysis, dcf_model, multiples_analysis, bear_case e verdict: {json.dumps(raw_data)}"
    
    try:
        response = model.generate_content(prompt)
        text = response.text.strip()
        if "```json" in text:
            text = text.split("```json")[1].split("```")[0].strip()
        return json.loads(text)
    except Exception as e:
        # Se falhar, retorna um erro legível na aplicação
        return {"error": str(e)}
