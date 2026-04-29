import streamlit as st
import pandas as pd
import os
import json
from dotenv import load_dotenv
from data_fetcher import fetch_all_data
from claude_agent import analyze_with_claude
from pdf_generator import generate_pdf
from pptx_generator import generate_pptx

load_dotenv()

st.set_page_config(page_title="DeepResearch AI", layout="wide")

st.title("📊 DeepResearch: Análise Institucional")

ticker = st.text_input("Introduza o Ticker (ex: NVDA, AAPL):", "").upper()

if st.button("Executar Análise"):
    if ticker:
        with st.spinner(f"A recolher dados de {ticker}..."):
            try:
                # 1. Busca dados do Yahoo Finance
                raw_data = fetch_all_data(ticker)
                
                # 2. IA analisa os dados
                analysis = analyze_with_claude(raw_data)
                
                if "verdict" in analysis:
                    v = analysis["verdict"]
                    
                    # Mostrar Quadrados de Resumo (os que estavam N/A)
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("Preço Atual", f"${v.get('current_price', 0)}")
                    col2.metric("Alvo DCF", f"${v.get('dcf_target_price', 0)}")
                    col3.metric("Upside", f"{v.get('upside_pct', 0)}%")
                    col4.subheader(f"Rating: {v.get('rating', 'N/A')}")
                    
                    st.success("Análise completa!")
                    
                    # Botões de Download
                    pdf_bytes = generate_pdf(analysis, ticker)
                    st.download_button("Descarregar PDF", pdf_bytes, f"{ticker}_Report.pdf")
                    
                    pptx_bytes = generate_pptx(analysis, ticker)
                    st.download_button("Descarregar PPTX", pptx_bytes, f"{ticker}_Presentation.pptx")
                else:
                    st.error("A IA não conseguiu estruturar os dados.")
                    
            except Exception as e:
                st.error(f"Erro: {e}")
    else:
        st.warning("Por favor, insira um ticker.")
