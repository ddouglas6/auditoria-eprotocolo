import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
import time
import json
import os
import requests
import urllib3
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from fpdf import FPDF
import docx
from docx.shared import Pt, RGBColor
import plotly.express as px

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Auditoria e-Protocolo", page_icon="🛡️", layout="wide")

st.markdown("""
    <style>
    .main-header {background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); padding: 20px; border-radius: 10px; color: white; text-align: center; margin-bottom: 20px;}
    </style>
    <div class="main-header">
        <h1>🛡️ Sistema de Monitoramento e-Protocolo</h1>
        <p>Auditoria Corporativa em Nuvem | Execução em Background</p>
    </div>
""", unsafe_allow_html=True)

# --- FUNÇÕES AUXILIARES ---
def calcular_dias(data_str):
    try:
        d = re.search(r'\d{2}/\d{2}/\d{4}', str(data_str)).group()
        return max(0, (datetime.now() - datetime.strptime(d, "%d/%m/%Y")).days)
    except Exception: return 0

def e_nao_atribuido(texto): return not texto or texto.strip() == "-"
def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
def analisar_risco(texto, p_custom, r_custom):
    t = str(texto).lower()
    pc = str(p_custom).lower().strip()
    if pc and pc in t: return r_custom
    if any(x in t for x in ['urgente', 'imediato', 'mandado', 'judicial', 'liminar']): return "🔴 ALTO"
    if any(x in t for x in ['solicitação', 'memorando', 'informação']): return "🟡 MÉDIO"
    return "🟢 BAIXO"
def limpa_pdf(texto):
    if not isinstance(texto, str): texto = str(texto)
    texto = texto.replace('“', '"').replace('”', '"').replace('‘', "'").replace('’', "'")
    return texto.replace('–', '-').replace('—', '-').replace('º', 'o.').replace('ª', 'a.').encode('latin-1', 'replace').decode('latin-1')

def gerar_resumo_documento_ia(texto, chave_api):
    if not chave_api or len(str(texto).strip()) < 10: return "Resumo não disponível."
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={chave_api}"
        headers = {'Content-Type': 'application/json'}
        prompt = f"Você é um auditor rigoroso. Crie um resumo MUITO CONCISO (1 parágrafo) do documento abaixo. Destaque: 1) Assunto; 2) Próximo passo; 3) Último andamento.\n\nDocumento:\n{str(texto)[:30000]}"
        payload = {"contents": [{"parts": [{"text": prompt}]}], "generationConfig": {"temperature": 0.3}}
        resposta = requests.post(url, headers=headers, json=payload, timeout=20)
        if resposta.status_code == 200:
            return resposta.json()['candidates'][0]['content']['parts'][0]['text'].strip().replace('*', '')
        return f"Falha de conexão com a IA (Erro {resposta.status_code})."
    except Exception as e: return f"Erro de processamento da IA: {str(e)[:50]}"

# --- INTERFACE DO APLICATIVO ---
with st.sidebar:
    st.header("🔐 Acesso")
    usr_input = st.text_input("CPF:", type="password")
    pwd_input = st.text_input("Senha:", type="password")
    st.divider()
    st.header("🤖 Configuração IA")
    gemini_key = st.text_input("Gemini API Key:", type="password")
    cb_ia_risco = st.checkbox("Ativar IA de Urgência", value=True)
    cb_resumo_ia = st.checkbox("Ativar Resumo de PDF (Avançado)", value=False)
    
tab1, tab2 = st.tabs(["⚙️ Filtros e Relatórios", "🚀 Execução e Terminal"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Filtros de Varredura")
        filtro_dias = st.number_input("Ignorar protocolos < que (Dias):", min_value=0, value=0)
        alerta_dias = st.number_input("Alerta Vermelho a partir de (Dias):", min_value=1, value=30)
        cb_somente_nao_atribuidos = st.checkbox("Somente Não Atribuídos")
        cb_historico = st.checkbox("Coletar Histórico de Movimentações")
    with col2:
        st.subheader("Formatos de Saída")
        cb_pdf = st.checkbox("PDF", value=True)
        cb_excel = st.checkbox("Excel", value=True)
        cb_word = st.checkbox("Word", value=False)
        ordem_relatorio = st.radio("Ordem:", ["Decrescente", "Crescente"])

with tab2:
    if st.button("🚀 INICIAR VARREDURA COMPLETA", use_container_width=True, type="primary"):
        if not usr_input or not pwd_input:
            st.error("Preencha CPF e Senha no menu lateral!")
            st.stop()
            
        st.info("💡 Dica para celular: Você pode bloquear a tela agora. O servidor fará o trabalho.")
        status_text = st.empty()
        progress_bar = st.progress(0)
        log_box = st.empty()
        
        logs = []
        def escreve_log(msg, prog=None):
            agora = datetime.now().strftime('%H:%M:%S')
            logs.append(f"[{agora}] {msg}")
            # Mantém apenas os últimos 15 logs na tela para não travar o celular
            log_box.code("\n".join(logs[-15:]), language="text")
            if prog is not None: progress_bar.progress(prog / 100.0)
            
        escreve_log("🚀 Iniciando Motor de Auditoria...", 5)
        
        # CONFIGURAÇÕES DO CHROME PARA STREAMLIT CLOUD
        download_dir = "/tmp/downloads"
        os.makedirs(download_dir, exist_ok=True)
        chrome_prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "pdfjs.disabled": True
        }
        
        options = Options()
        options.page_load_strategy = 'eager'
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')
        options.add_experimental_option("prefs", chrome_prefs)
        # O Streamlit instala o chromium aqui:
        options.binary_location = "/usr/bin/chromium" 
        
        try:
            driver = webdriver.Chrome(options=options)
            wait = WebDriverWait(driver, 35)
            
            escreve_log("🚪 Acessando portal PR...", 10)
            try: driver.get("https://auth-cs.identidadedigital.pr.gov.br/centralautenticacao/login.html?response_type=code&client_id=9188905e74c28e489b44e954ec0b9bca&redirect_uri=https%3A%2F%2Fwww.eprotocolo.pr.gov.br%2Fspiweb")
            except: driver.execute_script("window.stop();")
            
            time.sleep(2)
            try: driver.execute_script("arguments[0].click();", wait.until(EC.presence_of_element_located((By.ID, "btnCentral")))); time.sleep(2)
            except: pass
            
            escreve_log("🔑 Injetando credenciais...", 15)
            usr_f = wait.until(EC.visibility_of_element_located((By.ID, "attribute_central")))
            usr_f.send_keys(usr_input)
            driver.find_element(By.ID, "password").send_keys(pwd_input)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btn-central-acessar"))
            
            escreve_log("⏳ Aguardando Sistema...", 20)
            time.sleep(8)
            if "telaInicial" not in driver.current_url and "iniciarProcesso" not in driver.current_url:
                try: driver.get("https://www.eprotocolo.pr.gov.br/spiweb/telaInicial.do?action=iniciarProcesso")
                except: driver.execute_script("window.stop();")
                time.sleep(5)
                
            escreve_log("📍 Direcionando para Locais...", 25)
            aba_local = wait.until(EC.presence_of_element_located((By.XPATH, "//a[@href='#div_protocolos' or contains(text(), 'Protocolos no Local')]")))
            driver.execute_script("arguments[0].click();", aba_local)
            time.sleep(4)
            
            sel = Select(wait.until(EC.presence_of_element_located((By.ID, "codLocal"))))
            opcoes_locais = [o.text.strip() for o in sel.options if o.text.strip() and "Selecione" not in o.text]
            
            janela_principal = driver.current_window_handle
            dados_finais = []
            escreve_log(f"📋 Mapeados {len(opcoes_locais)} locais para varredura.", 30)
            
            passo = 60 / max(len(opcoes_locais), 1)
            
            for i, loc in enumerate(opcoes_locais):
                escreve_log(f"🔎 Lendo: {loc}", 30 + (i * passo))
                Select(wait.until(EC.presence_of_element_located((By.ID, "codLocal")))).select_by_visible_text(loc)
                time.sleep(3)
                try: driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "botaoPesquisar"))
                except: pass
                time.sleep(6)
                
                prots = len(driver.find_elements(By.XPATH, "//table//tbody//tr[.//img[contains(@src, 'icon_exibir.svg')]]"))
                if prots == 0: continue
                escreve_log(f"   ✅ {prots} protocolos achados.")
                
                for p_idx in range(prots):
                    wait.until(EC.presence_of_element_located((By.XPATH, "//table//tbody//tr")))
                    time.sleep(2)
                    linhas = driver.find_elements(By.XPATH, "//table//tbody//tr[.//img[contains(@src, 'icon_exibir.svg')]]")
                    if p_idx >= len(linhas): continue
                    
                    linha = linhas[p_idx]
                    try: atrib = linha.find_element(By.XPATH, ".//*[contains(@id, '_atribuidoPara')]").text.strip() or "-"
                    except: atrib = "-"
                    
                    if cb_somente_nao_atribuidos and not e_nao_atribuido(atrib): continue
                        
                    driver.execute_script("arguments[0].click();", linha.find_element(By.XPATH, ".//img[contains(@src, 'icon_exibir.svg')]/ancestor::a"))
                    time.sleep(4)
                    
                    def pt(xp):
                        try: return driver.find_element(By.XPATH, xp).text.strip()
                        except: return "-"
                    try: num = driver.find_element(By.ID, "numeroProtocolo").get_attribute("value")
                    except: 
                        m = re.search(r'\d{2}\.\d{3}\.\d{3}-\d', driver.page_source)
                        num = m.group() if m else "-"
                        
                    data_e = pt("//div[contains(text(), 'Enviado em:')]/following-sibling::div[1] | //label[contains(text(), 'Enviado em:')]/../following-sibling::div[1]")
                    dias = calcular_dias(data_e)
                    
                    if dias >= filtro_dias:
                        escreve_log(f"      📄 {num} ({p_idx+1}/{prots})")
                        dict_d = {
                            "Local do Filtro": loc, "Numero do Eprotocolo": num, "Atribuído para": atrib,
                            "Detalhamento": pt("//div[contains(text(), 'Detalhamento:')]/following-sibling::div[1]"),
                            "Situação": pt("//div[contains(text(), 'Situação:')]/following-sibling::div[1]"),
                            "Local de envio": pt("//div[contains(text(), 'Local de Envio:')]/following-sibling::div[1]"),
                            "Onde esta": pt("//div[contains(text(), 'Onde está:')]/following-sibling::div[1]"),
                            "Motivo": pt("//div[contains(text(), 'Motivo:')]/following-sibling::div[1]"),
                            "Enviado em": data_e, "Dias no mesmo local": dias
                        }
                        if cb_ia_risco: dict_d["Grau de Risco (IA)"] = analisar_risco(dict_d["Detalhamento"], "", "🔴 ALTO")
                        
                        if cb_resumo_ia:
                            escreve_log(f"      🧠 Baixando PDF para IA...")
                            try:
                                for f in os.listdir(download_dir): os.remove(os.path.join(download_dir, f))
                                btn_d = wait.until(EC.presence_of_element_located((By.XPATH, "//img[contains(@src, 'icon_download.svg')]/ancestor::a | //img[contains(@src, 'icon_download.svg')]")))
                                driver.execute_script("arguments[0].click();", btn_d)
                                
                                arq = None
                                for _ in range(30): # Espera 30 seg na nuvem
                                    time.sleep(1)
                                    fs = [f for f in os.listdir(download_dir) if not f.endswith('.tmp') and not f.endswith('.crdownload')]
                                    if fs:
                                        arq = os.path.join(download_dir, fs[0]); time.sleep(1); break
                                        
                                if arq:
                                    with pdfplumber.open(arq) as pdf:
                                        pgs = pdf.pages if len(pdf.pages) <= 30 else pdf.pages[:15] + pdf.pages[-15:]
                                        txt = "\n".join([p.extract_text() or "" for p in pgs])
                                    if txt.strip(): dict_d["Resumo Avançado (IA)"] = gerar_resumo_documento_ia(txt, gemini_key)
                                    else: dict_d["Resumo Avançado (IA)"] = "Documento s/ texto."
                                else: dict_d["Resumo Avançado (IA)"] = "Falha no download/Timeout."
                            except: dict_d["Resumo Avançado (IA)"] = "Acesso Restrito/Falha."
                            finally:
                                while len(driver.window_handles) > 1:
                                    driver.switch_to.window(driver.window_handles[-1]); driver.close()
                                driver.switch_to.window(janela_principal)
                                
                        if cb_historico:
                            try:
                                driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//h2[contains(., 'Andamentos')]"))
                                time.sleep(2)
                                dict_d["Movimentações"] = len(driver.find_elements(By.XPATH, "//div[@id='Andamentos_menos']//table//tbody//tr"))
                            except: dict_d["Movimentações"] = 0
                            
                        dados_finais.append(dict_d)
                        
                    try: driver.execute_script("arguments[0].click();", wait.until(EC.presence_of_element_located((By.XPATH, "//input[@value='Voltar'] | //button[contains(text(), 'Voltar')]"))))
                    except: driver.execute_script("window.history.go(-1)")
                    time.sleep(2)
                    
            driver.quit()
            escreve_log("✅ Varredura Concluída!", 95)
            
            if dados_finais:
                escreve_log("📊 Gerando Relatórios...", 98)
                fator = -1 if ordem_relatorio == "Decrescente" else 1
                dados_finais.sort(key=lambda x: (x.get('Local do Filtro', ''), fator * x.get('Dias no mesmo local', 0)))
                df = pd.DataFrame(dados_finais)
                
                st.subheader("Pré-visualização dos Dados")
                st.dataframe(df)
                
                # --- BOTÕES DE DOWNLOAD ---
                col_d1, col_d2, col_d3 = st.columns(3)
                
                if cb_excel:
                    buffer_exc = io.BytesIO()
                    df.to_excel(buffer_exc, index=False)
                    col_d1.download_button(label="📥 Baixar Planilha EXCEL", data=buffer_exc.getvalue(), file_name="Auditoria.xlsx", mime="application/vnd.ms-excel", type="primary")
                
                if cb_pdf:
                    pdf = FPDF()
                    pdf.add_page(); pdf.set_font("Arial", 'B', 16)
                    pdf.cell(0, 10, txt="Auditoria E-protocolo", ln=True, align='C')
                    for p in dados_finais:
                        pdf.set_font("Arial", 'B', 10); pdf.ln(3)
                        pdf.cell(0, 6, limpa_pdf(f"Protocolo: {p.get('Numero do Eprotocolo', '-')} | Local: {p.get('Local do Filtro', '-')}"), 0, 1)
                        pdf.set_font("Arial", '', 10)
                        for k, v in p.items():
                            if k not in ["Numero do Eprotocolo", "Local do Filtro"]:
                                pdf.multi_cell(0, 5, limpa_pdf(f"{k}: {v}"))
                        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
                    buffer_pdf = pdf.output(dest='S').encode('latin-1')
                    col_d2.download_button(label="📥 Baixar Relatório PDF", data=buffer_pdf, file_name="Auditoria.pdf", mime="application/pdf", type="primary")

                escreve_log("🎉 SUCESSO! Relatórios Prontos para Download.", 100)
            else:
                escreve_log("⚠️ Nenhum dado capturado.", 100)
                
        except Exception as e:
            escreve_log(f"❌ Erro Crítico: {str(e)[:150]}")
            try: driver.quit()
            except: pass
