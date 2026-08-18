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
import gc
import base64
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from fpdf import FPDF
import docx

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- 1. CONFIGURAÇÃO E MEMÓRIA ---
st.set_page_config(page_title="Auditoria e-Protocolo", page_icon="🛡️", layout="wide", initial_sidebar_state="expanded")

chaves_padrao = {
    "usr": "", "pwd": "", "gemini_key": "", "cb_ia_risco": True, "cb_resumo_ia": False,
    "txt_palavra_ia": "Mandado", "dd_prioridade_ia": "🔴 ALTO", "filtro_dias": 0, "alerta_dias": 30,
    "cb_somente_nao_atribuidos": False, "cb_historico": False, "cb_pdf": True, "cb_excel": True, 
    "cb_word": False, "ordem_relatorio": "Decrescente",
    "fase_app": "inicio", "dados_auditoria": [], "downloads_feitos": False
}
for k, v in chaves_padrao.items():
    if k not in st.session_state: 
        st.session_state[k] = v

# --- 2. TRUQUE DO DOWNLOAD AUTOMÁTICO (Estilo Colab) ---
def forcar_download_automatico(dados_bytes, nome_arquivo, tipo_mime):
    b64 = base64.b64encode(dados_bytes).decode()
    id_link = re.sub(r'\W+', '', nome_arquivo)
    html_str = f"""
        <a id="{id_link}" href="data:{tipo_mime};base64,{b64}" download="{nome_arquivo}" style="display:none;">AutoDownload</a>
        <script>
            document.getElementById("{id_link}").click();
        </script>
    """
    st.components.v1.html(html_str, height=0, width=0)

# --- 3. DESIGN E ESTILO ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
    [data-testid="stConnectionStatus"] {display: none !important;}
    .main-header {background: linear-gradient(135deg, #4f46e5 0%, #3730a3 100%); padding: 35px 25px; border-radius: 20px; color: white; text-align: center; margin-bottom: 30px; box-shadow: 0 10px 30px -10px rgba(79, 70, 229, 0.4);}
    .main-header h1 {margin: 0; font-size: 2.4em; font-weight: 700; color: #ffffff; letter-spacing: -0.5px;}
    .main-header p {margin: 8px 0 0 0; font-size: 1.1em; color: #e0e7ff; font-weight: 300;}
    div[data-testid="stVerticalBlock"] > div[style*="border"] {background-color: #ffffff !important; border: 1px solid #f1f5f9 !important; border-radius: 20px !important; box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.03) !important; padding: 10px;}
    .log-box {background-color: #f8fafc; color: #475569; padding: 20px; border-radius: 16px; font-family: 'SFMono-Regular', Consolas, monospace; font-size: 13px; height: 300px; overflow-y: auto; border: 1px solid #e2e8f0; box-shadow: inset 0 2px 4px 0 rgba(0,0,0,0.02);}
    .stButton>button {border-radius: 14px; font-weight: 600; padding: 0.6rem 1rem; transition: all 0.2s ease; border: none;}
    .stButton>button[kind="primary"] {background-color: #4f46e5; color: white; font-size: 1.1em; padding: 1rem;}
    .stButton>button[kind="primary"]:hover {background-color: #4338ca; transform: translateY(-1px);}
    .stDownloadButton>button {border-radius: 12px; font-weight: 600; color: #ffffff; background-color: #10b981; border: none; padding: 0.8rem; font-size: 1.05em;}
    .stDownloadButton>button:hover {background-color: #059669;}
    @media (max-width: 768px) { .main-header {padding: 25px 15px;} .main-header h1 {font-size: 1.8em;} .log-box {height: 250px; font-size: 11px;} }
    </style>
    <script>
    async function keepAwake() { if ('wakeLock' in navigator) { try { await navigator.wakeLock.request('screen'); } catch (err) {} } }
    keepAwake(); document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') { keepAwake(); } });
    </script>
    <div class="main-header">
        <h1>🛡️ Auditoria de e-Protocolo</h1>
        <p>Sistema Limpo • Autônomo • Em Nuvem</p>
    </div>
""", unsafe_allow_html=True)

# --- 4. FUNÇÕES DE IA E PROCESSAMENTO ---
def calcular_dias(data_str):
    try: return max(0, (datetime.now() - datetime.strptime(re.search(r'\d{2}/\d{2}/\d{4}', str(data_str)).group(), "%d/%m/%Y")).days)
    except: return 0

def e_nao_atribuido(texto): return not texto or texto.strip() == "-"

def analisar_risco(texto, p_custom, r_custom):
    t = str(texto).lower(); pc = str(p_custom).lower().strip()
    if pc and pc in t: return r_custom
    if any(x in t for x in ['urgente', 'imediato', 'mandado', 'judicial', 'liminar']): return "🔴 ALTO"
    if any(x in t for x in ['solicitação', 'memorando', 'informação']): return "🟡 MÉDIO"
    return "🟢 BAIXO"

def limpa_pdf(texto):
    if not isinstance(texto, str): texto = str(texto)
    return texto.replace('“', '"').replace('”', '"').replace('‘', "'").replace('’', "'").replace('–', '-').replace('—', '-').encode('latin-1', 'replace').decode('latin-1')

def gerar_resumo_documento_ia(texto, chave_api):
    if not chave_api or len(str(texto).strip()) < 10: return "Resumo indisponível."
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={chave_api}"
        payload = {"contents": [{"parts": [{"text": f"Resuma em 1 parágrafo: 1) Assunto; 2) Próximo passo; 3) Último andamento.\n\n{str(texto)[:30000]}"}]}], "generationConfig": {"temperature": 0.3}}
        resp = requests.post(url, headers={'Content-Type': 'application/json'}, json=payload, timeout=20)
        if resp.status_code == 200: return resp.json()['candidates'][0]['content']['parts'][0]['text'].strip().replace('*', '')
        return f"Falha na IA ({resp.status_code})."
    except Exception as e: return f"Erro IA: {str(e)[:40]}"

# --- 5. SIDEBAR ---
with st.sidebar:
    st.markdown("### 💾 Seu Perfil (Local)")
    arquivo_perfil = st.file_uploader("📂 Importar Perfil (.json)", type="json", label_visibility="collapsed")
    if arquivo_perfil:
        try:
            dados_carregados = json.load(arquivo_perfil)
            for k, v in dados_carregados.items():
                # VACINA CONTRA O ERRO DO RÁDIO ANTIGO
                if k == "ordem_relatorio":
                    if "antigos" in str(v).lower() or "Decrescente" in str(v): v = "Decrescente"
                    else: v = "Crescente"
                
                if k in st.session_state: st.session_state[k] = v
            st.success("Perfil carregado!")
        except: st.error("Erro JSON.")
    json_perfil = json.dumps({k: st.session_state[k] for k in chaves_padrao.keys() if k not in ["fase_app", "dados_auditoria", "downloads_feitos"]}, indent=4)
    st.download_button("💾 Baixar Perfil Atual", json_perfil, "meu_perfil_eprotocolo.json", "application/json", use_container_width=True)
    st.divider(); st.markdown("### 🔐 Credenciais")
    st.text_input("CPF do Usuário:", key="usr")
    st.text_input("Senha de Acesso:", key="pwd", type="password")
    st.divider(); st.markdown("### 🤖 Integração I.A.")
    st.text_input("Gemini API Key:", key="gemini_key", type="password")
    st.checkbox("Alerta Semântico (Risco)", key="cb_ia_risco")
    st.checkbox("Resumo de PDFs (Avançado)", key="cb_resumo_ia")
    with st.expander("Parâmetros da I.A."):
        st.text_input("Palavra Gatilho:", key="txt_palavra_ia")
        st.selectbox("Prioridade:", ["🔴 ALTO", "🟡 MÉDIO", "🟢 BAIXO"], key="dd_prioridade_ia")

# =====================================================================
# MÁQUINA DE ESTADOS
# =====================================================================

if st.session_state.fase_app == "inicio":
    st.markdown("### ⚙️ Painel de Configurações")
    with st.container(border=True):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 🎯 Refinamento")
            st.number_input("Ignorar mais novos que (Dias):", min_value=0, key="filtro_dias")
            st.number_input("Destacar atraso acima de (Dias):", min_value=1, key="alerta_dias")
            st.checkbox("Somente Protocolos Não Atribuídos", key="cb_somente_nao_atribuidos")
            st.checkbox("Incluir Histórico de Movimentações", key="cb_historico")
        with col2:
            st.markdown("#### 📑 Relatórios")
            st.checkbox("Gerar Relatório em PDF", key="cb_pdf")
            st.checkbox("Gerar Planilha em Excel", key="cb_excel")
            st.checkbox("Gerar Ofício em Word", key="cb_word")
            st.radio("Organização Cronológica:", ["Decrescente", "Crescente"], key="ordem_relatorio")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 INICIAR AUDITORIA", type="primary", use_container_width=True):
        usr_val = st.session_state.get('usr', '').strip()
        pwd_val = st.session_state.get('pwd', '').strip()
        
        if not usr_val or not pwd_val:
            faltando = []
            if not usr_val: faltando.append("CPF")
            if not pwd_val: faltando.append("Senha")
            st.error(f"⚠️ Atenção: Preencha o(a) {' e '.join(faltando)} no menu lateral esquerdo.")
            st.info("📱 Dica de Celular: Após digitar a senha, toque em 'Concluído/Return' no teclado antes de iniciar.")
            st.stop()
            
        st.session_state.downloads_feitos = False
        st.session_state.fase_app = "processando"
        st.rerun()

elif st.session_state.fase_app == "processando":
    st.markdown("<hr>", unsafe_allow_html=True)
    progress_bar = st.progress(0)
    log_placeholder = st.empty()
    logs = []

    def escreve_log(msg, prog=None):
        logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
        html_log = f"<div class='log-box'>{'<br>'.join(logs[-20:])}</div>"
        log_placeholder.markdown(html_log, unsafe_allow_html=True)
        if prog is not None: progress_bar.progress(prog / 100.0)

    escreve_log("🚀 Despertando Servidores em Nuvem...", 2)
    download_dir = os.path.abspath(os.path.join(os.getcwd(), "downloads_eprotocolo"))
    os.makedirs(download_dir, exist_ok=True)
    
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir, "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True, "pdfjs.disabled": True,
        "profile.managed_default_content_settings.images": 2 
    })
    options.binary_location = "/usr/bin/chromium"
    
    dados_finais = []
    erro_fatal = None
    
    try:
        driver = webdriver.Chrome(options=options)
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': download_dir})
        wait = WebDriverWait(driver, 45)
        
        escreve_log("🚪 Acessando o portal...", 10)
        try: driver.get("https://auth-cs.identidadedigital.pr.gov.br/centralautenticacao/login.html?response_type=code&client_id=9188905e74c28e489b44e954ec0b9bca&redirect_uri=https%3A%2F%2Fwww.eprotocolo.pr.gov.br%2Fspiweb")
        except: driver.execute_script("window.stop();")
        time.sleep(2)
        try: driver.execute_script("arguments[0].click();", wait.until(EC.presence_of_element_located((By.ID, "btnCentral")))); time.sleep(2)
        except: pass
        
        escreve_log("🔑 Injetando credenciais...", 15)
        usr_f = wait.until(EC.visibility_of_element_located((By.ID, "attribute_central")))
        usr_f.clear(); usr_f.send_keys(st.session_state.usr)
        pwd_f = driver.find_element(By.ID, "password")
        pwd_f.clear(); pwd_f.send_keys(st.session_state.pwd)
        driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btn-central-acessar"))
        
        escreve_log("⏳ Negociando Acesso...", 20)
        for _ in range(25):
            if "telaInicial" in driver.current_url or "iniciarProcesso" in driver.current_url: break
            time.sleep(1)
            
        if "telaInicial" not in driver.current_url and "iniciarProcesso" not in driver.current_url:
            try: driver.get("https://www.eprotocolo.pr.gov.br/spiweb/telaInicial.do?action=iniciarProcesso")
            except: driver.execute_script("window.stop();")
            time.sleep(5)
        
        escreve_log("📍 Mapeando 'Protocolos no Local'...", 25)
        driver.execute_script("arguments[0].click();", wait.until(EC.presence_of_element_located((By.XPATH, "//a[@href='#div_protocolos' or contains(text(), 'Protocolos no Local')]"))))
        time.sleep(5)
        
        opcoes_locais = [opt.text.strip() for opt in Select(wait.until(EC.presence_of_element_located((By.ID, "codLocal")))).options if opt.text.strip() and "Selecione" not in opt.text]
        janela_principal = driver.current_window_handle
        escreve_log(f"📋 Encontrados {len(opcoes_locais)} locais.", 30)
        
        salto_progresso = 60 / max(len(opcoes_locais), 1)
        
        for idx_loc, nome_local in enumerate(opcoes_locais):
            escreve_log(f"🔎 Leitura em: {nome_local}", 30 + (idx_loc * salto_progresso))
            Select(wait.until(EC.presence_of_element_located((By.ID, "codLocal")))).select_by_visible_text(nome_local)
            time.sleep(3) 
            try: driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "botaoPesquisar"))
            except: pass
            time.sleep(6) 
            
            total_prot = len(driver.find_elements(By.XPATH, "//table//tbody//tr[.//img[contains(@src, 'icon_exibir.svg')]]"))
            if total_prot == 0: continue
            
            for idx_prot in range(total_prot):
                try:
                    wait.until(EC.presence_of_element_located((By.XPATH, "//table//tbody//tr"))); time.sleep(2) 
                    linhas = driver.find_elements(By.XPATH, "//table//tbody//tr[.//img[contains(@src, 'icon_exibir.svg')]]")
                    if idx_prot >= len(linhas): continue
                        
                    linha_atual = linhas[idx_prot]
                    try: atribuido_para = linha_atual.find_element(By.XPATH, ".//*[contains(@id, '_atribuidoPara')]").text.strip() or "-"
                    except: atribuido_para = "-"
                    
                    if st.session_state.cb_somente_nao_atribuidos and not e_nao_atribuido(atribuido_para): continue
                    driver.execute_script("arguments[0].click();", linha_atual.find_element(By.XPATH, ".//img[contains(@src, 'icon_exibir.svg')]/ancestor::a")); time.sleep(4) 
                    
                    def pt(xpath):
                        try: return driver.find_element(By.XPATH, xpath).text.strip()
                        except: return "-"
                    
                    try: num_prot = driver.find_element(By.ID, "numeroProtocolo").get_attribute("value")
                    except: match = re.search(r'\d{2}\.\d{3}\.\d{3}-\d', driver.page_source); num_prot = match.group() if match else "-"
                    
                    data_envio = pt("//div[contains(text(), 'Enviado em:')]/following-sibling::div[1]")
                    dias_calculados = calcular_dias(data_envio)
                    
                    if dias_calculados >= st.session_state.filtro_dias:
                        escreve_log(f"      📄 Protocolo: {num_prot}")
                        dict_dados = {
                            "Local do Filtro": nome_local, "Numero do Eprotocolo": num_prot, "Atribuído para": atribuido_para,
                            "Detalhamento": pt("//div[contains(text(), 'Detalhamento:')]/following-sibling::div[1]"),
                            "Situação": pt("//div[contains(text(), 'Situação:')]/following-sibling::div[1]"),
                            "Local de envio": pt("//div[contains(text(), 'Local de Envio:')]/following-sibling::div[1]"),
                            "Onde esta": pt("//div[contains(text(), 'Onde está:')]/following-sibling::div[1]"),
                            "Motivo": pt("//div[contains(text(), 'Motivo:')]/following-sibling::div[1]"),
                            "Enviado em": data_envio, "Dias no mesmo local": dias_calculados
                        }
                        
                        if st.session_state.cb_ia_risco: dict_dados["Grau de Risco (IA)"] = analisar_risco(dict_dados["Detalhamento"], st.session_state.txt_palavra_ia, st.session_state.dd_prioridade_ia)
                        
                        if st.session_state.cb_resumo_ia:
                            escreve_log(f"      📥 Baixando anexo para IA...")
                            try:
                                for f in os.listdir(download_dir): os.remove(os.path.join(download_dir, f))
                                driver.execute_script("arguments[0].click();", wait.until(EC.presence_of_element_located((By.XPATH, "//a[.//img[contains(@src, 'icon_download.svg')]] | //img[contains(@src, 'icon_download.svg')]"))))
                                
                                arquivo_baixado = None
                                for _ in range(60):
                                    time.sleep(1)
                                    arquivos = os.listdir(download_dir)
                                    arq_prontos = [f for f in arquivos if not f.endswith('.crdownload') and not f.endswith('.tmp')]
                                    if arq_prontos:
                                        arq_prontos.sort(key=lambda x: os.path.getmtime(os.path.join(download_dir, x)), reverse=True)
                                        arquivo_baixado = os.path.join(download_dir, arq_prontos[0]); time.sleep(1); break
                                        
                                if arquivo_baixado:
                                    escreve_log(f"      📖 Lendo PDF...")
                                    with pdfplumber.open(arquivo_baixado) as pdf:
                                        paginas = pdf.pages if len(pdf.pages) <= 30 else pdf.pages[:15] + pdf.pages[-15:]
                                        texto_documento = "\n".join([page.extract_text() or "" for page in paginas])
                                    dict_dados["Resumo Avançado (IA)"] = gerar_resumo_documento_ia(texto_documento, st.session_state.gemini_key) if texto_documento.strip() else "Documento sem OCR."
                                else: dict_dados["Resumo Avançado (IA)"] = "Timeout no Download."
                            except Exception: dict_dados["Resumo Avançado (IA)"] = "Documento Restrito."
                            finally:
                                while len(driver.window_handles) > 1: driver.switch_to.window(driver.window_handles[-1]); driver.close()
                                driver.switch_to.window(janela_principal)

                        if st.session_state.cb_historico:
                            try:
                                driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//h2[contains(., 'Andamentos')]")); time.sleep(2)
                                dict_dados["Movimentações Totais"] = len(driver.find_elements(By.XPATH, "//div[@id='Andamentos_menos']//table//tbody//tr"))
                            except: dict_dados["Movimentações Totais"] = 0
                                
                        dados_finais.append(dict_dados)
                except: pass
                finally:
                    try:
                        if driver.current_window_handle != janela_principal: driver.switch_to.window(janela_principal)
                        driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Voltar'] | //button[contains(text(), 'Voltar')]"))))
                    except: driver.execute_script("window.history.go(-1)")
                    time.sleep(2); gc.collect() 
                    
        escreve_log("✨ Finalizando...", 100)
        
    except Exception as e: erro_fatal = str(e)
    finally:
        try: driver.quit() 
        except: pass
        
        st.session_state.dados_auditoria = dados_finais
        st.session_state.fase_app = "erro" if erro_fatal and not dados_finais else "concluido"
        st.rerun()

elif st.session_state.fase_app == "concluido":
    st.success("🎉 Processamento concluído com sucesso!")
    
    df = pd.DataFrame(st.session_state.dados_auditoria)
    
    if len(df) > 0:
        fator = -1 if "Decrescente" in st.session_state.ordem_relatorio else 1
        df = df.sort_values(by=['Local do Filtro', 'Dias no mesmo local'], ascending=[True, fator == 1])
        
        with st.expander("👁️ Pré-visualização de Dados", expanded=True): 
            st.dataframe(df, use_container_width=True)
            
        st.markdown("### 📥 Seus Relatórios")
        col_btn1, col_btn2, col_btn3 = st.columns(3)
        
        if st.session_state.cb_excel:
            output_excel = io.BytesIO()
            df.to_excel(output_excel, index=False)
            dados_excel = output_excel.getvalue()
            col_btn1.download_button("📊 Baixar Excel", dados_excel, "Auditoria_eProtocolo.xlsx", "application/vnd.ms-excel", use_container_width=True)
            if not st.session_state.downloads_feitos:
                forcar_download_automatico(dados_excel, "Auditoria_eProtocolo.xlsx", "application/vnd.ms-excel")
        
        if st.session_state.cb_pdf:
            pdf = FPDF(); pdf.add_page(); pdf.set_font("Arial", 'B', 16); pdf.cell(0, 10, txt="Auditoria E-protocolo", ln=True, align='C')
            locais = {}
            for _, p in df.iterrows(): locais.setdefault(p['Local do Filtro'], []).append(p)
            for loc, group in locais.items():
                pdf.ln(5); pdf.set_font("Arial", 'B', 11); pdf.set_fill_color(79, 70, 229); pdf.set_text_color(255, 255, 255); pdf.multi_cell(0, 8, limpa_pdf(f" SETOR: {loc} - ({len(group)} Processos)"), fill=True)
                for p in group:
                    pdf.set_font("Arial", 'B', 10); pdf.set_text_color(204, 0, 0) if p.get('Dias no mesmo local', 0) >= st.session_state.alerta_dias else pdf.set_text_color(0, 0, 0)
                    pdf.ln(3); pdf.cell(0, 6, limpa_pdf(f"Protocolo: {p.get('Numero do Eprotocolo', '-')} | Dias Parado: {p.get('Dias no mesmo local', 0)}"), 0, 1)
                    pdf.set_font("Arial", '', 10); pdf.set_text_color(50, 50, 50)
                    for k, v in p.items():
                        if k not in ["Numero do Eprotocolo", "Local do Filtro", "Dias no mesmo local"]: pdf.multi_cell(0, 5, limpa_pdf(f"{k}: {v}"))
                    pdf.set_draw_color(226,232,240); pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            
            dados_pdf = pdf.output(dest='S').encode('latin-1')
            col_btn2.download_button("📕 Baixar PDF", dados_pdf, "Auditoria_eProtocolo.pdf", "application/pdf", use_container_width=True)
            if not st.session_state.downloads_feitos:
                forcar_download_automatico(dados_pdf, "Auditoria_eProtocolo.pdf", "application/pdf")
        
        if st.session_state.cb_word:
            doc = docx.Document(); doc.add_heading('Auditoria E-protocolo', 0).alignment = 1 
            for _, p in df.iterrows():
                doc.add_heading(f"ID: {p.get('Numero do Eprotocolo', '-')}", level=2)
                for k, v in p.items():
                    if k != "Numero do Eprotocolo": doc.add_paragraph(f"{k}: {v}")
            output_word = io.BytesIO(); doc.save(output_word)
            dados_word = output_word.getvalue()
            col_btn3.download_button("📘 Baixar Word", dados_word, "Auditoria_eProtocolo.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            if not st.session_state.downloads_feitos:
                forcar_download_automatico(dados_word, "Auditoria_eProtocolo.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        st.session_state.downloads_feitos = True 

    else:
        st.warning("Nenhum processo atendeu aos critérios estabelecidos.")

    st.markdown("<br><hr>", unsafe_allow_html=True)
    if st.button("🔄 Fazer Nova Auditoria", type="secondary", use_container_width=True):
        st.session_state.fase_app = "inicio"
        st.session_state.dados_auditoria = []
        st.rerun()

elif st.session_state.fase_app == "erro":
    st.error("❌ Ocorreu uma instabilidade no servidor (provavelmente falta de memória) e a coleta foi interrompida.")
    if st.button("🔄 Voltar e Tentar Novamente", type="secondary", use_container_width=True):
        st.session_state.fase_app = "inicio"
        st.rerun()
