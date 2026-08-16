import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
import time
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
from docx.shared import RGBColor
import plotly.express as px

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Auditoria e-Protocolo", page_icon="🛡️", layout="wide")

# Estilo Customizado para Celular e Desktop
st.markdown("""
    <style>
    .main-header {background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); padding: 20px; border-radius: 12px; color: white; text-align: center; margin-bottom: 20px;}
    .stButton>button {font-weight: bold; border-radius: 8px;}
    .log-box {background-color: #1e293b; color: #38bdf8; padding: 15px; border-radius: 8px; font-family: monospace; font-size: 13px; height: 300px; overflow-y: auto;}
    </style>
    <div class="main-header">
        <h1 style='color: white;'>🛡️ Monitoramento e-Protocolo</h1>
        <p>Auditoria Corporativa em Nuvem | Modo Background</p>
    </div>
""", unsafe_allow_html=True)

# --- FUNÇÕES AUXILIARES ---
def calcular_dias(data_str):
    try:
        d = re.search(r'\d{2}/\d{2}/\d{4}', str(data_str)).group()
        return max(0, (datetime.now() - datetime.strptime(d, "%d/%m/%Y")).days)
    except Exception: return 0

def e_nao_atribuido(texto): return not texto or texto.strip() == "-"

def analisar_risco(texto, p_custom, r_custom):
    t = str(texto).lower()
    pc = str(p_custom).lower().strip()
    if pc and pc in t: return r_custom
    if any(x in t for x in ['urgente', 'imediato', 'mandado', 'judicial', 'liminar']): return "🔴 ALTO"
    if any(x in t for x in ['solicitação', 'memorando', 'informação']): return "🟡 MÉDIO"
    return "🟢 BAIXO"

def limpa_pdf(texto):
    if not isinstance(texto, str): texto = str(texto)
    return texto.replace('“', '"').replace('”', '"').replace('‘', "'").replace('’', "'").replace('–', '-').replace('—', '-').encode('latin-1', 'replace').decode('latin-1')

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
        return f"Falha na IA (Erro {resposta.status_code})."
    except Exception as e: return f"Erro na IA: {str(e)[:50]}"

# --- INTERFACE LATERAL (SIDEBAR) ---
with st.sidebar:
    st.header("🔐 Acesso Restrito")
    usr_input = st.text_input("CPF do Usuário:", placeholder="Apenas números")
    pwd_input = st.text_input("Senha do Estado:", type="password")
    
    st.divider()
    st.header("🤖 Inteligência Artificial")
    gemini_key = st.text_input("Gemini API Key:", type="password", help="Necessário para os resumos executivos.")
    cb_ia_risco = st.checkbox("Ativar Alerta de Risco (Semântico)", value=True)
    cb_resumo_ia = st.checkbox("Ativar Leitura Profunda de PDFs", value=False)
    
    with st.expander("Parâmetros de Risco"):
        txt_palavra_ia = st.text_input("Palavra Gatilho Externa:", "Mandado")
        dd_prioridade_ia = st.selectbox("Prioridade:", ["🔴 ALTO", "🟡 MÉDIO", "🟢 BAIXO"])

# --- INTERFACE PRINCIPAL (ABAS) ---
tab_filtros, tab_execucao = st.tabs(["⚙️ Filtros e Configurações", "🚀 Execução e Terminal"])

with tab_filtros:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Regras de Varredura")
        filtro_dias = st.number_input("Ignorar protocolos mais novos que (Dias):", min_value=0, value=0)
        alerta_dias = st.number_input("Destacar atrasos maiores que (Dias):", min_value=1, value=30)
        cb_somente_nao_atribuidos = st.checkbox("Varredura focada: Só Não Atribuídos", value=False)
        cb_historico = st.checkbox("Coletar Histórico Completo de Movimentações", value=False)
    
    with col2:
        st.subheader("Formatos de Saída")
        cb_pdf = st.checkbox("Gerar Relatório PDF", value=True)
        cb_excel = st.checkbox("Gerar Planilha Excel", value=True)
        cb_word = st.checkbox("Gerar Relatório Word", value=False)
        ordem_relatorio = st.radio("Ordem do Relatório:", ["Decrescente (Mais antigos primeiro)", "Crescente (Mais novos primeiro)"])

# --- MOTOR DE EXECUÇÃO ---
with tab_execucao:
    st.info("📱 Dica: Após iniciar, você pode colocar o aplicativo em segundo plano. O servidor concluirá o trabalho.")
    
    if st.button("🚀 INICIAR AUDITORIA", type="primary", use_container_width=True):
        if not usr_input or not pwd_input:
            st.error("⚠️ Preencha seu CPF e Senha no menu lateral esquerdo antes de iniciar.")
            st.stop()

        # Placeholders para a UI dinâmica
        progress_bar = st.progress(0)
        status_text = st.empty()
        log_placeholder = st.empty()
        
        logs = []
        def escreve_log(msg, prog=None):
            logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
            # Exibe apenas as últimas 15 linhas para não travar a tela do celular
            html_log = f"<div class='log-box'>{'<br>'.join(logs[-15:])}</div>"
            log_placeholder.markdown(html_log, unsafe_allow_html=True)
            if prog is not None: progress_bar.progress(prog / 100.0)

        escreve_log("Inicializando Servidores em Nuvem...", 5)
        
        # Preparando ambiente de download seguro para o Streamlit Cloud
        download_dir = "/tmp/downloads_eprotocolo"
        os.makedirs(download_dir, exist_ok=True)
        
        chrome_prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "pdfjs.disabled": True,
            "profile.managed_default_content_settings.images": 2 # MODO ULTRA LEVE: Bloqueia imagens
        }
        
        options = Options()
        options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--disable-popup-blocking')
        options.add_experimental_option("prefs", chrome_prefs)
        
        # O Streamlit Cloud instala o chromium automaticamente via packages.txt
        options.binary_location = "/usr/bin/chromium"
        
        try:
            driver = webdriver.Chrome(options=options)
            driver.set_page_load_timeout(60)
            wait = WebDriverWait(driver, 45)
            
            escreve_log("🚪 Acessando portal de Autenticação...", 10)
            try: driver.get("https://auth-cs.identidadedigital.pr.gov.br/centralautenticacao/login.html?response_type=code&client_id=9188905e74c28e489b44e954ec0b9bca&redirect_uri=https%3A%2F%2Fwww.eprotocolo.pr.gov.br%2Fspiweb")
            except: driver.execute_script("window.stop();")
            time.sleep(2)
            
            try:
                driver.execute_script("arguments[0].click();", wait.until(EC.presence_of_element_located((By.ID, "btnCentral"))))
                time.sleep(2)
            except: pass
            
            escreve_log("🔑 Injetando credenciais seguras...", 15)
            usr_f = wait.until(EC.visibility_of_element_located((By.ID, "attribute_central")))
            usr_f.clear(); usr_f.send_keys(usr_input)
            pwd_f = driver.find_element(By.ID, "password")
            pwd_f.clear(); pwd_f.send_keys(pwd_input)
            time.sleep(1)
            
            # Clique fantasma para evitar travamento de cache
            driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btn-central-acessar"))
            
            escreve_log("⏳ Negociando acesso com o e-Protocolo...", 20)
            for _ in range(25):
                if "telaInicial" in driver.current_url or "iniciarProcesso" in driver.current_url: break
                time.sleep(1)
                
            if "telaInicial" not in driver.current_url and "iniciarProcesso" not in driver.current_url:
                try: driver.get("https://www.eprotocolo.pr.gov.br/spiweb/telaInicial.do?action=iniciarProcesso")
                except: driver.execute_script("window.stop();")
                time.sleep(5)
            
            escreve_log("📍 Direcionando para 'Protocolos no Local'...", 25)
            aba_local = wait.until(EC.presence_of_element_located((By.XPATH, "//a[@href='#div_protocolos' or contains(text(), 'Protocolos no Local')]")))
            driver.execute_script("arguments[0].click();", aba_local)
            time.sleep(5)
            
            select_element = wait.until(EC.presence_of_element_located((By.ID, "codLocal")))
            opcoes_locais = [opt.text.strip() for opt in Select(select_element).options if opt.text.strip() and "Selecione" not in opt.text]
            
            janela_principal = driver.current_window_handle
            dados_finais = []
            escreve_log(f"📋 Encontrados {len(opcoes_locais)} locais de trabalho.", 30)
            
            salto_progresso = 60 / max(len(opcoes_locais), 1)
            
            for idx_loc, nome_local in enumerate(opcoes_locais):
                escreve_log(f"🔎 Analisando local: {nome_local}", 30 + (idx_loc * salto_progresso))
                Select(wait.until(EC.presence_of_element_located((By.ID, "codLocal")))).select_by_visible_text(nome_local)
                time.sleep(3) 
                
                try: driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "botaoPesquisar"))
                except: pass
                time.sleep(6) 
                
                total_prot = len(driver.find_elements(By.XPATH, "//table//tbody//tr[.//img[contains(@src, 'icon_exibir.svg')]]"))
                if total_prot == 0: continue
                    
                escreve_log(f"   ✅ {total_prot} processos localizados.")
                
                for idx_prot in range(total_prot):
                    try:
                        wait.until(EC.presence_of_element_located((By.XPATH, "//table//tbody//tr")))
                        time.sleep(2) 
                        linhas = driver.find_elements(By.XPATH, "//table//tbody//tr[.//img[contains(@src, 'icon_exibir.svg')]]")
                        if idx_prot >= len(linhas): continue
                            
                        linha_atual = linhas[idx_prot]
                        try: atribuido_para = linha_atual.find_element(By.XPATH, ".//*[contains(@id, '_atribuidoPara')]").text.strip() or "-"
                        except: atribuido_para = "-"
                        
                        if cb_somente_nao_atribuidos and not e_nao_atribuido(atribuido_para): continue

                        driver.execute_script("arguments[0].click();", linha_atual.find_element(By.XPATH, ".//img[contains(@src, 'icon_exibir.svg')]/ancestor::a"))
                        time.sleep(4) 
                        
                        def pt(xpath):
                            try: return driver.find_element(By.XPATH, xpath).text.strip()
                            except: return "-"
                        
                        try: num_prot = driver.find_element(By.ID, "numeroProtocolo").get_attribute("value")
                        except: 
                            match = re.search(r'\d{2}\.\d{3}\.\d{3}-\d', driver.page_source)
                            num_prot = match.group() if match else "-"
                        
                        data_envio = pt("//div[contains(text(), 'Enviado em:')]/following-sibling::div[1] | //label[contains(text(), 'Enviado em:')]/../following-sibling::div[1]")
                        dias_calculados = calcular_dias(data_envio)
                        
                        if dias_calculados >= filtro_dias:
                            escreve_log(f"      📄 Lendo ID: {num_prot}")
                            
                            dict_dados = {
                                "Local do Filtro": nome_local, "Numero do Eprotocolo": num_prot, "Atribuído para": atribuido_para,
                                "Detalhamento": pt("//div[contains(text(), 'Detalhamento:')]/following-sibling::div[1]"),
                                "Situação": pt("//div[contains(text(), 'Situação:')]/following-sibling::div[1]"),
                                "Local de envio": pt("//div[contains(text(), 'Local de Envio:')]/following-sibling::div[1]"),
                                "Onde esta": pt("//div[contains(text(), 'Onde está:')]/following-sibling::div[1]"),
                                "Motivo": pt("//div[contains(text(), 'Motivo:')]/following-sibling::div[1]"),
                                "Enviado em": data_envio, "Dias no mesmo local": dias_calculados
                            }
                            
                            if cb_ia_risco: dict_dados["Grau de Risco (IA)"] = analisar_risco(dict_dados["Detalhamento"], txt_palavra_ia, dd_prioridade_ia)
                            
                            if cb_resumo_ia:
                                escreve_log(f"      🧠 Baixando anexo para leitura de IA...")
                                try:
                                    for f in os.listdir(download_dir): os.remove(os.path.join(download_dir, f))
                                    btn_download = wait.until(EC.presence_of_element_located((By.XPATH, "//a[.//img[contains(@src, 'icon_download.svg')]] | //img[contains(@src, 'icon_download.svg')]")))
                                    driver.execute_script("arguments[0].click();", btn_download)
                                    
                                    arquivo_baixado = None
                                    for _ in range(45): # timeout de 45seg para download na nuvem
                                        time.sleep(1)
                                        arquivos = os.listdir(download_dir)
                                        arquivos_prontos = [f for f in arquivos if not f.endswith('.crdownload') and not f.endswith('.tmp')]
                                        if arquivos_prontos:
                                            arquivos_prontos.sort(key=lambda x: os.path.getmtime(os.path.join(download_dir, x)), reverse=True)
                                            arquivo_baixado = os.path.join(download_dir, arquivos_prontos[0])
                                            time.sleep(1); break
                                            
                                    if arquivo_baixado:
                                        with pdfplumber.open(arquivo_baixado) as pdf:
                                            total_pages = len(pdf.pages)
                                            # LEITURA EM PINÇA (Otimização Extrema de Memória)
                                            paginas = pdf.pages if total_pages <= 30 else pdf.pages[:15] + pdf.pages[-15:]
                                            texto_documento = "\n".join([page.extract_text() or "" for page in paginas])
                                            
                                        if texto_documento.strip():
                                            dict_dados["Resumo Avançado (IA)"] = gerar_resumo_documento_ia(texto_documento, gemini_key)
                                        else: dict_dados["Resumo Avançado (IA)"] = "Documento sem texto escaneável."
                                    else: dict_dados["Resumo Avançado (IA)"] = "Falha no download (Arquivo muito pesado)."
                                        
                                except Exception: dict_dados["Resumo Avançado (IA)"] = "Erro ao processar anexo."
                                finally:
                                    while len(driver.window_handles) > 1:
                                        driver.switch_to.window(driver.window_handles[-1]); driver.close()
                                    driver.switch_to.window(janela_principal)

                            if cb_historico:
                                try:
                                    aba_andamentos = driver.find_element(By.XPATH, "//h2[contains(., 'Andamentos')]")
                                    driver.execute_script("arguments[0].click();", aba_andamentos)
                                    time.sleep(2)
                                    linhas_andamentos = driver.find_elements(By.XPATH, "//div[@id='Andamentos_menos']//table//tbody//tr")
                                    dict_dados["Movimentações Totais"] = len(linhas_andamentos)
                                except: dict_dados["Movimentações Totais"] = 0
                                    
                            dados_finais.append(dict_dados)
                    except: pass
                    finally:
                        try:
                            if driver.current_window_handle != janela_principal: driver.switch_to.window(janela_principal)
                        except: pass
                        try: driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Voltar'] | //button[contains(text(), 'Voltar')]"))))
                        except: driver.execute_script("window.history.go(-1)")
                        time.sleep(2)
                        
            driver.quit()
            escreve_log("✅ Varredura Concluída com Sucesso!", 95)
            
            # --- PROCESSAMENTO DOS ARQUIVOS E EXIBIÇÃO ---
            if dados_finais:
                escreve_log("📊 Diagramando relatórios para download...", 98)
                fator = -1 if "Decrescente" in ordem_relatorio else 1
                dados_finais.sort(key=lambda x: (x.get('Local do Filtro', ''), fator * x.get('Dias no mesmo local', 0)))
                df = pd.DataFrame(dados_finais)
                
                st.success("🎉 Varredura finalizada! Visualize ou baixe seus dados abaixo.")
                
                # Exibição Interativa dos Dados
                with st.expander("👁️ Visualizar Tabela de Dados", expanded=True):
                    st.dataframe(df, use_container_width=True)
                
                # Geração em Memória (BytesIO) para Download no Web App
                col_btn1, col_btn2, col_btn3 = st.columns(3)
                
                if cb_excel:
                    output_excel = io.BytesIO()
                    df.to_excel(output_excel, index=False)
                    col_btn1.download_button(label="📊 Baixar Planilha Excel", data=output_excel.getvalue(), file_name=f"Auditoria_eProtocolo_{datetime.now().strftime('%d%m%Y')}.xlsx", mime="application/vnd.ms-excel", type="primary", use_container_width=True)
                
                if cb_pdf:
                    pdf = FPDF()
                    pdf.add_page(); pdf.set_font("Arial", 'B', 16)
                    pdf.cell(0, 10, txt=limpa_pdf("Auditoria E-protocolo (Automática)"), ln=True, align='C')
                    pdf.set_font("Arial", 'I', 10)
                    pdf.cell(0, 6, txt=limpa_pdf(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"), ln=True, align='C')
                    
                    locais = {}
                    for p in dados_finais:
                        loc = p['Local do Filtro']
                        if loc not in locais: locais[loc] = []
                        locais[loc].append(p)
                        
                    for loc, group in locais.items():
                        pdf.ln(5)
                        pdf.set_font("Arial", 'B', 11); pdf.set_fill_color(30, 60, 114); pdf.set_text_color(255, 255, 255)
                        pdf.multi_cell(0, 8, limpa_pdf(f" {loc} - ({len(group)} Processos)"), fill=True)
                        for p in group:
                            pdf.set_font("Arial", 'B', 10); pdf.set_text_color(0, 0, 0); pdf.ln(3)
                            pdf.cell(0, 6, limpa_pdf(f"Protocolo: {p.get('Numero do Eprotocolo', '-')} | Dias: {p.get('Dias no mesmo local', 0)}"), 0, 1)
                            pdf.set_font("Arial", '', 10)
                            for k, v in p.items():
                                if k not in ["Numero do Eprotocolo", "Local do Filtro", "Dias no mesmo local"]:
                                    pdf.multi_cell(0, 5, limpa_pdf(f"{k}: {v}"))
                            pdf.set_draw_color(200,200,200); pdf.line(10, pdf.get_y(), 200, pdf.get_y())
                            
                    output_pdf = pdf.output(dest='S').encode('latin-1')
                    col_btn2.download_button(label="📕 Baixar Relatório PDF", data=output_pdf, file_name=f"Auditoria_{datetime.now().strftime('%d%m%Y')}.pdf", mime="application/pdf", type="primary", use_container_width=True)
                
                if cb_word:
                    doc = docx.Document()
                    doc.add_heading('Auditoria E-protocolo', 0).alignment = 1 
                    for p in dados_finais:
                        doc.add_heading(f"Protocolo: {p.get('Numero do Eprotocolo', '-')}", level=2)
                        for k, v in p.items():
                            if k != "Numero do Eprotocolo": doc.add_paragraph(f"{k}: {v}")
                        doc.add_paragraph("_" * 40)
                    output_word = io.BytesIO()
                    doc.save(output_word)
                    col_btn3.download_button(label="📘 Baixar Relatório Word", data=output_word.getvalue(), file_name=f"Auditoria_{datetime.now().strftime('%d%m%Y')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
                
                escreve_log("Tudo pronto! Fim da operação.", 100)
                
            else:
                st.warning("A varredura foi concluída, mas nenhum dado atendeu aos critérios dos filtros.")
                escreve_log("⚠️ Nenhum dado capturado.", 100)

        except Exception as e:
            st.error(f"Ocorreu um erro no motor de varredura. Detalhes no terminal.")
            escreve_log(f"❌ Erro Crítico: {str(e)[:150]}...")
            try: driver.quit() 
            except: pass


