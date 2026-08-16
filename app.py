import streamlit as st

# Configuração inicial da página
st.set_page_config(
    page_title="Auditoria Corporativa",
    page_icon="🔎",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Estilo visual do banner azul topo
st.markdown("""
    <style>
    .banner-azul {
        background: linear-gradient(135deg, #0052cc, #003d99);
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
    }
    .banner-azul h2 {
        margin: 0;
        font-weight: 600;
        font-size: 1.5rem;
    }
    </style>
""", unsafe_allow_html=True)

# Cria a memória do aplicativo para saber se o usuário já colocou a senha
if 'logado' not in st.session_state:
    st.session_state['logado'] = False

# Tela de Login Principal
def tela_login():
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("<h2 style='text-align: center; color: #0052cc;'>🔒 Acesso Restrito</h2>", unsafe_allow_html=True)
        st.markdown("---")
        
        cpf = st.text_input("CPF:")
        senha = st.text_input("Senha:", type="password")
        
        if st.button("Entrar", use_container_width=True):
            # Validação da senha (mude o 123456 para a senha desejada)
            if cpf != "" and senha == "123456": 
                st.session_state['logado'] = True
                st.rerun()
            else:
                st.error("CPF ou Senha incorretos.")

# Tela Principal da Auditoria (Aparece após o login)
def tela_principal():
    # Menu lateral escondido
    with st.sidebar:
        st.header("🤖 Configuração IA")
        gemini_key = st.text_input("Gemini API Key:", type="password")
        
        st.checkbox("Ativar IA de Urgência")
        st.checkbox("Ativar Resumo de PDF (Avançado)")
        
        st.markdown("---")
        if st.button("Sair / Logout", use_container_width=True):
            st.session_state['logado'] = False
            st.rerun()

    # Banner Azul
    st.markdown("""
        <div class="banner-azul">
            <h2>Auditoria Corporativa em Nuvem | Execução em Background</h2>
        </div>
    """, unsafe_allow_html=True)

    # Abas da tela
    aba1, aba2 = st.tabs(["⚙️ Filtros e Relatórios", "🚀 Execução e Terminal"])

    # Aba de Filtros
    with aba1:
        st.subheader("Filtros de Varredura")
        
        col_filtro1, col_filtro2 = st.columns(2)
        with col_filtro1:
            ignorar_dias = st.number_input("Ignorar protocolos < que (Dias):", min_value=0, value=0, step=1)
        with col_filtro2:
            alerta_vermelho = st.number_input("Alerta Vermelho a partir de (Dias):", min_value=0, value=30, step=1)

        st.checkbox("Somente Não Atribuídos")
        st.checkbox("Coletar Histórico de Movimentações")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.subheader("Formatos de Saída")
        col_pdf, col_excel, col_word = st.columns(3)
        with col_pdf:
            gerar_pdf = st.checkbox("PDF", value=True)
        with col_excel:
            gerar_excel = st.checkbox("Excel", value=True)
        with col_word:
            gerar_word = st.checkbox("Word")
            
        st.markdown("<br>", unsafe_allow_html=True)
        ordem = st.radio("Ordem:", ["Decrescente", "Crescente"])

    # Aba de Terminal
    with aba2:
        st.write("A tela de execução e os resultados do terminal aparecerão aqui...")

# Lógica que decide qual tela mostrar
if not st.session_state['logado']:
    tela_login()
else:
    tela_principal()
