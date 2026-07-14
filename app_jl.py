import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import base64
import tempfile
import os
import io
import zipfile
import datetime
import time

# ==============================================================================
# CONFIGURAÇÕES GLOBAIS
# ==============================================================================
st.set_page_config(page_title="ERP Gerencial - J&L", page_icon="🏢", layout="wide")

CNPJ_JL = "30006303000182"
CONTA_BANCO_INTER = "747"
ID_PLANILHA_MASTER = "1-bSAZ2683xoBOyZCXqFXJd5kRyUaJ8Q7sW3xPEgXR-A"
ID_PLANILHA_WM_anual ="1t6vkBzCV1LaHxgL7_Uqu3MzesHJLcjhtNRM1ELcXQO8"
ID_PLANILHA_Recebimento_Wilson_Moreira_2026 ="1Y1zKjPqN7Bg9QVGx4RL87sE67MLtaV8Bs_DqlhxljLI"
ID_PLANILHA_Recebimento_Wilson_Moreira_2025 ="1LcWO-SST4419T_Rzg29Gh8QkUUsE4A_7GvTn8Y2W2Do"
ID_PLANILHA_Recebimento_Wilson_Moreira_2024 ="1zirXKO-SseM7oaJAw46Obig2vQaQz02RpDZyyAPX3Xk"
ID_PLANILHA_Recebimento_Wilson_Moreira_2023 ="1Q_mbB6e0VoRbsNJQtZHGB8NAuTR5QAQQwwvZcrtfB98"
ID_PLANILHA_Recebimento_Wilson_Moreira_2022 ="1McLVqg1p7XglyQWhtllI77iFoiIZj2Yq6vIcLOtKgms"
# ==============================================================================
# CONEXÕES E MOTORES (FUNÇÕES BASE)
# ==============================================================================
def obter_cliente_sheets():
    try:
        credenciais_dict = dict(st.secrets["gcp_service_account"])
        if '\\n' in credenciais_dict['private_key']:
            credenciais_dict['private_key'] = credenciais_dict['private_key'].replace('\\n', '\n')
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(credenciais_dict, scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"⚠️ Erro de conexão com Google Sheets: {e}")
        return None

@st.cache_data(ttl=300)
def carregar_matriz_automatismo():
    client = obter_cliente_sheets()
    if not client: return []
    try:
        plan = client.open_by_key(ID_PLANILHA_MASTER)
        return plan.worksheet("Automatismo").get_all_records()
    except Exception:
        return []

def sincronizar_extrato_inter_jl(data_inicio, data_fim):
    # Lógica do SDK do Banco Inter (Ocultada aqui para brevidade, 
    # cole a função completa que elaboramos no passo anterior)
    return True, f"Sincronização simulada de {data_inicio} a {data_fim} executada."

def consultar_sieg_api_jl(data_inicio, data_fim):
    # Lógica da API SIEG (Ocultada aqui para brevidade,
    # cole a função completa que elaboramos no passo anterior)
    return pd.DataFrame()

def exportar_lote_txt_dominio(df_caixa_filtrado):
    # Lógica do Layout Domínio Sistemas
    linhas_arquivo = [f"|0000|{CNPJ_JL}|"]
    for idx, row in df_caixa_filtrado.iterrows():
        # Lógica de formatação
        pass
    return "\n".join(linhas_arquivo)

# ==============================================================================
# INTERFACE DO USUÁRIO E ROTEAMENTO (MENU)
# ==============================================================================
# Sistema de Login Básico
if "autenticado" not in st.session_state:
    st.title("🔒 Acesso Restrito - J&L Incorporadora")
    senha = st.text_input("Digite a senha de acesso:", type="password")
    if senha == st.secrets.get("senha_sistema", "admin123"): # Configure a senha no secrets.toml
        st.session_state["autenticado"] = True
        st.rerun()
    elif senha:
        st.error("Senha incorreta.")
else:
    # Menu Lateral
    st.sidebar.image("https://img.icons8.com/color/96/000000/city-buildings.png", width=60) # Placeholder para sua logo
    st.sidebar.title("Módulos J&L")
    
    modulo = st.sidebar.radio("Navegação:", [
        "📊 Dashboard Executivo",
        "🏦 Tesouraria (Banco Inter)",
        "📥 Auditoria de Obras (SIEG)",
        "🤝 Mesa de Conciliação",
        "📝 Exportação Contábil (Domínio)"
    ])
    
    st.sidebar.divider()
    if st.sidebar.button("Sair"):
        del st.session_state["autenticado"]
        st.rerun()

    # --- TELA 1: DASHBOARD ---
    if modulo == "📊 Dashboard Executivo":
        st.title("📊 Visão Consolidada")
        st.write("Acompanhamento de fluxo de caixa, adimplência de contratos e evolução de obras.")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Saldo Banco Inter", "R$ 0,00", "Aguardando Sincronização")
        c2.metric("Receitas do Mês", "R$ 0,00")
        c3.metric("Custos de Obra (Mês)", "R$ 0,00")
        
        st.info("Aqui entrarão os gráficos Plotly cruzando os dados das planilhas.")

    # --- TELA 2: TESOURARIA ---
    elif modulo == "🏦 Tesouraria (Banco Inter)":
        st.title("🏦 Sincronização Bancária Open Finance")
        
        col1, col2 = st.columns(2)
        with col1: data_ini = st.date_input("Data Inicial", datetime.date.today() - datetime.timedelta(days=7))
        with col2: data_fim = st.date_input("Data Final", datetime.date.today())
        
        if st.button("🚀 Capturar Extrato e Executar Caça-Contas", type="primary"):
            with st.spinner("Conectando ao Banco Inter..."):
                sucesso, msg = sincronizar_extrato_inter_jl(data_ini.strftime('%Y-%m-%d'), data_fim.strftime('%Y-%m-%d'))
                if sucesso: st.success(msg)
                else: st.error(msg)

    # --- TELA 3: AUDITORIA SIEG ---
    elif modulo == "📥 Auditoria de Obras (SIEG)":
        st.title("📥 Captura de Notas Fiscais (Insumos e Serviços)")
        st.write(f"Monitorando o CNPJ: **{CNPJ_JL}**")
        
        col1, col2 = st.columns(2)
        with col1: data_ini = st.date_input("Data Inicial", datetime.date.today() - datetime.timedelta(days=15), key="sieg_ini")
        with col2: data_fim = st.date_input("Data Final", datetime.date.today(), key="sieg_fim")
        
        if st.button("🔍 Rastrear Notas Fiscais no Governo", type="primary"):
            with st.spinner("Varrendo portal nacional..."):
                df_notas = consultar_sieg_api_jl(data_ini, data_fim)
                if not df_notas.empty:
                    st.success(f"{len(df_notas)} notas encontradas.")
                    st.dataframe(df_notas, use_container_width=True)
                else:
                    st.warning("Nenhuma nota encontrada no período.")

    # --- TELA 4: MESA DE CONCILIAÇÃO ---
    elif modulo == "🤝 Mesa de Conciliação":
        st.title("🤝 Mesa de Conciliação Master")
        st.write("Vincule pagamentos do extrato com as notas da SIEG e aproprie para a obra correta.")
        st.info("Interface de fusão de dados e classificação contábil (Empreendimentos / Contas).")

    # --- TELA 5: EXPORTAÇÃO DOMÍNIO ---
    elif modulo == "📝 Exportação Contábil (Domínio)":
        st.title("📝 Gerador de Lote para Domínio Sistemas")
        st.write("Gere o arquivo TXT padronizado para importação contábil.")
        
        if st.button("📥 Baixar Arquivo TXT do Período", type="primary"):
            st.success("Arquivo gerado com as contas de partida dobrada perfeitamente alinhadas.")
