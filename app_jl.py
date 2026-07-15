import streamlit as st
import os
import json
import re
import zipfile
import io
import base64
import tempfile
import datetime
import time
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests

# ==========================================
# 1. CONFIGURAÇÕES GLOBAIS - J&L INCORPORADORA
# ==========================================
st.set_page_config(page_title="ERP Gerencial - J&L", page_icon="🏢", layout="wide")

CNPJ_JL = "30006303000182"
EMAIL_SIEG = "empresarialdldp@gmail.com"

# --- IDs DAS PLANILHAS ---
ID_PLANILHA_MASTER = "1-bSAZ2683xoBOyZCXqFXJd5kRyUaJ8Q7sW3xPEgXR-A"
ID_PLANILHA_WM_anual ="1t6vkBzCV1LaHxgL7_Uqu3MzesHJLcjhtNRM1ELcXQO8"
ID_PLANILHA_Recebimento_Wilson_Moreira_2026 ="1Y1zKjPqN7Bg9QVGx4RL87sE67MLtaV8Bs_DqlhxljLI"
ID_PLANILHA_Recebimento_Wilson_Moreira_2025 ="1LcWO-SST4419T_Rzg29Gh8QkUUsE4A_7GvTn8Y2W2Do"
ID_PLANILHA_Recebimento_Wilson_Moreira_2024 ="1zirXKO-SseM7oaJAw46Obig2vQaQz02RpDZyyAPX3Xk"
ID_PLANILHA_Recebimento_Wilson_Moreira_2023 ="1Q_mbB6e0VoRbsNJQtZHGB8NAuTR5QAQQwwvZcrtfB98"
ID_PLANILHA_Recebimento_Wilson_Moreira_2022 ="1McLVqg1p7XglyQWhtllI77iFoiIZj2Yq6vIcLOtKgms"

# ==========================================
# EXTRAÇÃO AUTOMÁTICA DO SDK DO BANCO INTER
# ==========================================
if not os.path.exists('inter_sdk_python') and os.path.exists('inter_sdk_python.zip'):
    with zipfile.ZipFile('inter_sdk_python.zip', 'r') as zip_ref:
        zip_ref.extractall('.')
    st.toast("📦 SDK do Banco Inter extraído com sucesso na nuvem!", icon="⚙️")

# ==========================================
# 2. CONEXÕES E MOTORES BLINDADOS
# ==========================================
def obter_cliente_sheets():
    """Conecta ao Google Sheets resolvendo as quebras de linha da chave JSON."""
    try:
        credenciais_dict = dict(st.secrets["gcp_service_account"])
        if '\\n' in credenciais_dict['private_key']:
            credenciais_dict['private_key'] = credenciais_dict['private_key'].replace('\\n', '\n')
            
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(credenciais_dict, scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Erro na autenticação do Google Sheets: {e}")
        return None

def extrair_dados_xml_rapido(conteudo_xml):
    """Extrator otimizado via regex para Notas Fiscais (NFe e NFSe)."""
    def limpar_texto(texto):
        if not texto: return ""
        return re.sub(r'[\r\n\t]+', ' ', str(texto)).strip()

    dados = {
        "Tipo": "DESPESA", "Numero_NF": "", "Chave_Acesso": "", 
        "Emitente": "NÃO IDENTIFICADO", "CNPJ_Emitente": "",
        "Destinatario": "NÃO IDENTIFICADO", "CNPJ_Destinatario": "",
        "Data_Emissao": "", "Valor_Total": 0.0, "Itens_Comprados": []
    }
    
    def buscar_tag(nome_tag, texto):
        match = re.search(rf'<[^>]*?\b{nome_tag}\b[^>]*>([^<]+)</[^>]*?\b{nome_tag}\b>', texto, re.IGNORECASE)
        return limpar_texto(match.group(1)) if match else None

    # Numero e Chave
    dados["Numero_NF"] = buscar_tag(r'nNF', conteudo_xml) or buscar_tag(r'Numero', conteudo_xml)
    dados["Chave_Acesso"] = buscar_tag(r'chNFe', conteudo_xml)
    
    if not dados["Chave_Acesso"]:
        match_id = re.search(r'<infNFe[^>]*?Id="NFe(\d{44})"', conteudo_xml, re.IGNORECASE)
        if match_id: dados["Chave_Acesso"] = limpar_texto(match_id.group(1))

    # Valores e Datas
    data = buscar_tag(r'dhEmi', conteudo_xml) or buscar_tag(r'DataEmissao', conteudo_xml)
    if data: dados["Data_Emissao"] = data[:10]

    valor = buscar_tag(r'vNF', conteudo_xml) or buscar_tag(r'vServ', conteudo_xml)
    if valor: dados["Valor_Total"] = float(valor.replace(',', '.'))

    # Emitente (Fornecedor)
    cnpj_emit = re.search(r'<emit>.*?<CNPJ>(\d+)</CNPJ>', conteudo_xml, re.IGNORECASE | re.DOTALL)
    if cnpj_emit: dados["CNPJ_Emitente"] = cnpj_emit.group(1)
    
    nome_emit = re.search(r'<emit>.*?<xNome>([^<]+)</xNome>', conteudo_xml, re.IGNORECASE | re.DOTALL)
    if nome_emit: dados["Emitente"] = limpar_texto(nome_emit.group(1))

    # Identificação Automática (Se for a J&L emitindo, é Receita)
    if dados["CNPJ_Emitente"] == CNPJ_JL:
        dados["Tipo"] = "RECEITA"
        
    return dados

def consultar_sieg_api(data_inicio, data_fim):
    """Busca notas na SIEG (Entradas e Saídas) para o CNPJ da J&L."""
    try:
        client_id = str(st.secrets["sieg_api"]["CLIENT_ID"]).strip()
        secret_key = str(st.secrets["sieg_api"]["SECRET_KEY"]).strip()
        api_key = str(st.secrets["sieg_api"]["SIEG_API_KEY"]).strip()
        
        # 1. Gerar Token JWT
        res_jwt = requests.post("https://api.sieg.com/api/v1/create-jwt", headers={"X-Client-Id": client_id, "X-Secret-Key": secret_key}, json={})
        if res_jwt.status_code != 200: return pd.DataFrame()
        token_jwt = res_jwt.json().get("Token", "").replace('"', '').strip()

        # 2. Busca
        url_busca = "https://api.sieg.com/api/v1/baixar-xmls"
        headers_busca = {"Authorization": f"Bearer {token_jwt}", "X-API-Key": api_key, "X-Client-Id": client_id, "X-Secret-Key": secret_key}
        
        notas_processadas = []
        tipos_de_nota = [1, 3] # 1: NFe (Insumos), 3: NFSe (Serviços)
        papeis = [{"CnpjDest": CNPJ_JL}, {"CnpjEmit": CNPJ_JL}] # Busca o que comprou e o que vendeu

        for tipo in tipos_de_nota:
            for papel in papeis:
                payload = {
                    "TipoXml": tipo, "Take": 50, "Skip": 0,
                    "DataEmissaoInicio": data_inicio.strftime('%Y-%m-%dT00:00:00.000Z') if tipo == 1 else data_inicio.strftime('%Y-%m-%d'),
                    "DataEmissaoFim": data_fim.strftime('%Y-%m-%dT23:59:59.999Z') if tipo == 1 else data_fim.strftime('%Y-%m-%d'),
                    "BaixarEventos": False
                }
                payload.update(papel)
                
                resposta = requests.post(url_busca, json=payload, headers=headers_busca)
                if 'application/zip' in resposta.headers.get('Content-Type', '').lower() or resposta.content.startswith(b'PK'):
                    with zipfile.ZipFile(io.BytesIO(resposta.content)) as z:
                        for nome_arquivo in [x for x in z.namelist() if x.lower().endswith('.xml')]:
                            with z.open(nome_arquivo) as f:
                                dados = extrair_dados_xml_rapido(f.read().decode('utf-8', errors='ignore'))
                                if dados.get("Chave_Acesso"): notas_processadas.append(dados)

        return pd.DataFrame(notas_processadas)
    except Exception as e:
        st.error(f"Erro no motor SIEG: {e}")
        return pd.DataFrame()

def sincronizar_inter_nuvem(str_data_ini, str_data_fim):
    """Lê a API do Inter num intervalo de datas específico e aplica regras contábeis."""
    caminho_pfx_temp = None
    try:
        creds_inter = st.secrets["inter_api"]
        client_id = creds_inter["client_id"]
        client_secret = creds_inter["client_secret"]
        senha_pfx = creds_inter["pfx_senha"]
        conta = creds_inter["conta_corrente"]
        pfx_b64 = creds_inter["pfx_base64"]
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pfx') as tmp_pfx:
            tmp_pfx.write(base64.b64decode(pfx_b64))
            caminho_pfx_temp = tmp_pfx.name
            
        from inter_sdk_python.InterSdk import InterSdk 
        motor = InterSdk(
            environment="PRODUCTION", 
            client_id=client_id, 
            client_secret=client_secret,
            certificate=caminho_pfx_temp, 
            certificate_password=senha_pfx
        )
        motor.set_account(conta.replace("-",""))
        
        data_ini_buscada = datetime.datetime.strptime(str_data_ini, '%Y-%m-%d')
        data_fim_buscada = datetime.datetime.strptime(str_data_fim, '%Y-%m-%d')
        
        lista_transacoes_total = []
        data_fim_atual = data_fim_buscada
        
        while data_fim_atual >= data_ini_buscada:
            data_ini_atual = max(data_fim_atual - datetime.timedelta(days=29), data_ini_buscada)
            try:
                extrato_obj = motor.banking().retrieve_statement(data_ini_atual.strftime('%Y-%m-%d'), data_fim_atual.strftime('%Y-%m-%d'))
                if hasattr(extrato_obj, 'transactions') and extrato_obj.transactions:
                    lista_transacoes_total.extend(extrato_obj.transactions)
                elif hasattr(extrato_obj, 'transacoes') and extrato_obj.transacoes:
                    lista_transacoes_total.extend(extrato_obj.transacoes)
            except: pass 
            data_fim_atual = data_ini_atual - datetime.timedelta(days=1)
            
        if not lista_transacoes_total:
            if caminho_pfx_temp and os.path.exists(caminho_pfx_temp): os.remove(caminho_pfx_temp)
            return True, "Nenhuma transação encontrada no período selecionado.", []
            
        linhas_injetar = []
        client_gspread = obter_cliente_sheets()
        planilha = client_gspread.open_by_key(ID_PLANILHA_MASTER)
        
        # APONTAMENTO PARA A ABA CORRETA DA IMAGEM
        try:
            aba_caixa = planilha.worksheet("Fluxo_Caixa") 
        except:
            return False, "Aba 'Fluxo_Caixa' não encontrada no Sheets.", []
            
        dados_existentes = aba_caixa.col_values(1) 
        
        regras_list = []
        try:
            aba_regras = planilha.worksheet("Regras_automaticas")
            regras_list = aba_regras.get_all_records()
        except: pass
        
        for idx, transacao in enumerate(lista_transacoes_total):
            def get_val(chaves, padrao=""):
                if isinstance(transacao, dict):
                    for c in chaves:
                        if c in transacao: return transacao[c]
                else:
                    for c in chaves:
                        if hasattr(transacao, c): return getattr(transacao, c)
                return padrao

            titulo = get_val(['title', 'titulo'], '')
            desc_original = get_val(['description', 'descricao'], '')
            descricao = f"{titulo} {desc_original}".strip().upper()
            
            if "SALDO" in descricao: continue 
            
            valor = float(get_val(['value', 'valor', 'valorTransacao', 'valor_transacao'], 0))
            if valor == 0: continue
            
            tipo_op = str(get_val(['operationType', 'tipoOperacao', 'operation_type', 'tipo_operacao'], 'C')).upper()
            tipo = "DESPESA" if tipo_op in ['D', 'DEBIT', 'DEBITO', 'PAYMENT'] else "RECEITA"
                
            import re
            data_lancamento = str_data_fim
            id_transacao = f"INTER-{idx}"
            atributos = transacao.keys() if isinstance(transacao, dict) else dir(transacao)
            
            for attr in atributos:
                if attr.startswith('_'): continue
                try:
                    valor_attr = str(transacao[attr]) if isinstance(transacao, dict) else str(getattr(transacao, attr))
                    if 'dat' in attr.lower():
                        match_dt = re.search(r'(202\d-[0-1]\d-[0-3]\d)', valor_attr)
                        if match_dt: data_lancamento = match_dt.group(1)
                        else:
                            match_dt_br = re.search(r'([0-3]\d/[0-1]\d/202\d)', valor_attr)
                            if match_dt_br:
                                dt_br = match_dt_br.group(1)
                                data_lancamento = f"{dt_br[6:10]}-{dt_br[3:5]}-{dt_br[0:2]}"
                    if attr.lower() in ['id', 'idtransacao', 'id_transacao', 'transactionid', 'transaction_id', 'identificador']:
                        if valor_attr and len(valor_attr) > 5: id_transacao = valor_attr
                except: continue
                    
            if id_transacao == f"INTER-{idx}": id_transacao = f"INTER-{data_lancamento}-{idx}"
                
            try: data_formatada = datetime.datetime.strptime(data_lancamento, '%Y-%m-%d').strftime('%Y-%m-%d')
            except: data_formatada = data_lancamento
                
            if id_transacao in dados_existentes: continue 
            
            conta_contrapartida = ""
            categoria_dash = "⚠️ A Classificar"
            
            # LEITURA EXATA DAS COLUNAS DA SUA IMAGEM DAS REGRAS
            for regra in regras_list:
                termo = str(regra.get('Palavra_Chave no Extrato', '')).upper()
                if termo and termo in descricao:
                    conta_contrapartida = str(regra.get('Conta_Contrapartida (Dominio)', '')).strip()
                    cat_temp = str(regra.get('Categoria_Gerencial (Dashboard)', '')).strip()
                    if cat_temp: categoria_dash = cat_temp
                    break
                    
            if conta_contrapartida == "":
                if tipo == "RECEITA":
                    conta_contrapartida = "504"
                    if categoria_dash == "⚠️ A Classificar": categoria_dash = "Receitas de serviços"
                else: 
                    conta_contrapartida = "506"
                    if categoria_dash == "⚠️ A Classificar": categoria_dash = "Fornecedores"
            
            c_deb = "747" if tipo == "RECEITA" else conta_contrapartida
            c_cred = conta_contrapartida if tipo == "RECEITA" else "747"
            
            # INJEÇÃO DAS 9 COLUNAS EXATAS DA ABA FLUXO_CAIXA
            nova_linha = [
                id_transacao,            # A: ID_Transacao
                data_formatada,          # B: Data
                "Banco Inter",           # C: Conta/Banco
                tipo,                    # D: Tipo
                f"{abs(valor):.2f}".replace('.', ','), # E: Valor
                descricao,               # F: Descricao_Original
                categoria_dash,          # G: Categoria_Gerencial
                c_deb,                   # H: Conta_Debito
                c_cred                   # I: Conta_Credito
            ]
            linhas_injetar.append(nova_linha)
                
        if caminho_pfx_temp and os.path.exists(caminho_pfx_temp): os.remove(caminho_pfx_temp)

        if linhas_injetar:
            aba_caixa.append_rows(linhas_injetar)
            return True, f"Sucesso! {len(linhas_injetar)} movimentações salvas na aba Fluxo_Caixa.", lista_transacoes_total
        else:
            return True, "Sincronizado. Nenhuma movimentação inédita no período.", []
            
    except Exception as e:
        if caminho_pfx_temp and os.path.exists(caminho_pfx_temp):
            try: os.remove(caminho_pfx_temp)
            except: pass
        return False, f"Falha na conexão com o Banco Inter: {str(e)}", []

# ==========================================
# 3. INTERFACE VISUAL E NAVEGAÇÃO
# ==========================================
if "autenticado" not in st.session_state:
    st.title("🔒 Acesso Restrito - J&L Incorporadora")
    senha = st.text_input("Digite a senha:", type="password")
    if senha == st.secrets.get("senha_sistema", "Dldp2023"):
        st.session_state["autenticado"] = True
        st.rerun()
else:
    st.sidebar.title("🏢 Módulos J&L")
    modulo = st.sidebar.radio("Navegação:", [
        "📊 Dashboard Executivo",
        "🏦 Tesouraria (Banco Inter)",
        "📥 Auditoria de Obras (SIEG)",
        "📝 Exportação Contábil (Domínio)"
    ])
    
    st.sidebar.divider()
    if st.sidebar.button("Sair"):
        del st.session_state["autenticado"]
        st.rerun()

    # --- TELA 1: TESOURARIA (INTER) ---
    if modulo == "🏦 Tesouraria (Banco Inter)":
        st.title("🏦 Sincronização Open Finance (Inter)")
        col1, col2 = st.columns(2)
        with col1: data_ini = st.date_input("Data Inicial", datetime.date.today() - datetime.timedelta(days=7))
        with col2: data_fim = st.date_input("Data Final", datetime.date.today())
        
        if st.button("🚀 Capturar Extrato", type="primary"):
            with st.spinner("Conectando ao Banco Inter..."):
                sucesso, msg, transacoes = sincronizar_inter_nuvem(data_ini.strftime('%Y-%m-%d'), data_fim.strftime('%Y-%m-%d'))
                if sucesso:
                    st.success(msg)
                    if transacoes:
                        # Extrai os dicionários de forma segura ignorando metadados da API
                        df_transacoes = pd.DataFrame([t if isinstance(t, dict) else {a: getattr(t, a) for a in dir(t) if not a.startswith('_')} for t in transacoes])
                        st.dataframe(df_transacoes)
                else: st.error(msg)

    # --- TELA 2: AUDITORIA (SIEG) ---
    elif modulo == "📥 Auditoria de Obras (SIEG)":
        st.title("📥 Captura de Notas de Insumos e Serviços")
        col1, col2 = st.columns(2)
        with col1: data_ini = st.date_input("Data Inicial", datetime.date.today() - datetime.timedelta(days=15), key="sieg_ini")
        with col2: data_fim = st.date_input("Data Final", datetime.date.today(), key="sieg_fim")
        
        if st.button("🔍 Rastrear Notas Fiscais", type="primary"):
            with st.spinner("Varrendo portal da Receita Federal..."):
                df_notas = consultar_sieg_api(data_ini, data_fim)
                if not df_notas.empty:
                    st.success(f"{len(df_notas)} notas encontradas.")
                    st.dataframe(df_notas, use_container_width=True)
                else:
                    st.warning("Nenhuma nota encontrada.")

    # --- TELA 3: DOMÍNIO ---
    elif modulo == "📝 Exportação Contábil (Domínio)":
        st.title("📝 Fechamento Contábil e Exportação (Domínio)")
        st.markdown("Edite as contas contábeis e salve as correções direto na planilha.")
        
        col_d1, col_d2, col_d3 = st.columns([1.5, 1.5, 2])
        with col_d1: data_ini_txt = st.date_input("Data Inicial", datetime.date.today().replace(day=1), key="dt_ini_contabil")
        with col_d2: data_fim_txt = st.date_input("Data Final", datetime.date.today(), key="dt_fim_contabil")
        with col_d3: 
            banco_filtro = st.selectbox(
                "Filtrar por Banco:", 
                ["Todos os Bancos", "Banco Inter", "SICOOB"],
                key="sel_banco_contabil"
            )
        
        try:
            client_gspread = obter_cliente_sheets()
            planilha = client_gspread.open_by_key(ID_PLANILHA_MASTER)
            try: aba_caixa = planilha.worksheet("Fluxo_Caixa")
            except: 
                st.error("Aba 'Fluxo_Caixa' não encontrada.")
                st.stop()
                
            dados_rows = aba_caixa.get_all_values()
            
            if len(dados_rows) > 1:
                df_caixa = pd.DataFrame(dados_rows[1:], columns=dados_rows[0])
                df_caixa.columns = df_caixa.columns.str.strip()
                
                #MAPEAMENTO DAS COLUNAS EXATAS DA ABA FLUXO_CAIXA
                colunas_necessarias = ['Data', 'Conta/Banco', 'Tipo', 'Valor', 'Descricao_Original', 'Categoria_Gerencial', 'Conta_Debito', 'Conta_Credito']
                for col in colunas_necessarias:
                    if col not in df_caixa.columns:
                        df_caixa[col] = ""
                
                df_caixa['Data_Filtro'] = pd.to_datetime(df_caixa['Data'], errors='coerce').dt.date
                mask_datas = (df_caixa['Data_Filtro'] >= data_ini_txt) & (df_caixa['Data_Filtro'] <= data_fim_txt)
                
                if banco_filtro != "Todos os Bancos":
                    mask_bancos = df_caixa['Conta/Banco'].str.strip().str.upper() == banco_filtro.upper()
                else:
                    mask_bancos = pd.Series(True, index=df_caixa.index)
                
                df_filtrado = df_caixa[mask_datas & mask_bancos].copy()
                
                st.write(f"📌 Exibindo **{len(df_filtrado)}** lançamentos da conta **{banco_filtro}** no período selecionado.")
                
                if not df_filtrado.empty:
                    df_view = df_filtrado[colunas_necessarias].copy()
                    
                    df_editado = st.data_editor(
                        df_view,
                        column_config={
                            "Data": st.column_config.TextColumn("Data", disabled=True),
                            "Conta/Banco": st.column_config.TextColumn("Banco", disabled=True),
                            "Tipo": st.column_config.TextColumn("Tipo", disabled=True),
                            "Valor": st.column_config.TextColumn("Valor", disabled=True),
                            "Descricao_Original": st.column_config.TextColumn("Histórico Extrato", disabled=True),
                            "Categoria_Gerencial": st.column_config.TextColumn("Categoria Dash", disabled=True),
                            "Conta_Debito": st.column_config.TextColumn("Débito (Duplo Clique)"),
                            "Conta_Credito": st.column_config.TextColumn("Crédito (Duplo Clique)"),
                        },
                        use_container_width=True,
                        num_rows="fixed",
                        key="editor_contabil"
                    )
                    
                    st.divider()
                    col_info, col_save = st.columns([2, 2])
                    
                    with col_info:
                        st.info("💡 Dê 2 cliques na célula para editar a conta. Depois, clique em Salvar para gravar na base de dados.")
                        
                    with col_save:
                        if st.button("💾 Salvar Correções na Planilha Central", type="secondary", use_container_width=True):
                            with st.spinner("Gravando alterações..."):
                                try:
                                    df_caixa.loc[df_editado.index, 'Conta_Debito'] = df_editado['Conta_Debito']
                                    df_caixa.loc[df_editado.index, 'Conta_Credito'] = df_editado['Conta_Credito']
                                    
                                    df_caixa = df_caixa.drop(columns=['Data_Filtro'])
                                    df_caixa = df_caixa.fillna("").astype(str)
                                    
                                    dados_salvar = [df_caixa.columns.values.tolist()] + df_caixa.values.tolist()
                                    
                                    aba_caixa.clear()
                                    aba_caixa.append_rows(dados_salvar)
                                    
                                    st.success("Alterações salvas com sucesso!")
                                    time.sleep(1.2)
                                    st.rerun()
                                except Exception as err_sheets:
                                    st.error(f"Erro ao salvar dados: {err_sheets}")
                else:
                    st.warning("Nenhum lançamento encontrado neste intervalo.")
            else:
                st.caption("A base central de caixas e bancos está vazia.")
                
        except Exception as e:
            st.error(f"Erro ao processar tela contábil: {e}")
    # --- TELA 4: DASHBOARD ---
    elif modulo == "📊 Dashboard Executivo":
        st.title("📊 Visão Consolidada das Obras")
        st.write("Acompanhamento de fluxo financeiro.")
