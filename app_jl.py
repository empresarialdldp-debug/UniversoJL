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
ID_PLANILHA_MASTER = "1azNi85ir10HXcJGd1pa7Wh2ceZ5hQb_U9qMPqldjFjs"
ID_PLANILHA_WM_anual ="1O499mTddJPCTpAC5oPEATmYgd1BBKcWL7oVU_vFumBU"
ID_PLANILHA_Recebimento_Wilson_Moreira_2026 ="1N3f4sjn-xCDl2wDORhsHVEKWDLxRQGp7A8bgh1h3J8k"
ID_PLANILHA_Recebimento_Wilson_Moreira_2025 ="1xVJNLe6oLfGr-7sjsz8Yte8zSIHxfLTo811de_tACeo"
ID_PLANILHA_Recebimento_Wilson_Moreira_2024 ="1UKclnSZfP1MQQHMDHxRb4beHMBJqHAZIoduaeu7DYDE"
ID_PLANILHA_Recebimento_Wilson_Moreira_2023 ="1V_zqHga4gkH6geLi06MOtrJDC8OPc6QVBjFhdFOusC0"
ID_PLANILHA_Recebimento_Wilson_Moreira_2022 ="1P6HPr4vdiH0V9buhc2iBYLNtiv4b1OoaUO9nvOs73OM"

# ==========================================
# EXTRAÇÃO AUTOMÁTICA DO SDK DO BANCO INTER
# ==========================================
if not os.path.exists('inter_sdk_python') and os.path.exists('inter_sdk_python.zip'):
    with zipfile.ZipFile('inter_sdk_python.zip', 'r') as zip_ref:
        zip_ref.extractall('.')
    st.toast("📦 SDK do Banco Inter extraído com sucesso na nuvem!", icon="⚙️")

# ==========================================
# IMPORTAR O SDK EXTRAÍDO
# ==========================================
from inter_sdk_python.billing.models.BillingIssueRequest import BillingIssueRequest
from inter_sdk_python.billing.models.Person import Person 
try:
    from inter_sdk_python.commons.models.PersonType import PersonType
except ImportError:
    from enum import Enum
    class PersonType(Enum):
        FISICA = "FISICA"
        JURIDICA = "JURIDICA"

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

def salvar_extrato_padronizado(transacoes_padrao, nome_banco):
    """
    Recebe uma lista padrão [{'id', 'data', 'valor', 'tipo', 'descricao'}]
    Aplica as regras automáticas + a Regra do 747 e salva na planilha Fluxo_Caixa.
    """
    if not transacoes_padrao:
        return True, "Nenhuma transação para processar.", []

    client_gspread = obter_cliente_sheets()
    if client_gspread is None:
        return False, "Falha ao conectar no Google Sheets.", []
        
    planilha = client_gspread.open_by_key(ID_PLANILHA_MASTER)
    
    try:
        aba_caixa = planilha.worksheet("Fluxo_Caixa") 
    except Exception as e:
        return False, f"Aba 'Fluxo_Caixa' não encontrada.", []
        
    dados_existentes = aba_caixa.col_values(1) 
    
    regras_list = []
    try:
        aba_regras = planilha.worksheet("Regras_automaticas")
        regras_list = aba_regras.get_all_records()
    except: pass
    
    linhas_injetar = []
    for t in transacoes_padrao:
        if "SALDO" in t['descricao']: continue
        if t['id'] in dados_existentes: continue 
        
        conta_contrapartida = ""
        categoria_dash = "⚠️ A Classificar"
        
        for regra in regras_list:
            termo = str(regra.get('Palavra_Chave no Extrato', '')).upper()
            if termo and termo in t['descricao']:
                conta_contrapartida = str(regra.get('Conta_Contrapartida (Dominio)', '')).strip()
                cat_temp = str(regra.get('Categoria_Gerencial (Dashboard)', '')).strip()
                if cat_temp: categoria_dash = cat_temp
                break
        
        # REGRA CONTÁBIL J&L (747 Fixa)
        if t['tipo'] == "RECEITA":
            c_deb = "747"
            c_cred = conta_contrapartida
        else: # DESPESA
            c_deb = conta_contrapartida
            c_cred = "747"
        
        nova_linha = [
            t['id'], t['data'], nome_banco, t['tipo'], 
            f"{t['valor']:.2f}".replace('.', ','), 
            t['descricao'], categoria_dash, c_deb, c_cred                   
        ]
        linhas_injetar.append(nova_linha)

    if linhas_injetar:
        aba_caixa.append_rows(linhas_injetar)
        return True, f"Sucesso! {len(linhas_injetar)} movimentações do {nome_banco} salvas.", linhas_injetar
    else:
        return True, "Sincronizado. Nenhuma movimentação inédita no período.", []

def processar_ofx_bb(conteudo_ofx):
    """Lê o arquivo OFX do Banco do Brasil e converte para o formato padrão do sistema."""
    transacoes = []
    blocos = re.findall(r'<STMTTRN>(.*?)</STMTTRN>', conteudo_ofx, re.DOTALL)
    
    for b in blocos:
        match_dt = re.search(r'<DTPOSTED>(\d{8})', b)
        if not match_dt: continue
        dt_raw = match_dt.group(1)
        data_fmt = f"{dt_raw[:4]}-{dt_raw[4:6]}-{dt_raw[6:8]}"
        
        match_val = re.search(r'<TRNAMT>([\-\d\.]+)', b)
        val = float(match_val.group(1)) if match_val else 0.0
        if val == 0: continue
        tipo = "RECEITA" if val > 0 else "DESPESA"
        
        match_memo = re.search(r'<MEMO>(.*?)(?:\r|\n|<)', b)
        memo = match_memo.group(1).strip().upper() if match_memo else "TRANSACAO SEM NOME"
        
        match_id = re.search(r'<FITID>(.*?)(?:\r|\n|<)', b)
        tid = match_id.group(1).strip() if match_id else f"BB-{dt_raw}-{abs(val)}"
        
        transacoes.append({
            "id": tid, "data": data_fmt, "valor": abs(val), "tipo": tipo, "descricao": memo
        })
    return transacoes

def buscar_inter_api(str_data_ini, str_data_fim):
    """Conecta no Banco Inter e extrai os dados crus no formato padronizado."""
    caminho_pfx_temp = None
    try:
        creds_inter = st.secrets["inter_api"]
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pfx') as tmp_pfx:
            tmp_pfx.write(base64.b64decode(creds_inter["pfx_base64"]))
            caminho_pfx_temp = tmp_pfx.name
            
        from inter_sdk_python.InterSdk import InterSdk 
        motor = InterSdk(
            environment="PRODUCTION", client_id=creds_inter["client_id"], 
            client_secret=creds_inter["client_secret"],
            certificate=caminho_pfx_temp, certificate_password=creds_inter["pfx_senha"]
        )
        motor.set_account(creds_inter["conta_corrente"].replace("-",""))
        
        dt_ini = datetime.datetime.strptime(str_data_ini, '%Y-%m-%d')
        dt_fim = datetime.datetime.strptime(str_data_fim, '%Y-%m-%d')
        
        lista_raw = []
        dt_fim_atual = dt_fim
        while dt_fim_atual >= dt_ini:
            dt_ini_atual = max(dt_fim_atual - datetime.timedelta(days=29), dt_ini)
            try:
                ext = motor.banking().retrieve_statement(dt_ini_atual.strftime('%Y-%m-%d'), dt_fim_atual.strftime('%Y-%m-%d'))
                if hasattr(ext, 'transactions') and ext.transactions: lista_raw.extend(ext.transactions)
                elif hasattr(ext, 'transacoes') and ext.transacoes: lista_raw.extend(ext.transacoes)
            except: pass 
            dt_fim_atual = dt_ini_atual - datetime.timedelta(days=1)
            
        if caminho_pfx_temp and os.path.exists(caminho_pfx_temp): os.remove(caminho_pfx_temp)
        
        # Padronizando dados do Inter
        transacoes_padrao = []
        for t in lista_raw:
            def get_val(chaves, padrao=""):
                if isinstance(t, dict):
                    for c in chaves:
                        if c in t: return t[c]
                else:
                    for c in chaves:
                        if hasattr(t, c): return getattr(t, c)
                return padrao
                
            val = float(get_val(['value', 'valor', 'valorTransacao'], 0))
            if val == 0: continue
            
            tipo_op = str(get_val(['operationType', 'tipoOperacao', 'operation_type'], 'C')).upper()
            tipo = "DESPESA" if tipo_op in ['D', 'DEBIT', 'DEBITO', 'PAYMENT'] else "RECEITA"
            
            titulo = get_val(['title', 'titulo'], '')
            desc = get_val(['description', 'descricao'], '')
            memo = f"{titulo} {desc}".strip().upper()
            
            dt_lanc = str_data_fim
            tid = f"INTER-{len(transacoes_padrao)}"
            
            atributos = t.keys() if isinstance(t, dict) else dir(t)
            for attr in atributos:
                if attr.startswith('_'): continue
                try:
                    v_attr = str(t[attr]) if isinstance(t, dict) else str(getattr(t, attr))
                    if 'dat' in attr.lower():
                        match_dt = re.search(r'(202\d-[0-1]\d-[0-3]\d)', v_attr)
                        if match_dt: dt_lanc = match_dt.group(1)
                    if attr.lower() in ['id', 'idtransacao', 'transactionid', 'identificador']:
                        if v_attr and len(v_attr) > 5: tid = v_attr
                except: continue
                
            transacoes_padrao.append({
                "id": tid, "data": dt_lanc, "valor": abs(val), "tipo": tipo, "descricao": memo
            })
            
        return transacoes_padrao
    except Exception as e:
        if caminho_pfx_temp and os.path.exists(caminho_pfx_temp): os.remove(caminho_pfx_temp)
        raise e

def emitir_boletos_lote(df_boletos):
    """Gera boletos no Banco Inter a partir de um DataFrame e retorna os PDFs."""
    caminho_pfx_temp = None
    resultados = []
    
    try:
        creds_inter = st.secrets["inter_api"]
        client_id = creds_inter["client_id"]
        client_secret = creds_inter["client_secret"]
        senha_pfx = creds_inter["pfx_senha"]
        conta = creds_inter["conta_corrente"]
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pfx') as tmp_pfx:
            tmp_pfx.write(base64.b64decode(creds_inter["pfx_base64"]))
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
        
        from decimal import Decimal
        
        for index, row in df_boletos.iterrows():
            nome_cliente = str(row['Nome_Cliente']).strip()[:100]
            cpf_limpo = re.sub(r'\D', '', str(row['CPF_CNPJ']))
            controle = re.sub(r'[^a-zA-Z0-9]', '', str(row['seuNumero']))[:14] + "S"
            
            try:
                data_vencimento = pd.to_datetime(row['Vencimento']).strftime('%Y-%m-%d')
            except:
                resultados.append({"Cliente": nome_cliente, "Status": "Erro", "Motivo": "Data Inválida", "PDF": None})
                continue

            try:
                # Monta Pagador
                pagador = Person()
                pagador.nome = pagador.name = nome_cliente
                pagador.cpf_cnpj = pagador.cpfCnpj = cpf_limpo
                pagador.cep = pagador.zip_code = pagador.zipCode = re.sub(r'\D', '', str(row['CEP']))
                pagador.numero = pagador.number = str(row['Número']).strip() if not pd.isna(row['Número']) else "0"
                pagador.endereco = pagador.address = pagador.logradouro = "Logradouro"
                pagador.cidade = pagador.city = "Cidade"
                pagador.uf = pagador.state = "MG"
                pagador.bairro = pagador.neighborhood = "Bairro"
                pagador.tipo_pessoa = pagador.personType = PersonType.FISICA if len(cpf_limpo) <= 11 else PersonType.JURIDICA

                # Monta Boleto
                boleto = BillingIssueRequest()
                boleto.seu_numero = boleto.seuNumero = controle
                boleto.valor_nominal = boleto.valorNominal = Decimal(str(round(row['Valor'], 2)))
                boleto.data_vencimento = boleto.dataVencimento = data_vencimento
                boleto.pagador = boleto.payer = pagador
                boleto.num_dias_agenda = 0

                # Emissão
                res = motor.billing().issue_billing(boleto)
                n_num = getattr(res, 'nossoNumero', None) or getattr(res, 'nosso_numero', None)
                
                if n_num:
                    time.sleep(3) # Pausa para o banco renderizar
                    pdf_path = os.path.join(tempfile.gettempdir(), f"{controle}.pdf")
                    motor.billing().retrieve_billing_pdf(str(n_num), file=pdf_path)
                    
                    with open(pdf_path, "rb") as f:
                        pdf_bytes = f.read()
                        
                    resultados.append({"Cliente": nome_cliente, "Nosso Número": n_num, "Status": "Sucesso", "PDF": pdf_bytes})
                    
            except Exception as e:
                erro_msg = str(e)
                if hasattr(e, 'error') and e.error: erro_msg = e.error.detail
                resultados.append({"Cliente": nome_cliente, "Status": "Rejeitado", "Motivo": erro_msg, "PDF": None})
                
        if caminho_pfx_temp and os.path.exists(caminho_pfx_temp): os.remove(caminho_pfx_temp)
        return resultados
        
    except Exception as e:
        if caminho_pfx_temp and os.path.exists(caminho_pfx_temp): os.remove(caminho_pfx_temp)
        st.error(f"Erro fatal de conexão: {e}")
        return None
# ==========================================
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
        "🏦 Tesouraria Central (Inter e BB)",
        "📝 Exportação Contábil (Domínio)",
        "📥 Auditoria de Obras (SIEG)",
        "📊 Dashboard Executivo",
        "💸 Faturamento e Boletos",
        "⚙️ Engenharia (Transbordo)"
    ])
    
    st.sidebar.divider()
    if st.sidebar.button("Sair"):
        del st.session_state["autenticado"]
        st.rerun()

    # --- TELA 1: TESOURARIA (UNIFICADA) ---
    if modulo == "🏦 Tesouraria Central (Inter e BB)":
        st.title("🏦 Central de Tesouraria e Caixas")
        st.markdown("Sincronize ou importe os extratos bancários. O sistema aplicará a **Regra de Ouro (Conta 747)** automaticamente.")
        
        col_inter, col_bb = st.columns(2)
        
        # --- PAINEL BANCO INTER ---
        with col_inter:
            st.subheader("🟠 API Banco Inter")
            st.markdown("Conexão direta e automática (Open Finance)")
            data_ini = st.date_input("Data Inicial", datetime.date.today() - datetime.timedelta(days=7), key="dt_inter_ini")
            data_fim = st.date_input("Data Final", datetime.date.today(), key="dt_inter_fim")
            
            if st.button("🚀 Puxar Extrato Inter", type="primary", use_container_width=True):
                with st.spinner("Puxando API do Inter..."):
                    try:
                        t_padrao = buscar_inter_api(data_ini.strftime('%Y-%m-%d'), data_fim.strftime('%Y-%m-%d'))
                        sucesso, msg, t_inseridas = salvar_extrato_padronizado(t_padrao, "Banco Inter")
                        if sucesso:
                            st.success(msg)
                            if t_inseridas:
                                df_disp = pd.DataFrame(t_inseridas, columns=['ID', 'Data', 'Banco', 'Tipo', 'Valor', 'Descrição Original', 'Categoria', 'Conta Débito', 'Conta Crédito'])
                                st.dataframe(df_disp, hide_index=True)
                        else: st.error(msg)
                    except Exception as e:
                        st.error(f"Erro no Inter: {e}")

        # --- PAINEL BANCO DO BRASIL ---
        with col_bb:
            st.subheader("🔵 Arquivo OFX - Banco do Brasil")
            st.markdown("Importação manual via arquivo do banco")
            arquivo_ofx = st.file_uploader("Arraste o extrato .OFX aqui", type=["ofx"])
            
            if st.button("📂 Processar OFX", type="primary", use_container_width=True) and arquivo_ofx:
                with st.spinner("Lendo arquivo e aplicando regras..."):
                    try:
                        conteudo = arquivo_ofx.read().decode('utf-8', errors='ignore')
                        t_padrao = processar_ofx_bb(conteudo)
                        sucesso, msg, t_inseridas = salvar_extrato_padronizado(t_padrao, "Banco do Brasil")
                        if sucesso:
                            st.success(msg)
                            if t_inseridas:
                                df_disp = pd.DataFrame(t_inseridas, columns=['ID', 'Data', 'Banco', 'Tipo', 'Valor', 'Descrição Original', 'Categoria', 'Conta Débito', 'Conta Crédito'])
                                st.dataframe(df_disp, hide_index=True)
                        else: st.error(msg)
                    except Exception as e:
                        st.error(f"Erro no processamento OFX: {e}")

    # --- TELA 2: DOMÍNIO (CONTABILIDADE) ---
    elif modulo == "📝 Exportação Contábil (Domínio)":
        st.title("📝 Fechamento Contábil e Exportação (Domínio)")
        st.markdown("Preencha as contas em branco. O sistema salvará no Caixa e **criará regras automáticas** para o futuro!")
        
        col_d1, col_d2, col_d3 = st.columns([1.5, 1.5, 2])
        with col_d1: data_ini_txt = st.date_input("Data Inicial", datetime.date.today().replace(day=1), key="dt_ini_contabil")
        with col_d2: data_fim_txt = st.date_input("Data Final", datetime.date.today(), key="dt_fim_contabil")
        with col_d3: 
            banco_filtro = st.selectbox("Filtrar por Banco:", ["Todos os Bancos", "Banco Inter", "Banco do Brasil"], key="sel_banco_contabil")
        
        try:
            client_gspread = obter_cliente_sheets()
            if client_gspread is None:
                st.error("Falha de conexão com o Google Sheets.")
                st.stop()
                
            planilha = client_gspread.open_by_key(ID_PLANILHA_MASTER)
            aba_caixa = planilha.worksheet("Fluxo_Caixa")
            dados_rows = aba_caixa.get_all_values()
            
            if len(dados_rows) > 1:
                df_caixa = pd.DataFrame(dados_rows[1:], columns=dados_rows[0])
                df_caixa.columns = df_caixa.columns.str.strip()
                
                colunas_necessarias = ['Data', 'Conta/Banco', 'Tipo', 'Valor', 'Descricao_Original', 'Categoria_Gerencial', 'Conta_Debito', 'Conta_Credito']
                for col in colunas_necessarias:
                    if col not in df_caixa.columns: df_caixa[col] = ""
                
                df_caixa['Data_Filtro'] = pd.to_datetime(df_caixa['Data'], errors='coerce').dt.date
                mask_datas = (df_caixa['Data_Filtro'] >= data_ini_txt) & (df_caixa['Data_Filtro'] <= data_fim_txt)
                mask_bancos = pd.Series(True, index=df_caixa.index) if banco_filtro == "Todos os Bancos" else df_caixa['Conta/Banco'].str.strip().str.upper() == banco_filtro.upper()
                
                df_filtrado = df_caixa[mask_datas & mask_bancos].copy()
                
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
                            "Categoria_Gerencial": st.column_config.TextColumn("Categoria Dash"),
                            "Conta_Debito": st.column_config.TextColumn("Débito (Duplo Clique)"),
                            "Conta_Credito": st.column_config.TextColumn("Crédito (Duplo Clique)"),
                        },
                        use_container_width=True, num_rows="fixed", key="editor_contabil"
                    )
                    
                    st.divider()
                    col_info, col_save = st.columns([2, 2])
                    with col_info:
                        st.info("💡 As contas preenchidas viram regras eternas. Você não precisará classificá-las de novo no próximo mês.")
                        
                    with col_save:
                        if st.button("💾 Salvar Planilha & Criar Regras Automáticas", type="secondary", use_container_width=True):
                            with st.spinner("Gravando alterações e treinando o sistema..."):
                                try:
                                    novas_regras = []
                                    for idx, row in df_editado.iterrows():
                                        deb_novo = str(row['Conta_Debito']).strip()
                                        cred_novo = str(row['Conta_Credito']).strip()
                                        deb_velho = str(df_view.loc[idx, 'Conta_Debito']).strip()
                                        cred_velho = str(df_view.loc[idx, 'Conta_Credito']).strip()
                                        
                                        contrapartida_aprendida = ""
                                        if row['Tipo'] == "DESPESA" and deb_novo != deb_velho and deb_novo != "747" and deb_novo != "":
                                            contrapartida_aprendida = deb_novo
                                        elif row['Tipo'] == "RECEITA" and cred_novo != cred_velho and cred_novo != "747" and cred_novo != "":
                                            contrapartida_aprendida = cred_novo
                                            
                                        if contrapartida_aprendida:
                                            novas_regras.append([row['Descricao_Original'], row['Categoria_Gerencial'], contrapartida_aprendida, "Lançamento Automático"])

                                    df_caixa.loc[df_editado.index, 'Conta_Debito'] = df_editado['Conta_Debito']
                                    df_caixa.loc[df_editado.index, 'Conta_Credito'] = df_editado['Conta_Credito']
                                    df_caixa.loc[df_editado.index, 'Categoria_Gerencial'] = df_editado['Categoria_Gerencial']
                                    
                                    df_caixa = df_caixa.drop(columns=['Data_Filtro'])
                                    df_caixa = df_caixa.fillna("").astype(str)
                                    dados_salvar = [df_caixa.columns.values.tolist()] + df_caixa.values.tolist()
                                    aba_caixa.clear()
                                    aba_caixa.append_rows(dados_salvar)
                                    
                                    if novas_regras:
                                        aba_regras = planilha.worksheet("Regras_automaticas")
                                        aba_regras.append_rows(novas_regras)
                                        st.toast(f"🤖 {len(novas_regras)} novas regras aprendidas com sucesso!", icon="🧠")
                                    
                                    st.success("Tudo salvo com sucesso!")
                                    time.sleep(1.5)
                                    st.rerun()
                                except Exception as err_sheets:
                                    st.error(f"Erro ao salvar dados: {err_sheets}")
                else:
                    st.warning("Nenhum lançamento encontrado neste intervalo.")
            else:
                st.caption("A base central está vazia.")
        except Exception as e:
            st.error(f"Erro ao processar tela contábil: {e}")
            
   # --- TELA 4: DASHBOARD EXECUTIVO ---
    elif modulo == "📊 Dashboard Executivo":
        st.title("📊 Painel de Recebíveis - Wilson Moreira")
        st.markdown("Acompanhe os recebimentos e lance **Documentação ou Correções** diretamente no saldo contratado.")

        client_gspread = obter_cliente_sheets()
        if not client_gspread:
            st.error("Sem conexão com o Google Sheets.")
            st.stop()
            
        planilha_master = client_gspread.open_by_key(ID_PLANILHA_MASTER)

        # --- CÉREBRO MATEMÁTICO BLINDADO ---
        def safe_to_float(val):
            if isinstance(val, (int, float)): return float(val)
            val = str(val).replace('R$', '').replace(' ', '').strip()
            if not val or val.lower() in ['nan', 'none', '']: return 0.0
            
            if ',' in val and '.' in val:
                if val.rfind(',') > val.rfind('.'):
                    val = val.replace('.', '').replace(',', '.')
                else:
                    val = val.replace(',', '')
            elif ',' in val:
                val = val.replace(',', '.')
            try: return float(val)
            except: return 0.0

        # ==========================================
        # 1. PAINEL DE AJUSTES (FORMULÁRIO DE LANÇAMENTO)
        # ==========================================
        st.subheader("🛠️ Lançar Documentação ou Correção")
        with st.expander("Clique para adicionar um valor ao saldo de um contrato"):
            with st.form("form_correcao", clear_on_submit=True):
                col1, col2, col3, col4 = st.columns([3, 2, 2, 2])
                
                try:
                    p_2026 = client_gspread.open_by_key(ID_PLANILHA_Recebimento_Wilson_Moreira_2026)
                    abas = p_2026.worksheets()
                    df_lista = pd.DataFrame()
                    for aba in abas:
                        dados_aba = aba.get_all_values()
                        if len(dados_aba) > 3:
                            cab = [str(c).strip().upper() for c in dados_aba[2]]
                            if "NOME DO ADQUIRENTE" in cab:
                                df_lista = pd.DataFrame(dados_aba[3:], columns=cab)
                                break
                    
                    if not df_lista.empty:
                        termos_excluir = ['BASE IR', 'IR', 'IR ADICIONAL', 'CSLL', 'VALOR LIQUIDO', 'VALOR BRUTO', 'PIS', 'COFINS', 'VALOR DA VENDA']
                        df_lista = df_lista[~df_lista['NOME DO ADQUIRENTE'].str.strip().str.upper().isin(termos_excluir)]
                        df_lista = df_lista[~df_lista['NOME DO ADQUIRENTE'].str.strip().str.upper().str.startswith('R$')]
                        df_lista = df_lista[df_lista['NOME DO ADQUIRENTE'].astype(str).str.replace(r'[^a-zA-Z]', '', regex=True).str.len() > 2]
                        
                        lista_formatada = []
                        for _, row in df_lista.iterrows():
                            nome = str(row.get('NOME DO ADQUIRENTE', '')).strip()
                            unidade = str(row.get('DESCRIÇÃO RESUMIDA DA UNIDADE', '')).strip()
                            
                            texto_combo = f"{nome}"
                            if unidade: texto_combo += f" - {unidade}"
                            if nome and texto_combo not in lista_formatada:
                                lista_formatada.append(texto_combo)
                                
                        lista_clientes = sorted(lista_formatada)
                    else:
                        lista_clientes = ["Tabela não encontrada nas abas..."]
                except Exception as e:
                    lista_clientes = [f"Erro na conexão com o Sheets"]

                cliente_combo = col1.selectbox("Selecione o Contrato:", lista_clientes)
                motivo_ajuste = col2.text_input("Motivo (Ex: Doc, INCC):")
                valor_digitado = col3.text_input("Valor (R$):", placeholder="Ex: 1750.10 ou 1750,10")
                
                import datetime
                data_selecionada = col4.date_input("Data do Lançamento:", value=datetime.date.today(), format="DD/MM/YYYY")
                btn_salvar_ajuste = st.form_submit_button("💾 Salvar Correção no Contrato")
                
                valor_ajuste = safe_to_float(valor_digitado)
                
                if btn_salvar_ajuste and cliente_combo and valor_ajuste > 0 and "Erro" not in cliente_combo:
                    try:
                        chave_banco = cliente_combo.strip()
                        try:
                            aba_ajustes = planilha_master.worksheet("Ajustes_Contratos")
                        except:
                            aba_ajustes = planilha_master.add_worksheet(title="Ajustes_Contratos", rows="100", cols="4")
                            aba_ajustes.append_row(["Data_Registro", "Cliente", "Motivo", "Valor_Ajuste"])
                            
                        data_formatada = data_selecionada.strftime('%d/%m/%Y')
                        aba_ajustes.append_row([data_formatada, chave_banco, motivo_ajuste, float(valor_ajuste)])
                        
                        st.success(f"Acréscimo de R$ {valor_ajuste:.2f} salvo com sucesso para {chave_banco}!")
                        import time
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao salvar ajuste: {e}")

        st.divider()

        # ==========================================
        # 2. CONSOLIDAÇÃO MATEMÁTICA E ABAS
        # ==========================================
        with st.spinner("Consolidando Livro Razão, Contratos, Correções e Tributos..."):
            try:
                # A. PAGAMENTOS
                aba_recebimentos = planilha_master.worksheet("Recebimentos_Master")
                df_pagamentos = pd.DataFrame(aba_recebimentos.get_all_records())
                if 'Unidade' not in df_pagamentos.columns: df_pagamentos['Unidade'] = ""
                
                df_pagamentos['Valor_Recebido'] = df_pagamentos['Valor_Recebido'].apply(safe_to_float)
                df_pagamentos['Chave'] = df_pagamentos['Cliente'].astype(str).str.strip()
                df_pagamentos.loc[df_pagamentos['Unidade'].str.strip() != "", 'Chave'] = df_pagamentos['Chave'] + " - " + df_pagamentos['Unidade'].astype(str).str.strip()
                
                df_resumo_pago = df_pagamentos.groupby('Chave')['Valor_Recebido'].sum().reset_index()
                df_resumo_pago.rename(columns={'Valor_Recebido': 'Total_Pago'}, inplace=True)

                # B. CONTRATOS
                if not df_lista.empty:
                    df_contratos = df_lista.copy()
                    df_contratos['VALOR DA UNIDADE'] = df_contratos['VALOR DA UNIDADE'].apply(safe_to_float)
                    
                    df_contratos['Chave'] = df_contratos['NOME DO ADQUIRENTE'].astype(str).str.strip()
                    df_contratos.loc[df_contratos['DESCRIÇÃO RESUMIDA DA UNIDADE'].str.strip() != "", 'Chave'] = df_contratos['Chave'] + " - " + df_contratos['DESCRIÇÃO RESUMIDA DA UNIDADE'].astype(str).str.strip()
                    df_contratos = df_contratos[df_contratos['Chave'] != ""].drop_duplicates(subset=['Chave'])
                else:
                    df_contratos = pd.DataFrame(columns=['Chave', 'VALOR DA UNIDADE'])

                # C. CORREÇÕES
                try:
                    aba_ajustes = planilha_master.worksheet("Ajustes_Contratos")
                    df_aj = pd.DataFrame(aba_ajustes.get_all_records())
                    df_aj['Valor_Ajuste'] = df_aj['Valor_Ajuste'].apply(safe_to_float)
                    
                    df_resumo_ajustes = df_aj.groupby('Cliente')['Valor_Ajuste'].sum().reset_index()
                    df_resumo_ajustes.rename(columns={'Cliente': 'Chave'}, inplace=True)
                except:
                    df_resumo_ajustes = pd.DataFrame(columns=['Chave', 'Valor_Ajuste'])

                # D. O GRANDE CRUZAMENTO
                df_dash = pd.merge(df_contratos, df_resumo_ajustes, on='Chave', how='left').fillna(0)
                df_dash = pd.merge(df_dash, df_resumo_pago, on='Chave', how='left').fillna(0)

                df_dash['Valor_Atualizado'] = df_dash['VALOR DA UNIDADE'] + df_dash['Valor_Ajuste']
                df_dash['Saldo_Devedor'] = df_dash['Valor_Atualizado'] - df_dash['Total_Pago']
                df_dash['Saldo_Devedor'] = df_dash['Saldo_Devedor'].apply(lambda x: x if x > 0 else 0)

                vgv_total = df_dash['Valor_Atualizado'].sum()
                recebido_total = df_dash['Total_Pago'].sum()
                devedor_total = df_dash['Saldo_Devedor'].sum()

                # --- KPIs PRINCIPAIS ---
                c1, c2, c3 = st.columns(3)
                c1.metric("💰 VGV Atualizado", f"R$ {vgv_total:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'))
                c2.metric("✅ Caixa Realizado", f"R$ {recebido_total:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'))
                c3.metric("⚠️ Saldo Devedor", f"R$ {devedor_total:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'))
                st.divider()

                # ==========================================
                # ABAS DE ANÁLISE, AUDITORIA E BOLETOS
                # ==========================================
                aba_tabela, aba_graficos, aba_extrato, aba_boletos = st.tabs([
                    "🏢 Detalhamento (Investidores)", 
                    "📈 Análise Gráfica",
                    "🔍 Extrato do Investidor",
                    "🧾 Emissão de Boletos"
                ])
                
                with aba_tabela:
                    st.subheader("🏛️ Resumo Tributário Acumulado")
                    pis_total = recebido_total * 0.0065
                    cofins_total = recebido_total * 0.03
                    csll_total = recebido_total * 0.0108
                    ir_normal_total = recebido_total * 0.012
                    
                    df_pagamentos['Data_FMT'] = pd.to_datetime(df_pagamentos['Data_Pagamento'], errors='coerce', dayfirst=True)
                    df_pagamentos['Mes_Ano'] = df_pagamentos['Data_FMT'].dt.to_period('M')
                    receita_mensal = df_pagamentos.groupby('Mes_Ano')['Valor_Recebido'].sum().reset_index()
                    
                    ir_adicional_total = 0
                    for _, row in receita_mensal.iterrows():
                        base_pres = row['Valor_Recebido'] * 0.08
                        if base_pres > 20000:
                            ir_adicional_total += (base_pres - 20000) * 0.10

                    t1, t2, t3, t4, t5 = st.columns(5)
                    t1.metric("PIS (0,65%)", f"R$ {pis_total:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'))
                    t2.metric("COFINS (3%)", f"R$ {cofins_total:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'))
                    t3.metric("CSLL (1,08%)", f"R$ {csll_total:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'))
                    t4.metric("IR (1,2%)", f"R$ {ir_normal_total:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'))
                    t5.metric("IR Adicional", f"R$ {ir_adicional_total:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'))

                    st.divider()
                    st.subheader("Situação Individualizada por Contrato")
                    df_visual = df_dash[['Chave', 'VALOR DA UNIDADE', 'Valor_Ajuste', 'Valor_Atualizado', 'Total_Pago', 'Saldo_Devedor']].copy()
                    
                    df_visual['Progresso'] = df_visual.apply(
                        lambda row: (row['Total_Pago'] / row['Valor_Atualizado'] * 100) if row['Valor_Atualizado'] > 0 else 0.0, 
                        axis=1
                    )
                    df_visual['Progresso'] = df_visual['Progresso'].clip(upper=100)
                    
                    st.markdown("""<style>[data-testid="stDataFrame"] {zoom: 1.15;}</style>""", unsafe_allow_html=True)
                    
                    st.dataframe(
                        df_visual,
                        column_config={
                            "Chave": st.column_config.TextColumn("Contrato (Investidor/Unidade)"),
                            "VALOR DA UNIDADE": st.column_config.NumberColumn("Contrato Original", format="R$ %.2f"),
                            "Valor_Ajuste": st.column_config.NumberColumn("+ Correções", format="R$ %.2f"),
                            "Valor_Atualizado": st.column_config.NumberColumn("Total Atualizado", format="R$ %.2f"),
                            "Total_Pago": st.column_config.NumberColumn("Pago", format="R$ %.2f"),
                            "Saldo_Devedor": st.column_config.NumberColumn("Saldo Devedor", format="R$ %.2f"),
                            "Progresso": st.column_config.ProgressColumn("Quitação", format="%.1f %%", min_value=0, max_value=100)
                        },
                        hide_index=True, use_container_width=True
                    )
                
                with aba_graficos:
                    st.subheader("Evolução do Caixa e Carga Tributária")
                    if not df_pagamentos.empty:
                        df_grafico = df_pagamentos.dropna(subset=['Data_FMT']).copy()
                        df_grafico['Mes_Periodo'] = df_grafico['Data_FMT'].dt.to_period('M')
                        df_mensal_graf = df_grafico.groupby('Mes_Periodo')['Valor_Recebido'].sum().reset_index()

                        df_mensal_graf['Impostos_Comuns'] = df_mensal_graf['Valor_Recebido'] * (0.0065 + 0.03 + 0.0108 + 0.012)
                        df_mensal_graf['Base_IR'] = df_mensal_graf['Valor_Recebido'] * 0.08
                        df_mensal_graf['IR_Adicional'] = df_mensal_graf['Base_IR'].apply(lambda x: (x - 20000) * 0.10 if x > 20000 else 0)

                        df_mensal_graf['Total_Tributos'] = df_mensal_graf['Impostos_Comuns'] + df_mensal_graf['IR_Adicional']
                        df_mensal_graf['Caixa_Livre_Liquido'] = df_mensal_graf['Valor_Recebido'] - df_mensal_graf['Total_Tributos']
                        df_mensal_graf['Mês/Ano'] = df_mensal_graf['Mes_Periodo'].dt.strftime('%m/%Y')

                        st.markdown("### 📅 Desempenho Mensal")
                        df_chart_m = df_mensal_graf[['Mês/Ano', 'Caixa_Livre_Liquido', 'Total_Tributos']].set_index('Mês/Ano')
                        st.bar_chart(df_chart_m, color=["#1f77b4", "#d62728"], height=350) 
                        
                        st.divider()

                        df_grafico['Ano'] = df_grafico['Data_FMT'].dt.year.astype(str)
                        df_anual = df_grafico.groupby('Ano')['Valor_Recebido'].sum().reset_index()

                        df_anual['Impostos_Comuns'] = df_anual['Valor_Recebido'] * (0.0065 + 0.03 + 0.0108 + 0.012)
                        soma_ir_adicional_anual = df_mensal_graf.groupby(df_mensal_graf['Mes_Periodo'].dt.year)['IR_Adicional'].sum().reset_index()
                        soma_ir_adicional_anual.columns = ['Ano', 'IR_Adicional']
                        soma_ir_adicional_anual['Ano'] = soma_ir_adicional_anual['Ano'].astype(str)

                        df_anual = pd.merge(df_anual, soma_ir_adicional_anual, on='Ano', how='left').fillna(0)
                        df_anual['Total_Tributos'] = df_anual['Impostos_Comuns'] + df_anual['IR_Adicional']
                        df_anual['Caixa_Livre_Liquido'] = df_anual['Valor_Recebido'] - df_anual['Total_Tributos']

                        st.markdown("### 📆 Consolidado Anual")
                        df_chart_a = df_anual[['Ano', 'Caixa_Livre_Liquido', 'Total_Tributos']].set_index('Ano')
                        st.bar_chart(df_chart_a, color=["#2ca02c", "#d62728"], height=350)
                    else:
                        st.info("Ainda não há pagamentos registrados para gerar os gráficos.")

                with aba_extrato:
                    st.subheader("🔎 Auditoria: Extrato de Pagamentos e Ajustes")
                    st.markdown("Selecione um contrato para verificar todas as parcelas e taxas que o sistema encontrou para ele.")
                    
                    lista_chaves = sorted([str(c) for c in df_dash['Chave'].unique() if str(c).strip() != ""])
                    cliente_auditoria = st.selectbox("Selecione o Contrato para auditar:", lista_chaves)
                    
                    if cliente_auditoria:
                        saldo_atual = df_dash.loc[df_dash['Chave'] == cliente_auditoria, 'Saldo_Devedor'].values
                        saldo_exibicao = saldo_atual[0] if len(saldo_atual) > 0 else 0.0
                        st.warning(f"⚠️ Saldo Devedor Atualizado: **R$ {saldo_exibicao:,.2f}**".replace(',', '_').replace('.', ',').replace('_', '.'))
                        
                        col_ext1, col_ext2 = st.columns(2)
                        
                        with col_ext1:
                            st.markdown("#### 📥 Pagamentos Identificados")
                            df_pag_cli = df_pagamentos[df_pagamentos['Chave'] == cliente_auditoria].copy()
                            
                            if not df_pag_cli.empty:
                                soma_pags = df_pag_cli['Valor_Recebido'].sum()
                                st.success(f"Soma dos Pagamentos: R$ {soma_pags:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'))
                                
                                try:
                                    df_disp_pag = df_pag_cli[['Data_Pagamento', 'Valor_Recebido']].copy()
                                    df_disp_pag['Data_Real'] = pd.to_datetime(df_disp_pag['Data_Pagamento'], errors='coerce', format='mixed')
                                    df_disp_pag = df_disp_pag.sort_values(by='Data_Real', ascending=True)
                                    df_disp_pag['Data_Pagamento'] = df_disp_pag['Data_Real'].dt.strftime('%d/%m/%Y')
                                    df_disp_pag = df_disp_pag[['Data_Pagamento', 'Valor_Recebido']]
                                except:
                                    pass
                                
                                st.dataframe(
                                    df_disp_pag, 
                                    column_config={
                                        "Data_Pagamento": st.column_config.TextColumn("Data do Pagamento"),
                                        "Valor_Recebido": st.column_config.NumberColumn("Valor Pago", format="R$ %.2f")
                                    },
                                    hide_index=True, use_container_width=True
                                )
                            else:
                                st.warning("Nenhum pagamento localizado para esta chave.")

                        with col_ext2:
                            st.markdown("#### ⚙️ Ajustes e Correções")
                            try:
                                if 'df_ajustes' in locals() or 'df_ajustes' in globals():
                                    df_ajuste_cli = df_ajustes[df_ajustes['Chave'] == cliente_auditoria].copy()
                                    if not df_ajuste_cli.empty:
                                        soma_ajustes = df_ajuste_cli['Valor_Ajuste'].sum()
                                        st.info(f"Soma das Correções: R$ {soma_ajustes:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'))
                                        df_disp_ajuste = df_ajuste_cli[['Data_Ajuste', 'Valor_Ajuste', 'Motivo']].copy()
                                        st.dataframe(df_disp_ajuste, hide_index=True, use_container_width=True)
                                    else:
                                        st.info("Nenhuma correção/taxa lançada para este contrato.")
                                else:
                                    st.caption("A base de dados de correções não está carregada.")
                            except:
                                st.caption("Não foi possível carregar o histórico de ajustes.")

                # ==========================================
                # ABA DE BOLETOS 
                # ==========================================
                with aba_boletos:
                    st.subheader("🧾 Emissão Lote de Boletos - Banco Inter")
                    st.markdown("Selecione os clientes na tabela abaixo marcando a caixa **'Gerar?'** e clique no botão para emitir.")
                    
                    try:
                        # Corrigido: Usando a base df_dash (se a sua base de clientes tiver outro nome específico, basta trocar 'df_dash' abaixo)
                        df_boletos_tela = df_dash.copy()
                        if 'Emitir' not in df_boletos_tela.columns:
                            df_boletos_tela.insert(0, 'Emitir', False)
                            
                        df_editado = st.data_editor(
                            df_boletos_tela,
                            column_config={
                                "Emitir": st.column_config.CheckboxColumn("Gerar?", default=False)
                            },
                            hide_index=True,
                            use_container_width=True
                        )
                        
                        col_btn1, col_btn2 = st.columns([2, 2])
                        
                        with col_btn2:
                            if st.button("🚀 Processar Boletos Selecionados", type="primary"):
                                clientes_selecionados = df_editado[df_editado['Emitir'] == True]
                                
                                if clientes_selecionados.empty:
                                    st.warning("Selecione pelo menos um cliente marcando a caixa 'Gerar?'.")
                                else:
                                    st.success(f"Emitindo {len(clientes_selecionados)} boleto(s)...")
                                    st.session_state.boletos_processados = []
                                    
                                    import tempfile
                                    import base64
                                    import os
                                    import datetime as dt
                                    from decimal import Decimal
                                    import re
                                    import requests
                                    import io
                                    import math
                                    
                                    # ==========================================
                                    # 1. CONFIGURAÇÃO GOOGLE DRIVE (LINK PÚBLICO)
                                    # ==========================================
                                    PASTA_DRIVE_ID = "1yFTfudMhSBCfsmLx4q3o1krg7LZLWmiy"
                                    drive_service = None
                                    try:
                                        from google.oauth2.service_account import Credentials
                                        from googleapiclient.discovery import build
                                        from googleapiclient.http import MediaIoBaseUpload
                                        
                                        creds_gcp = Credentials.from_service_account_info(
                                            st.secrets["gcp_service_account"],
                                            scopes=['https://www.googleapis.com/auth/drive']
                                        )
                                        drive_service = build('drive', 'v3', credentials=creds_gcp)
                                    except Exception as e:
                                        st.warning(f"Drive Desconectado (O PDF não será salvo na nuvem). Erro: {e}")
                                    
                                    # ==========================================
                                    # 2. CONFIGURAÇÃO BANCO INTER
                                    # ==========================================
                                    caminho_pfx_temp = None
                                    try:
                                        creds_inter = st.secrets["inter_api"]
                                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pfx') as tmp_pfx:
                                            tmp_pfx.write(base64.b64decode(creds_inter["pfx_base64"]))
                                            caminho_pfx_temp = tmp_pfx.name
                                            
                                        from inter_sdk_python.InterSdk import InterSdk
                                        from inter_sdk_python.billing.models.BillingIssueRequest import BillingIssueRequest
                                        from inter_sdk_python.billing.models.Person import Person
                                        try:
                                            from inter_sdk_python.commons.models.PersonType import PersonType
                                        except:
                                            from enum import Enum
                                            class PersonType(Enum):
                                                FISICA = "FISICA"
                                                JURIDICA = "JURIDICA"
                                                
                                        sdk = InterSdk(
                                            "PRODUCTION",
                                            creds_inter["client_id"],
                                            creds_inter["client_secret"],
                                            caminho_pfx_temp,
                                            creds_inter["pfx_senha"]
                                        )
                                        sdk.set_account(creds_inter["conta_corrente"].replace("-",""))
                                        
                                        # ==========================================
                                        # 3. MOTOR DE EMISSÃO E LINKS
                                        # ==========================================
                                        for idx, row in clientes_selecionados.iterrows():
                                            nome_completo = str(row['Nome_Cliente'].split('-')[0]).strip()[:100]
                                            zap = str(row['WhatsApp']).replace(' ', '').replace('-', '').replace('(', '').replace(')', '')
                                            cep_limpo = re.sub(r'\D', '', str(row['CEP']))
                                            valor = float(row['Valor_Parcela'])
                                            vencimento = row['Vencimento']
                                            numero = str(row['Numero']).strip() if str(row['Numero']).strip() != "" else "0"
                                            cpf_cnpj_limpo = re.sub(r'\D', '', str(row['CPF_CNPJ']))
                                            
                                            saldo_devedor = float(row.get('Saldo_Devedor', 0))
                                            
                                            segundos = str(int(dt.datetime.now().timestamp()))[-6:]
                                            controle = f"JL{idx}{segundos}S"[:15]
                                            
                                            rua_encontrada, bairro_encontrado, cidade_encontrada, uf_encontrada = "Logradouro", "Bairro", "Cidade", "MG"
                                            if len(cep_limpo) == 8:
                                                try:
                                                    resp_cep = requests.get(f"https://viacep.com.br/ws/{cep_limpo}/json/", timeout=5)
                                                    if resp_cep.status_code == 200 and 'erro' not in resp_cep.json():
                                                        d_cep = resp_cep.json()
                                                        if d_cep.get('logradouro'): rua_encontrada = d_cep.get('logradouro')
                                                        if d_cep.get('bairro'): bairro_encontrado = d_cep.get('bairro')
                                                        if d_cep.get('localidade'): cidade_encontrada = d_cep.get('localidade')
                                                        if d_cep.get('uf'): uf_encontrada = d_cep.get('uf')
                                                except: pass
                                                
                                            try:
                                                data_vencimento = dt.datetime.strptime(vencimento, '%d/%m/%Y').strftime('%Y-%m-%d')
                                                
                                                pagador = Person()
                                                pagador.nome = pagador.name = nome_completo
                                                pagador.cpf_cnpj = pagador.cpfCnpj = cpf_cnpj_limpo
                                                pagador.cep = pagador.zip_code = pagador.zipCode = cep_limpo if len(cep_limpo) == 8 else "30000000"
                                                pagador.numero = pagador.number = numero
                                                pagador.endereco = pagador.address = pagador.logradouro = rua_encontrada
                                                pagador.cidade = pagador.city = cidade_encontrada
                                                pagador.uf = pagador.state = uf_encontrada
                                                pagador.bairro = pagador.neighborhood = bairro_encontrado
                                                if str(row['Complemento']).strip(): pagador.complemento = pagador.complement = str(row['Complemento']).strip()
                                                pagador.tipo_pessoa = pagador.tipoPessoa = pagador.person_type = pagador.personType = PersonType.FISICA if len(cpf_cnpj_limpo) <= 11 else PersonType.JURIDICA

                                                boleto = BillingIssueRequest()
                                                boleto.seu_numero = boleto.seuNumero = boleto.your_number = boleto.yourNumber = controle
                                                boleto.valor_nominal = boleto.valorNominal = boleto.nominal_value = boleto.nominalValue = Decimal(str(round(valor, 2)))
                                                boleto.data_vencimento = boleto.dataVencimento = boleto.due_date = boleto.dueDate = data_vencimento
                                                boleto.pagador = boleto.payer = pagador
                                                
                                                # ==========================================
                                                # REGRAS DE ATRASO (MULTA E MORA)
                                                # ==========================================
                                                boleto.num_dias_agenda = boleto.numDiasAgenda = boleto.scheduled_days = 30
                                                try:
                                                    from inter_sdk_python.billing.models.Fine import Fine
                                                    from inter_sdk_python.billing.models.Mora import Mora
                                                    
                                                    multa = Fine()
                                                    multa.codigo_multa = multa.codigoMulta = "PERCENTUAL"
                                                    multa.taxa = Decimal("2.00")
                                                    boleto.multa = boleto.fine = multa
                                                    
                                                    mora = Mora()
                                                    mora.codigo_mora = mora.codigoMora = "TAXA_MENSAL"
                                                    mora.taxa = Decimal("1.00")
                                                    boleto.mora = mora
                                                except Exception as err_regra:
                                                    st.toast(f"Aviso: As regras de multa não foram aplicadas. {err_regra}")

                                                # Emissão
                                                res = sdk.billing().issue_billing(boleto)
                                                n_num = getattr(res, 'nossoNumero', None) or getattr(res, 'nosso_numero', None) or (res.get('nossoNumero') if isinstance(res, dict) else None) or (res.get('request_code') if isinstance(res, dict) else getattr(res, 'request_code', None))
                                                
                                                if n_num:
                                                    st.info(f"⏳ Renderizando boleto (Cód: {str(n_num)[:8]}...).")
                                                    import time
                                                    time.sleep(4)
                                                    
                                                    pdf_path = os.path.join(tempfile.gettempdir(), f"{controle}.pdf")
                                                    link_do_drive = ""
                                                    
                                                    try:
                                                        sdk.billing().retrieve_billing_pdf(str(n_num), file=pdf_path)
                                                        with open(pdf_path, "rb") as f:
                                                            pdf_bytes = f.read()
                                                            
                                                        # UPLOAD E GERAÇÃO DE LINK PÚBLICO
                                                        if drive_service:
                                                            nome_arquivo = f"Boleto_JL_{nome_completo.replace(' ', '_')}_{vencimento.replace('/', '-')}.pdf"
                                                            media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype='application/pdf', resumable=True)
                                                            meta = {'name': nome_arquivo, 'parents': [PASTA_DRIVE_ID]}
                                                            
                                                            arquivo_drive = drive_service.files().create(body=meta, media_body=media, fields='id').execute()
                                                            file_id = arquivo_drive.get('id')
                                                            
                                                            drive_service.permissions().create(
                                                                fileId=file_id,
                                                                body={'type': 'anyone', 'role': 'reader'}
                                                            ).execute()
                                                            
                                                            file_info = drive_service.files().get(fileId=file_id, fields='webViewLink').execute()
                                                            link_do_drive = file_info.get('webViewLink')
                                                        
                                                        # MONTAGEM DA MENSAGEM DO ZAP
                                                        parcelas_restantes = math.ceil(saldo_devedor / valor) if valor > 0 else 0
                                                        saldo_formatado = f"{saldo_devedor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                                                        
                                                        texto_msg = (
                                                            f"Olá {nome_completo}, tudo bem? "
                                                            f"Segue o link para baixar o seu boleto da J&L Incorporadora no valor de R$ {valor:.2f} com vencimento para {vencimento}.\n\n"
                                                            f"👉 *Acessar Boleto:* {link_do_drive if link_do_drive else '(Baixe o PDF acima)'}\n\n"
                                                            f"Informamos que o seu Saldo Devedor atualizado é de R$ {saldo_formatado}, "
                                                            f"restando aproximadamente {parcelas_restantes} parcela(s) para a quitação do seu contrato."
                                                        )
                                                        
                                                        link_wa = f"https://wa.me/55{zap}?text={requests.utils.quote(texto_msg)}" if len(zap) >= 10 else None
                                                        
                                                        st.session_state.boletos_processados.append({
                                                            "nome": nome_completo,
                                                            "arquivo": pdf_bytes,
                                                            "link_wa": link_wa,
                                                            "link_drive": link_do_drive
                                                        })
                                                        st.success(f"🎉 PDF de {nome_completo} gerado e salvo!")
                                                    except Exception as erro_pdf:
                                                        st.warning(f"⚠️ Boleto gerado, mas erro no PDF/Drive: {erro_pdf}")
                                                else:
                                                    st.error(f"❌ Banco não retornou rastreio para {nome_completo}.")
                                            
                                            except Exception as erro_emissao:
                                                msg = str(erro_emissao)
                                                if hasattr(erro_emissao, 'error') and erro_emissao.error: msg = erro_emissao.error.detail
                                                st.error(f"❌ Falha ao emitir {nome_completo}: {msg}")
                                                
                                    finally:
                                        if caminho_pfx_temp and os.path.exists(caminho_pfx_temp):
                                            os.remove(caminho_pfx_temp)

                        # EXIBIÇÃO FINAL DOS BOLETOS
                        if st.session_state.get("boletos_processados"):
                            st.divider()
                            st.markdown("### 🗂️ Boletos Prontos para Envio")
                            
                            for i, bol in enumerate(st.session_state.boletos_processados):
                                colA, colB, colC = st.columns([3, 2, 2])
                                colA.markdown(f"**{bol['nome']}**")
                                
                                with colB:
                                    if bol['link_drive']:
                                        st.markdown(f"[🔗 Ver PDF no Drive]({bol['link_drive']})")
                                    else:
                                        st.caption("Salvo apenas na memória")
                                
                                with colC:
                                    if bol['link_wa']:
                                        st.markdown(f"**[📲 Enviar Mensagem Completa no WhatsApp]({bol['link_wa']})**")
                                    else:
                                        st.caption("Sem telefone cadastrado.")
            except Exception as e:
                 st.error(f"Erro na aba de Boletos: {e}")
                        

    
    # --- TELA 5: FATURAMENTO (BOLETOS) ---
    elif modulo == "💸 Faturamento e Boletos":
        st.title("💸 Gestão de Recebíveis e Emissão")
        st.markdown("Leia os clientes direto da planilha do Google Drive e emita os boletos.")
        
        # Lê a planilha de 2026 (usando a variável que já estava no seu código original)
        mes_faturamento = st.selectbox("Selecione a Aba (Mês):", ["Julho", "Agosto", "Setembro", "Outubro"])
        
        if st.button("📥 Carregar Clientes da Planilha"):
            with st.spinner("Conectando ao Google Drive..."):
                try:
                    client_gspread = obter_cliente_sheets()
                    planilha_2026 = client_gspread.open_by_key(ID_PLANILHA_Recebimento_Wilson_Moreira_2026)
                    aba_mes = planilha_2026.worksheet(mes_faturamento)
                    
                    dados = aba_mes.get_all_records()
                    if dados:
                        df_clientes = pd.DataFrame(dados)
                        st.session_state['df_boletos_pendentes'] = df_clientes
                        st.success("Planilha carregada com sucesso!")
                    else:
                        st.warning("A aba selecionada está vazia.")
                except Exception as e:
                    st.error(f"Erro ao ler a planilha: {e}")
                    
        if 'df_boletos_pendentes' in st.session_state:
            df_mostrar = st.session_state['df_boletos_pendentes']
            st.dataframe(df_mostrar)
            
            st.warning("⚠️ Certifique-se de que a planilha possui as colunas exatas: Nome_Cliente, CPF_CNPJ, seuNumero, Vencimento, Valor, CEP, Número.")
            
            if st.button("🚀 Emitir Boletos Selecionados (Banco Inter)", type="primary"):
                with st.spinner("Processando lote de boletos no Banco Inter. Isso pode levar alguns minutos..."):
                    resultados = emitir_boletos_lote(df_mostrar)
                    
                    if resultados:
                        st.subheader("📊 Resultado da Emissão")
                        for r in resultados:
                            if r["Status"] == "Sucesso":
                                st.success(f"✅ {r['Cliente']} - Nosso Número: {r['Nosso Número']}")
                                st.download_button(
                                    label=f"⬇️ Baixar PDF - {r['Cliente']}",
                                    data=r["PDF"],
                                    file_name=f"Boleto_{r['Cliente']}.pdf",
                                    mime="application/pdf"
                                )
                            else:
                                st.error(f"❌ {r['Cliente']} - Falhou: {r['Motivo']}")
# --- TELA SECRETA 6: TRANSBORDO DE DADOS (VERSÃO DEFINITIVA LINHA 3) ---


    elif modulo == "⚙️ Engenharia (Transbordo)":
        st.title("⚙️ Transbordo de Recebíveis (Leitura Mês a Mês)")
        st.markdown("Varrendo o histórico detalhado para criar o **Livro Razão Definitivo**.")

        # BOTÃO 1: LER AS PLANILHAS
        if st.button("🚀 Iniciar Transbordo Histórico com Datas", type="primary"):
            with st.spinner("Lendo planilhas. Sistema de anti-queda ativado..."):
                planilhas_ids = {
                    "2022": ID_PLANILHA_Recebimento_Wilson_Moreira_2022,
                    "2023": ID_PLANILHA_Recebimento_Wilson_Moreira_2023,
                    "2024": ID_PLANILHA_Recebimento_Wilson_Moreira_2024,
                    "2025": ID_PLANILHA_Recebimento_Wilson_Moreira_2025,
                    "2026": ID_PLANILHA_Recebimento_Wilson_Moreira_2026
                }
                
                client_gspread = obter_cliente_sheets()
                if not client_gspread:
                    st.error("Sem conexão com o Google Sheets.")
                    st.stop()
                    
                banco_de_dados_geral = []
                log_abas_lidas = []
                import re
                import time

                for ano, sheet_id in planilhas_ids.items():
                    # SISTEMA DE RETENTATIVA (Tenta 3 vezes se o Google der erro 503)
                    sucesso_leitura = False
                    tentativas = 0
                    
                    while not sucesso_leitura and tentativas < 3:
                        try:
                            planilha = client_gspread.open_by_key(sheet_id)
                            abas = planilha.worksheets()
                            sucesso_leitura = True # Se passou daqui, o Google liberou
                            
                            for aba in abas:
                                nome_aba = aba.title.lower()
                                if "resumo" in nome_aba or "planilha" in nome_aba or "totais" in nome_aba: continue
                                    
                                linhas_brutas = aba.get_all_values()
                                if len(linhas_brutas) < 2: continue
                                
                                linha_cab = -1
                                cabecalho_limpo = []
                                for i, row in enumerate(linhas_brutas):
                                    cols = [re.sub(r'\s+', ' ', str(c)).strip().upper() for c in row]
                                    if "NOME DO ADQUIRENTE" in cols and "VALOR" in cols and "DATA" in cols:
                                        linha_cab = i
                                        cabecalho_limpo = cols
                                        break
                                        
                                if linha_cab == -1: continue 
                                    
                                log_abas_lidas.append(f"{ano} - {aba.title}")
                                
                                try:
                                    idx_cliente = cabecalho_limpo.index('NOME DO ADQUIRENTE')
                                    idx_cpf = cabecalho_limpo.index('CPF DO ADQUIRENTE') if 'CPF DO ADQUIRENTE' in cabecalho_limpo else -1
                                    idx_data = cabecalho_limpo.index('DATA')
                                    idx_valor = cabecalho_limpo.index('VALOR')
                                    idx_contrato = cabecalho_limpo.index('Nº CONTRATO') if 'Nº CONTRATO' in cabecalho_limpo else -1
                                    idx_unidade = cabecalho_limpo.index('DESCRIÇÃO RESUMIDA DA UNIDADE') if 'DESCRIÇÃO RESUMIDA DA UNIDADE' in cabecalho_limpo else -1
                                except ValueError:
                                    continue

                                dados_tabela = linhas_brutas[linha_cab + 1:]
                                
                                for row in dados_tabela:
                                    row = row + [''] * (len(cabecalho_limpo) - len(row))
                                    cliente = str(row[idx_cliente]).strip()
                                    data_pagamento = str(row[idx_data]).strip()
                                    valor_bruto = str(row[idx_valor]).strip()
                                    
                                    if not cliente or not data_pagamento or data_pagamento.lower() in ['nan', '-', '']: continue
                                        
                                    v_limpo = valor_bruto.replace('R$', '').replace('.', '').replace(',', '.').strip()
                                    try: v_float = float(v_limpo)
                                    except: v_float = 0.0
                                    
                                    if v_float > 0:
                                        banco_de_dados_geral.append({
                                            "Ano_Origem": ano, 
                                            "Mes_Aba": aba.title, 
                                            "Cliente": cliente,
                                            "CPF_CNPJ": str(row[idx_cpf]).strip() if idx_cpf != -1 else "",
                                            "Contrato": str(row[idx_contrato]).strip() if idx_contrato != -1 else "",
                                            "Unidade": str(row[idx_unidade]).strip() if idx_unidade != -1 else "",
                                            "Data_Pagamento": data_pagamento, 
                                            "Valor_Recebido": v_float,
                                            "ID_Transacao_Banco": "",      # Coluna preparada para os novos lançamentos
                                            "Origem_Lancamento": "Histórico" # Identificador de auditoria
                                        })
                        except Exception as e:
                            tentativas += 1
                            if tentativas < 3:
                                st.warning(f"O Google falhou no ano {ano} (Erro 503). Tentando novamente em 3 segundos... (Tentativa {tentativas}/3)")
                                time.sleep(3)
                            else:
                                st.error(f"Aviso: Erro definitivo ao processar o ano {ano}. Detalhe: {e}")

                if banco_de_dados_geral:
                    df_final = pd.DataFrame(banco_de_dados_geral)
                    try:
                        df_final['Data_Pagamento_FMT'] = pd.to_datetime(df_final['Data_Pagamento'], errors='coerce', dayfirst=True)
                        df_final = df_final.sort_values(by='Data_Pagamento_FMT').drop(columns=['Data_Pagamento_FMT'])
                    except: pass
                    
                    st.session_state['df_transbordo'] = df_final
                else:
                    st.warning("Nenhuma linha encontrada.")

        # SE O DADO JÁ FOI LIDO, MOSTRA A TABELA E O BOTÃO DE SALVAR
        if 'df_transbordo' in st.session_state:
            df_mostrar = st.session_state['df_transbordo']
            st.success(f"✅ Transbordo Concluído! {len(df_mostrar)} liquidações exatas extraídas do histórico.")
            st.dataframe(df_mostrar, hide_index=True)
            
            st.divider()
            
            # BOTÃO 2: GRAVAR NO GOOGLE SHEETS
            if st.button("💾 GRAVAR NA PLANILHA MASTER", type="primary", use_container_width=True):
                with st.spinner("Criando a aba 'Recebimentos_Master' e gravando o histórico..."):
                    try:
                        client_gspread = obter_cliente_sheets()
                        planilha_master = client_gspread.open_by_key(ID_PLANILHA_MASTER)
                        
                        try:
                            aba_master = planilha_master.worksheet("Recebimentos_Master")
                            aba_master.clear() 
                        except:
                            aba_master = planilha_master.add_worksheet(title="Recebimentos_Master", rows="1000", cols="10")
                        
                        df_salvar = df_mostrar.fillna("").astype(str)
                        dados_para_inserir = [df_salvar.columns.values.tolist()] + df_salvar.values.tolist()
                        
                        aba_master.append_rows(dados_para_inserir)
                        
                        st.success("🎉 Golaço! Os dados foram gravados com sucesso na sua Planilha Master.")
                        st.balloons()
                        
                        del st.session_state['df_transbordo']
                    except Exception as e:
                        st.error(f"Erro ao tentar gravar na planilha: {e}")
