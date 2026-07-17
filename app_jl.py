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

from inter_sdk_python.billing.models.BillingIssueRequest import BillingIssueRequest
from inter_sdk_python.billing.models.Person import Person 
try:
    from inter_sdk_python.commons.models.PersonType import PersonType
except:
    from enum import Enum
    class PersonType(Enum):
        FISICA = "FISICA"
        JURIDICA = "JURIDICA"

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
    # --- TELA 4: DASHBOARD ---
    elif modulo == "📊 Dashboard Executivo":
        st.title("📊 Visão Consolidada das Obras")
        st.write("Acompanhamento de fluxo financeiro.")
        st.title("📊 Visão Consolidada das Obras")
        st.write("Acompanhamento de fluxo financeiro.")
        
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
        st.markdown("Varrendo o histórico detalhado em busca das **Datas** e **Valores** exatos para o Livro Razão.")

        if st.button("🚀 Iniciar Transbordo Histórico com Datas", type="primary"):
            with st.spinner("Motor blindado ativado: Caçando cabeçalhos e ignorando colunas vazias..."):
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

                for ano, sheet_id in planilhas_ids.items():
                    try:
                        planilha = client_gspread.open_by_key(sheet_id)
                        abas = planilha.worksheets()
                        
                        for aba in abas:
                            nome_aba = aba.title.lower()
                            
                            # Ignora as abas totalizadoras
                            if "resumo" in nome_aba or "planilha" in nome_aba or "totais" in nome_aba:
                                continue
                                
                            linhas_brutas = aba.get_all_values()
                            if len(linhas_brutas) < 2: continue
                            
                            # 1. O CAÇADOR DE CABEÇALHOS (Agora buscando "DATA" e "VALOR")
                            linha_cab = -1
                            cabecalho_limpo = []
                            for i, row in enumerate(linhas_brutas):
                                cols = [re.sub(r'\s+', ' ', str(c)).strip().upper() for c in row]
                                # Correção: Procurando os nomes exatos da sua imagem
                                if "NOME DO ADQUIRENTE" in cols and "VALOR" in cols and "DATA" in cols:
                                    linha_cab = i
                                    cabecalho_limpo = cols
                                    break
                                    
                            if linha_cab == -1:
                                continue 
                                
                            log_abas_lidas.append(f"{ano} - {aba.title}")
                            
                            # 2. MAPEAMENTO DE ÍNDICES 
                            try:
                                idx_cliente = cabecalho_limpo.index('NOME DO ADQUIRENTE')
                                idx_data = cabecalho_limpo.index('DATA')
                                idx_valor = cabecalho_limpo.index('VALOR')
                                idx_contrato = cabecalho_limpo.index('Nº CONTRATO') if 'Nº CONTRATO' in cabecalho_limpo else -1
                                idx_unidade = cabecalho_limpo.index('DESCRIÇÃO RESUMIDA DA UNIDADE') if 'DESCRIÇÃO RESUMIDA DA UNIDADE' in cabecalho_limpo else -1
                            except ValueError:
                                continue

                            dados_tabela = linhas_brutas[linha_cab + 1:]
                            
                            # 3. EXTRAÇÃO DIRETA 
                            for row in dados_tabela:
                                # Preenche a linha com vazio se for menor que o cabeçalho
                                row = row + [''] * (len(cabecalho_limpo) - len(row))
                                
                                cliente = str(row[idx_cliente]).strip()
                                data_pagamento = str(row[idx_data]).strip()
                                valor_bruto = str(row[idx_valor]).strip()
                                
                                # Ignora se não pagou (Célula vazia ou traço)
                                if not cliente or not data_pagamento or data_pagamento.lower() in ['nan', '-', '']: 
                                    continue
                                    
                                v_limpo = valor_bruto.replace('R$', '').replace('.', '').replace(',', '.').strip()
                                try: v_float = float(v_limpo)
                                except: v_float = 0.0
                                
                                if v_float > 0:
                                    banco_de_dados_geral.append({
                                        "Ano_Origem": ano,
                                        "Mes_Aba": aba.title,
                                        "Cliente": cliente,
                                        "Contrato": str(row[idx_contrato]).strip() if idx_contrato != -1 else "",
                                        "Unidade": str(row[idx_unidade]).strip() if idx_unidade != -1 else "",
                                        "Data_Pagamento": data_pagamento,
                                        "Valor_Recebido": v_float
                                    })
                                    
                    except Exception as e:
                        st.error(f"Aviso: Erro ao processar as abas do ano {ano}. Detalhe: {e}")

                # 4. EXIBIÇÃO DO RESULTADO
                if banco_de_dados_geral:
                    df_final = pd.DataFrame(banco_de_dados_geral)
                    
                    try:
                        df_final['Data_Pagamento_FMT'] = pd.to_datetime(df_final['Data_Pagamento'], errors='coerce', dayfirst=True)
                        df_final = df_final.sort_values(by='Data_Pagamento_FMT').drop(columns=['Data_Pagamento_FMT'])
                    except: pass

                    st.success(f"✅ Transbordo Concluído! {len(df_final)} liquidações exatas extraídas do histórico.")
                    st.dataframe(df_final, hide_index=True)
                    st.info(f"Abas vasculhadas com sucesso e mapeadas: {len(log_abas_lidas)}")
                else:
                    st.warning("A varredura foi concluída, mas nenhuma linha com 'DATA' preenchida com valor foi encontrada.")
                    st.write("Abas que o sistema conseguiu ler e identificou a tabela:", log_abas_lidas)
