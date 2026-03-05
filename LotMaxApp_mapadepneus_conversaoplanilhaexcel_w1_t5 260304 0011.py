import sys
import os
import io
import datetime
import pandas as pd
import xlsxwriter
import importlib
import streamlit as st

# Força o 'sys' no namespace global para bibliotecas que falham no importlib do 3.14
if 'sys' not in globals():
    globals()['sys'] = sys

titulo_app="App Lotmax - Mapa de Pneus - Mapeador de planilhas"
versao_app="w1.5"                 # Arquivo de origem - LotMaxApp_mapadepneus_conversaoplanilhaexcel_w1_t5 260304 0011.py

# --- 1. CONFIGURAÇÃO DA INTERFACE ---
st.set_page_config(page_title=titulo_app, layout="wide", initial_sidebar_state="expanded")
if 'idioma' not in st.session_state:
    st.session_state.idioma = 'pt-BR'


# --- 2. MATRIZ DE REGRAS UNIFORMIZADA (Integração Planilhas + App Lotmax) ---
MATRIZ_REGRAS = {
    "Placa ou Estoque":   {"critico": True,  "cod": "04-Local",             "tipo": "tamanho_texto", "limite": 7, "excecao": "ESTOQUE",      
                                                                            "warning": "⚠️ Você deve informar a placa do veículo instalado com até 7 caracteres ou palavra 'Estoque'."},
    "Marca":              {"critico": True,  "cod": "00-Marca",             "tipo": "tamanho_minimo", "limite": 4,                                                
                                                                            "warning": "⚠️ Você deve informar a marca com pelo menos 4 caracteres"},
    "Recapadora":         {"critico": False, "cod": "00-Recapadora",        "tipo": "nenhum",                                                
                                                                            "warning": ""},
    "Tipo":               {"critico": False, "cod": "00-Pneu_tipo",         "tipo": "lista",         "valores": ["liso", "borrachudo", "borrachudo florestal off -road pesado", "borrachudo off-road leve", "single", "misto", "liso-reboque", "comercial leve", "comercial médio", "passeio"],  
                                                                            "warning": ""},
    "Aplicacao":          {"critico": False, "cod": "01-Aplicacao",         "tipo": "lista",         "valores": ["pesado", "carreta", "leve ou medio", "passeio", "reboque"],  
                                                                            "warning": ""},
    "Código aplicado":    {"critico": True,  "cod": "01-Codigo_aplicado",   "tipo": "unico",                                            
                                                                            "warning": "❌ Não é permitido duplicações do código"},
    "Condicão":           {"critico": False, "cod": "01-Condicao",          "tipo": "lista",         "valores": ["novo", "novo - em uso", "recapado", "recapado - em uso"],  
                                                                            "warning": ""},
    "Medida":             {"critico": False, "cod": "02-Medida",            "tipo": "nenhum",                                        
                                                                            "warning": ""},
    "Vida util atual":    {"critico": True,  "cod": "01-Vida_util_atual",   "tipo": "numerico",                                      
                                                                            "warning": "⚠️ só é permitido valor"},
    "Recapes possíveis":  {"critico": True,  "cod": "01-Recapes",           "tipo": "lista",         "valores": ["0", "1", "2", "3"],     
                                                                            "warning": "⚠️ Lista valores: 0 a 3"},
    "Vida util recapes":  {"critico": True,  "cod": "01-Vida_util_recapado", "tipo": "numerico",                                      
                                                                            "warning": "⚠️ só é permitido valor"},
    "Código comercial":   {"critico": False, "cod": "02-Comercial",         "tipo": "nenhum",
                                                                            "warning": ""},
    "DOT fabricado":      {"critico": False, "cod": "02-DOT_Fabricado",     "tipo": "tamanho_fixo",   "limite": 4,                     
                                                                            "warning": "⚠️ Tamanho até 4 caracteres"},
    "Valor da compra":    {"critico": False, "cod": "02-Valor_compra",      "tipo": "numerico",                                       
                                                                            "warning": "⚠️ só é permitido valor"}
}

# Define a lista fixa globalmente para o botão de limpar funcionar
lista_fixa_base = list(MATRIZ_REGRAS.keys())

# --- 3. FUNÇÃO DE LEITURA (ESTÁVEL 3.12) ---
@st.cache_data(show_spinner="Lendo dados...", max_entries=10)
def ler_dados_excel(file, aba):
    try:
        engine_type = 'odf' if file.name.endswith('.ods') else 'openpyxl'
        df = pd.read_excel(file, sheet_name=aba, engine=engine_type)
        return df.copy()
    except Exception as e:
        st.error(f"Erro: {e}")
        return None

# --- 4. CSS ---
st.markdown("""
<style>
    /* 1. CONFIGURAÇÕES GERAIS DE PÁGINA */
    #MainMenu {visibility: hidden;} 
    footer {visibility: hidden;} 
    
    /* Esconde a barra superior mas mantém o botão de expandir lateral (>) funcional */
    header[data-testid="stHeader"] {
        background-color: rgba(0,0,0,0) !important;
        color: rgba(0,0,0,0) !important;
    }
    button[data-testid="stSidebarCollapseIcon"] {
        visibility: visible !important;
        color: #2c3e50 !important;
        z-index: 99999 !important;
    }

    .block-container { padding-top: 1.5rem !important; padding-bottom: 0rem !important; max-width: 98% !important; }

    /* 2. ESTILIZAÇÃO DOS SELECTBOX (ALTURA E FONTE) */
    div[data-baseweb="select"] > div { height: 28px !important; min-height: 28px !important; display: flex !important; align-items: center !important; }
    div[data-baseweb="select"] span { font-size: 0.8rem !important; line-height: 1 !important; }

    ul[role="listbox"] { padding: 0px !important; }
    ul[role="listbox"] li { padding: 0px !important; margin: 0px !important; min-height: 22px !important; display: flex !important; align-items: center !important; }

    /* 3. TRADUÇÃO E ESTILO DO UPLOADER (GESTÃO DE ARQUIVO) - REFORÇADO PARA DEPLOY */
    
    /* Moldura tracejada */
    [data-testid="stFileUploaderDropzone"] {
        padding: 12px !important;
        border: 1px dashed #d3d3d3 !important;
        border-radius: 8px !important;
        background-color: #f9f9f9 !important;
        margin-top: 5px !important;
    }

    /* Esconde textos nativos (Drag and Drop / Limit) */
    [data-testid="stFileUploaderDropzoneInstructions"] div span, 
    [data-testid="stFileUploaderDropzoneInstructions"] div small { 
        display: none !important; 
    }
    
    /* Insere o texto em Português no lugar do "Drag and drop" */
    [data-testid="stFileUploaderDropzoneInstructions"] div::before {
        content: "Arraste e solte o arquivo aqui";
        display: block !important;
        font-size: 13px !important;
        color: #555 !important;
        visibility: visible !important;
        margin-bottom: 5px !important;
    }

    /* Customizar o botão "Browse files" */
    [data-testid="stFileUploaderDropzone"] button { 
        color: transparent !important; 
        position: relative;
        width: 100% !important;
        border: 1px solid #d3d3d3 !important;
        background-color: white !important;
    }
    
    [data-testid="stFileUploaderDropzone"] button::after {
        content: "📁 Selecionar arquivo";
        visibility: visible;
        color: #2c3e50;
        font-weight: 600;
        position: absolute;
        width: 100%;
        display: flex;
        align-items: center;
        justify-content: center;
        left: 0;
        font-size: 0.75rem;
    }

    /* 4. LABELS, ERROS E ALERTAS */
    .mapping-label { font-weight: 700; color: #2c3e50; margin-bottom: 1px; font-size: 0.82rem; display: block; }
    div[data-testid="stSelectbox"] { margin-bottom: -10px !important; }
    .val-error { color: #d63031; font-size: 0.65rem; font-weight: 700; margin-top: 2px; line-height: 1.1; }
    .val-warning { color: #f39c12; font-size: 0.65rem; font-weight: 700; margin-top: 2px; line-height: 1.1; }
</style>
""", unsafe_allow_html=True)




# --- 5. CABEÇALHO ---
c_logo, c_titulo = st.columns([1, 4])
with c_logo:
    logo_nome = "Lotmax_app_lotmax_2026.png"
    if os.path.exists(logo_nome): 
        st.image(logo_nome, width=110)
    else: 
        st.markdown("### 🚀 App LotMax")

with c_titulo:
    # Título principal + Versão + Debug oculto (selecionável com mouse)
    st.markdown(f"""
        <h3 style='margin-top: 15px; margin-bottom: 0px;'>
            {titulo_app} 
            <span style='font-size: 0.85rem; font-weight: 400; color: #7f8c8d; margin-left: 8px;'> 
                {versao_app}
            </span>
            <span style='color: transparent; font-size: 0.5rem; user-select: text;'>
                 {sys.version}.{sys.executable}
            </span>
        </h3>
    """, unsafe_allow_html=True)



st.divider()

# --- 6. BARRA LATERAL (UPLOAD APENAS) ---
with st.sidebar:
    st.markdown("### 📂 Gestão de Arquivo")
    uploaded_file = st.file_uploader("Upload Excel/ODS", type=["xlsx", "xls", "ods"], label_visibility="collapsed")

# --- 7. LÓGICA CENTRAL ---
if uploaded_file:
    # DETECÇÃO AUTOMÁTICA DE TROCA DE ARQUIVO
    if st.session_state.get('ultimo_arquivo_nome') != uploaded_file.name:
        # 1. Limpa o dicionário de mapeamento
        st.session_state.map_state = {item: "(Pular)" for item in lista_fixa_base}
        # 2. Salva o nome do arquivo atual para a próxima comparação
        st.session_state.ultimo_arquivo_nome = uploaded_file.name
        # 3. MUITO IMPORTANTE: Incrementa o contador para forçar o reset dos widgets
        st.session_state.reset_ctr = st.session_state.get('reset_ctr', 0) + 1
        # 4. Rerun para aplicar a nova 'key' em todos os selectboxes abaixo
        st.rerun()

    # BLOCO VISUAL (Seu código original com botão)
    col_info, col_reset = st.columns([3, 1])
    with col_info:
        st.markdown(f"📄 **Arquivo:** `{uploaded_file.name}`")
    with col_reset:
        if st.button("🗑️ Limpar Seleções", use_container_width=True):
            st.session_state.map_state = {item: "(Pular)" for item in lista_fixa_base}
            st.session_state.reset_ctr = st.session_state.get('reset_ctr', 0) + 1
            st.rerun()


    r_key = st.session_state.get('reset_ctr', 0)
    xls = pd.ExcelFile(uploaded_file)
    aba_sel = st.selectbox("Selecione a Aba:", xls.sheet_names, key=f"aba_main_{r_key}")

    if aba_sel:
        df_origem = ler_dados_excel(uploaded_file, aba_sel)
        if df_origem is not None:
            colunas_planilha = df_origem.columns.tolist()

            if 'map_state' not in st.session_state:
                st.session_state.map_state = {item: "(Pular)" for item in lista_fixa_base}

            selecionados_atualmente = {v for k, v in st.session_state.map_state.items() if v != "(Pular)"}
            campos_com_erro_critico = []

            def format_rows(mask):
                lista_linhas = mask[mask].index.map(lambda x: int(x) + 2).tolist()
                t = len(lista_linhas)
                return f"{lista_linhas[:3]}... (+{t - 3})" if t > 3 else str(lista_linhas)

            grid = st.columns(4)
            for idx, item_fixo in enumerate(lista_fixa_base):
                with grid[idx % 4]:
                    st.markdown(f"<span class='mapping-label'>{item_fixo}</span>", unsafe_allow_html=True)
                    valor_salvo = st.session_state.map_state.get(item_fixo, "(Pular)")
                    opcoes_disponiveis = ["(Pular)"] + [c for c in colunas_planilha if c not in selecionados_atualmente or c == valor_salvo]
                    
                    idx_p = opcoes_disponiveis.index(valor_salvo) if valor_salvo in opcoes_disponiveis else 0
                    nova_escolha = st.selectbox(f"sel_{item_fixo}", options=opcoes_disponiveis, index=idx_p, key=f"f_{item_fixo}_{r_key}", label_visibility="collapsed")
                    
                    if nova_escolha != valor_salvo:
                        st.session_state.map_state[item_fixo] = nova_escolha
                        st.rerun()

                    # --- MOTOR DE VALIDAÇÃO ---
                    regra = MATRIZ_REGRAS[item_fixo]
                    if nova_escolha != "(Pular)" and regra["tipo"] != "nenhum":
                        col_data = df_origem[nova_escolha]
                        dados_limpos = col_data.dropna()
                        mask = None
                        msg_aviso = regra.get("warning", "")

                        if regra["tipo"] == "lista":
                            mask = ~dados_limpos.astype(str).str.lower().str.strip().isin(regra["valores"])
                            msg_aviso = f"⚠️ Use: {', '.join(regra['valores'])}"
                        elif regra["tipo"] == "tamanho_texto":
                            mask = dados_limpos.apply(lambda x: len(str(x)) > regra["limite"] and str(x).strip().upper() != regra["excecao"])
                        elif regra["tipo"] == "tamanho_minimo":
                            mask = dados_limpos.apply(lambda x: len(str(x)) < regra["limite"])
                        elif regra["tipo"] == "unico":
                            mask = col_data.duplicated(keep=False) & col_data.notna()
                        elif regra["tipo"] == "numerico":
                            mask = pd.to_numeric(dados_limpos, errors='coerce').isna()
                        elif regra["tipo"] == "tamanho_fixo":
                            mask = dados_limpos.apply(lambda x: len(str(x).strip()) != regra["limite"])

                        if mask is not None and mask.any():
                            if regra.get("critico"): campos_com_erro_critico.append(item_fixo)
                            classe = "val-error" if regra.get("critico") else "val-warning"
                            invalidos = f"<br>❌ Digitado: {dados_limpos[mask].unique().tolist()[:2]}" if regra["tipo"] == "lista" else ""
                            st.markdown(f"<p class='{classe}'>Linhas: {format_rows(mask)}{invalidos}<br>{msg_aviso}</p>", unsafe_allow_html=True)

            # --- 8. EXPORTAÇÃO DIFERENCIADA (EXCEL REVISÃO vs CSV pra LotMax) ---
            mapeamento_final = {v: k for k, v in st.session_state.map_state.items() if v != "(Pular)"}
            
            if mapeamento_final:
                st.divider()
                if len(campos_com_erro_critico) > 0:
                    st.error(f"⚠️ **Download Bloqueado.** Corrija erros críticos em: {', '.join(set(campos_com_erro_critico))}")
                else:
                    if st.button("🚀 PROCESSAR ARQUIVOS DE CARGA"):
                        with st.spinner("Gerando arquivos..."):
                            # A. CONSTRUÇÃO DO EXCEL (REVISÃO AMIGÁVEL - NOMES ORIGINAIS)
                            df_excel = pd.DataFrame(index=df_origem.index)
                            for item in lista_fixa_base:
                                col_user = st.session_state.map_state.get(item)
                                df_excel[item] = df_origem[col_user] if col_user and col_user != "(Pular)" else ""

                            # B. CONSTRUÇÃO DO CSV (CARGA LotMax - CODIFICADO + EXTRAS)
                            df_csv = df_excel.copy()

                            # 1. Renomeia as colunas amigáveis (com acento) para os códigos técnicos (sem acento)
                            nomes_tecnicos = {item: MATRIZ_REGRAS[item]["cod"] for item in lista_fixa_base}
                            df_csv = df_csv.rename(columns=nomes_tecnicos)

                            # --- CONCATENAÇÃO "TUDO JUNTO" (SEM PIPE) ---
                            c_marca = MATRIZ_REGRAS["Marca"]["cod"]
                            c_tipo  = MATRIZ_REGRAS["Tipo"]["cod"]
                            c_cod   = MATRIZ_REGRAS["Código aplicado"]["cod"]

                            # Removemos os pipes e espaços. Ex: pn-Pireliso12345
                            df_csv["01-Pneu_string"] = "pn-" + \
                            df_csv[c_marca].astype(str).str[:4] + \
                            df_csv[c_tipo].astype(str).str[:4] + \
                            df_csv[c_cod].astype(str)

                            # 3. Cálculo de Vida Útil Total
                            c_atual = MATRIZ_REGRAS["Vida util atual"]["cod"]
                            c_rec_q = MATRIZ_REGRAS["Recapes possíveis"]["cod"]
                            c_rec_v = MATRIZ_REGRAS["Vida util recapes"]["cod"]

                            v_at = pd.to_numeric(df_csv[c_atual], errors='coerce').fillna(0)
                            v_rq = pd.to_numeric(df_csv[c_rec_q], errors='coerce').fillna(0)
                            v_rv = pd.to_numeric(df_csv[c_rec_v], errors='coerce').fillna(0)

                            df_csv["03-Vida_novo_recapes"] = (v_at + (v_rq * v_rv)).astype(int)


                            # C. DOWNLOADS COM TIMESTAMP
                            agora = datetime.datetime.now().strftime("%y%m%d_%H%M")
                            nome_puro = os.path.splitext(uploaded_file.name)[0]
                            
                            # Preparar Excel
                            out_xlsx = io.BytesIO()
                            with pd.ExcelWriter(out_xlsx, engine='xlsxwriter') as writer:
                                df_excel.to_excel(writer, index=False)
                            
                            # Preparar CSV (UTF-8-SIG e Delimitador ;) conforme padrão LotMax
                            csv_data = df_csv.to_csv(index=False, sep=';', encoding='utf-8-sig')

                            st.success("✅ Arquivos processados com sucesso!")
                            col_a, col_b = st.columns(2)
                            with col_a:
                                st.download_button("📥 CSV LotMax (Técnico)", csv_data, f"{nome_puro}_App_LotMax_{agora}.csv", "text/csv", use_container_width=True)
                            with col_b:
                                st.download_button("📄 EXCEL para revisão (Nomes)", out_xlsx.getvalue(), f"{nome_puro}_revisão_{agora}.xlsx", "application/vnd.ms-excel", use_container_width=True)

else:
    st.info("Aguardando upload do arquivo Excel ou ODS...")
