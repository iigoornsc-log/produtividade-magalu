import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
import json
from datetime import datetime

# =========================================================================
# 1. CONFIGURAÇÕES INICIAIS E CSS
# =========================================================================
st.set_page_config(page_title="Torre de Controle | Logística", page_icon="⚡️", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #F8F9FA; }
    .kpi-card {
        background-color: #FFFFFF; border-radius: 12px; padding: 20px;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.04); border-left: 5px solid #0086FF;
        transition: transform 0.2s; margin-bottom: 20px;
    }
    .kpi-card:hover { transform: translateY(-3px); box-shadow: 0 6px 15px rgba(0, 0, 0, 0.08); }
    .kpi-title { margin: 0; font-size: 12px; color: #6C757D; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-value { margin: 5px 0; font-size: 32px; color: #212529; font-weight: 900; }
    .kpi-subtitle { margin: 0; font-size: 12px; color: #ADB5BD; font-weight: 500; }
    .bloco-header { color: #2C3E50; font-weight: 800; font-size: 22px; margin-top: 30px; margin-bottom: 10px; border-bottom: 2px solid #E9ECEF; padding-bottom: 5px;}
</style>
""", unsafe_allow_html=True)

def exibir_kpi(titulo, valor, subtitulo="", cor="#0086FF"):
    st.markdown(f"""
    <div class="kpi-card" style="border-left-color: {cor};">
        <p class="kpi-title">{titulo}</p>
        <p class="kpi-value">{valor}</p>
        <p class="kpi-subtitle">{subtitulo}</p>
    </div>
    """, unsafe_allow_html=True)

# --- A VACINA DOS NÚMEROS BRASILEIROS ---
def limpa_numero_br(valor):
    if pd.isna(valor) or str(valor).strip() in ['', 'NAN', 'NULL', 'NONE']: return 0
    v = str(valor).strip()
    if ',' in v:
        v = v.replace('.', '').replace(',', '.')
    else:
        v = v.replace('.', '')
    try:
        return float(v)
    except:
        return 0

# --- CONVERSÕES DE TEMPO ---
def time_to_mins(t_str):
    if pd.isna(t_str) or str(t_str).strip() == '': return 0
    try:
        partes = str(t_str).split(':')
        if len(partes) == 3: return int(partes[0]) * 60 + int(partes[1]) + float(partes[2]) / 60.0
        return 0
    except: return 0

def mins_to_text(mins):
    if pd.isna(mins) or mins <= 0: return "0m"
    total_m = int(round(mins))
    if total_m == 0: return "< 1m" 
    h = total_m // 60
    m = total_m % 60
    if h > 0: return f"{h}h {m}m"
    return f"{m}m"

# --- GRAVAR NO COFRE ---
def salvar_historico_fechamento(df_para_salvar):
    try:
        cred_dict = json.loads(st.secrets["google_json"])
        creds = Credentials.from_service_account_info(cred_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        client = gspread.authorize(creds)
        sh2 = client.open_by_key('1bj5vIu8LOIWqaW5evogwQeyrJd9yj1iQkXHbJKvTeks')
        aba = sh2.worksheet("FECHAMENTO")
        aba.append_rows(df_para_salvar.values.tolist())
        return True
    except Exception as e:
        st.error(f"Erro ao salvar na planilha: {e}")
        return False

# =========================================================================
# 2. MOTOR DE DADOS MULTI-PLANILHAS
# =========================================================================
@st.cache_data(ttl=300)
def carregar_dados_armazenagem():
    try:
        cred_dict = json.loads(st.secrets["google_json"])
        creds = Credentials.from_service_account_info(cred_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        client = gspread.authorize(creds)
        
        sh = client.open_by_key('1F4Qs5xGPMjgWSO6giHSwFfDf5F-mlv1RuT4riEVU0I0')
        ws = sh.worksheet("ACOMPANHAMENTO GERAL")
        data = ws.get_all_values()
        
        df = pd.DataFrame(data[1:], columns=data[0])
        df.columns = df.columns.str.strip().str.upper() 
        
        df['NU_ETIQUETA'] = df['NU_ETIQUETA'].astype(str).str.strip()
        df['AGENDA'] = df['AGENDA'].astype(str).str.strip().str.upper()
        df['PRODUTO'] = df['PRODUTO'].astype(str).str.strip()
        df['QT_PRODUTO'] = df['QT_PRODUTO'].apply(limpa_numero_br)
        df['SITUACAO'] = df['SITUACAO'].astype(str).str.strip()
        df['OPERADOR'] = df['OPERADOR'].astype(str).str.strip().str.upper()
        df['CONFERENTE'] = df['CONFERENTE'].astype(str).str.strip().str.upper()
        
        df['Data_Ref'] = pd.to_datetime(df['DATA'], format='%d/%m/%Y', errors='coerce').dt.date
        df['DT_CONFERENCIA_CALC'] = pd.to_datetime(df['DT_CONFERENCIA'], errors='coerce') 
        df['DT_ARMAZENAGEM_CALC'] = pd.to_datetime(df['DT_ARMAZENAGEM'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
        
        df['Data_Conf'] = df['DT_CONFERENCIA_CALC'].dt.date
        df['Data_Armz'] = df['Data_Ref'] 
        
        def formata_hora(h):
            if pd.isna(h) or str(h).strip() in ['', 'NAN', 'NULL', 'NONE']: return None
            try: return f"{int(float(h)):02d}:00"
            except: return None
                
        df['Hora_Conf'] = df['HORA CONF'].apply(formata_hora)
        df['Hora_Armz'] = df['HORA ARMZ'].apply(formata_hora)
        df['Tempo_Espera_Minutos'] = (df['DT_ARMAZENAGEM_CALC'] - df['DT_CONFERENCIA_CALC']).dt.total_seconds() / 60.0
        
        return df.dropna(subset=['Data_Ref'])
    except Exception as e:
        st.error(f"Erro Armazenagem: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=300)
def carregar_dados_conferencia():
    try:
        cred_dict = json.loads(st.secrets["google_json"])
        creds = Credentials.from_service_account_info(cred_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        client = gspread.authorize(creds)
        
        sh2 = client.open_by_key('1bj5vIu8LOIWqaW5evogwQeyrJd9yj1iQkXHbJKvTeks')
        todas_abas = sh2.worksheets()
        
        # --- 1. HISTÓRICO (BASE DE DADOS) ---
        aba_hist = next((aba for aba in todas_abas if "BASE DE DADOS" in aba.title.upper()), None)
        if aba_hist:
            data_hist = aba_hist.get("Q:W")
            df_hist = pd.DataFrame(data_hist[1:], columns=data_hist[0])
            df_hist.columns = df_hist.columns.str.strip().str.upper()
            if 'TMP APC' in df_hist.columns: df_hist['TMP APC'] = df_hist['TMP APC'].apply(limpa_numero_br)
            if 'PEÇAS' in df_hist.columns: df_hist['PEÇAS'] = df_hist['PEÇAS'].apply(limpa_numero_br)
            if 'SKU' in df_hist.columns: df_hist['SKU'] = df_hist['SKU'].apply(limpa_numero_br)
        else: df_hist = pd.DataFrame()
            
        # --- 2. PLANILHA DIA ATUAL ---
        aba_hoje = next((aba for aba in todas_abas if "DIA ATUAL" in aba.title.upper()), None)
        if aba_hoje:
            data_hoje = aba_hoje.get("A:I")
            df_hoje = pd.DataFrame(data_hoje[1:], columns=data_hoje[0])
            df_hoje.columns = df_hoje.columns.str.strip().str.upper()
            if 'STATUS' in df_hoje.columns: df_hoje.rename(columns={'STATUS': 'STATUS_FISICO'}, inplace=True)
            if 'PEÇAS' in df_hoje.columns: df_hoje['PEÇAS'] = df_hoje['PEÇAS'].apply(limpa_numero_br)
            if 'PÇS PENDENTES' in df_hoje.columns: df_hoje['PÇS PENDENTES'] = df_hoje['PÇS PENDENTES'].apply(limpa_numero_br)
            if 'SKU' in df_hoje.columns: df_hoje['SKU'] = df_hoje['SKU'].apply(limpa_numero_br)
            if 'DURAÇÃO CARGA' in df_hoje.columns: df_hoje['DURAÇÃO CARGA'] = df_hoje['DURAÇÃO CARGA'].astype(str).str.strip()
        else: df_hoje = pd.DataFrame()

        # --- 3. COFRE DE FECHAMENTO ---
        aba_fechamento = next((aba for aba in todas_abas if "FECHAMENTO" in aba.title.upper()), None)
        if aba_fechamento:
            data_fech = aba_fechamento.get_all_values()
            if len(data_fech) > 1:
                df_fechamento = pd.DataFrame(data_fech[1:], columns=data_fech[0])
                df_fechamento.columns = df_fechamento.columns.str.strip().str.upper()
                
                # A MÁGICA: Limpando a vírgula do Google Sheets antes de fazer conta!
                if 'META MINUTOS' in df_fechamento.columns:
                    df_fechamento['META MINUTOS'] = df_fechamento['META MINUTOS'].apply(limpa_numero_br)
                if 'REALIZADO MINUTOS' in df_fechamento.columns:
                    df_fechamento['REALIZADO MINUTOS'] = df_fechamento['REALIZADO MINUTOS'].apply(limpa_numero_br)
            else: df_fechamento = pd.DataFrame()
        else: df_fechamento = pd.DataFrame()
            
        return df_hist, df_hoje, df_fechamento
    except Exception as e:
        st.error(f"Erro Conferência: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# =========================================================================
# FUNÇÃO DO POP-UP (RAIO-X DA HORA)
# =========================================================================
@st.dialog("🔍 RAIO-X DA HORA: DETALHAMENTO", width="large")
def popup_detalhe_hora(hora, df_base, data_sel):
    df_conferido = df_base[(df_base['Hora_Conf'] == hora) & (df_base['Data_Conf'] == data_sel)].copy()
    df_armazenado = df_base[(df_base['Hora_Armz'] == hora) & (df_base['Data_Armz'] == data_sel) & (df_base['SITUACAO'] == '25')].copy()
    df_hora = pd.concat([df_conferido, df_armazenado]).drop_duplicates(subset=['NU_ETIQUETA'])
    
    if df_hora.empty:
        st.warning(f"Nenhuma movimentação às {hora}.")
        return
        
    st.markdown(f"### ⏱️ Resumo das **{hora}**")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Conferidas Aqui", df_conferido['NU_ETIQUETA'].nunique())
    c2.metric("Armazenadas Aqui", df_armazenado['NU_ETIQUETA'].nunique())
    c3.metric("Peças Movimentadas", f"{df_hora['QT_PRODUTO'].sum():,.0f}".replace(',','.'))
    c4.metric("Agendas Envolvidas", df_hora['AGENDA'].nunique())
    
    df_exibicao = df_hora[['NU_ETIQUETA', 'SITUACAO', 'PRODUTO', 'CONFERENTE', 'OPERADOR', 'Hora_Conf', 'Hora_Armz']].copy()
    df_exibicao['SITUACAO'] = df_exibicao['SITUACAO'].map({'23': '23 - Pendente', '25': '25 - Armazenado'})
    st.dataframe(df_exibicao, use_container_width=True, hide_index=True)

# =========================================================================
# 3. INTERFACE E ABAS
# =========================================================================
df_armz = carregar_dados_armazenagem()
df_hist_conf, df_hoje_conf, df_fechamento = carregar_dados_conferencia()

tab1, tab2, tab3 = st.tabs(["📦 Torre de Armazenagem (Doca)", "🔎 Torre de Conferência (Metas)", "🏅 Desempenho da Equipe"])

# -------------------------------------------------------------------------
# ABA 1: ARMAZENAGEM
# -------------------------------------------------------------------------
with tab1:
    if not df_armz.empty:
        df_armz_filtrado = df_armz[df_armz['SITUACAO'].isin(['23', '25'])]
        
        st.sidebar.image("https://magalog.com.br/opengraph-image.jpg?fdd536e7d35ec9da", width=250)
        st.sidebar.markdown("### 🎛️ Controles da Operação")
        
        data_max = df_armz_filtrado['Data_Conf'].dropna().max()
        data_sel = st.sidebar.date_input("🗓️ Data de Análise (Armaz.)", data_max)
        
        modo_visao = st.sidebar.radio("🔎 Modo de Análise", ["Líquida (Apenas do Dia)", "Global (Incluir Herança)"])
        
        df_hoje_c = df_armz_filtrado[df_armz_filtrado['Data_Conf'] == data_sel]
        df_hoje_a = df_armz_filtrado[df_armz_filtrado['Data_Armz'] == data_sel]
        
        if modo_visao == "Líquida (Apenas do Dia)":
            df_base_armz = df_hoje_c.copy()
        else:
            df_base_armz = pd.concat([df_hoje_c, df_hoje_a]).drop_duplicates(subset=['NU_ETIQUETA'])

        fantasmas = ['', 'NAN', 'NONE', 'NULL']
        conferentes_validos = sorted([c for c in df_base_armz['CONFERENTE'].unique() if pd.notna(c) and c not in fantasmas])
        conf_sel = st.sidebar.multiselect("📋 Filtrar Conferente:", options=conferentes_validos, default=conferentes_validos)
        df_base_armz = df_base_armz[df_base_armz['CONFERENTE'].isin(conf_sel)]

        operadores_validos = sorted([op for op in df_base_armz['OPERADOR'].unique() if pd.notna(op) and op not in fantasmas])
        op_sel = st.sidebar.multiselect("👥 Filtrar Operador:", options=operadores_validos, default=operadores_validos)
        
        df_producao_equipe = df_base_armz[(df_base_armz['Data_Armz'] == data_sel) & (df_base_armz['SITUACAO'] == '25') & (df_base_armz['OPERADOR'].isin(op_sel))]

        st.title(f"🚀 Torre de Controle | {data_sel.strftime('%d/%m/%Y')}")
        
        c1, c2, c3, c4 = st.columns(4)
        qtd_etiquetas_armz = df_producao_equipe['NU_ETIQUETA'].nunique()
        qtd_pendentes_doca = df_base_armz[df_base_armz['SITUACAO'] == '23']['NU_ETIQUETA'].nunique()
        if modo_visao == "Global (Incluir Herança)":
            qtd_pendentes_doca += df_armz_filtrado[(df_armz_filtrado['Data_Conf'] < data_sel) & (df_armz_filtrado['SITUACAO'] == '23') & (df_armz_filtrado['CONFERENTE'].isin(conf_sel))]['NU_ETIQUETA'].nunique()
            
        espera_valida = df_producao_equipe[df_producao_equipe['Tempo_Espera_Minutos'] > 0]['Tempo_Espera_Minutos']
        sla_medio = espera_valida.mean() if not espera_valida.empty else 0
        
        with c1: exibir_kpi("Armazenados (Sit. 25)", f"{qtd_etiquetas_armz:,.0f}", "Equipe selecionada", "#0086FF")
        with c2: exibir_kpi("Pendências (Sit. 23)", f"{qtd_pendentes_doca:,.0f}", "Fila total da Doca", "#E74C3C")
        with c3: exibir_kpi("SLA Médio Doca", mins_to_text(sla_medio), "Tempo em espera", "#F44336" if sla_medio > 120 else "#4CAF50")
        with c4: exibir_kpi("Operadores Ativos", str(len(op_sel)), "Filtro aplicado", "#FF9800")

        col_tit, col_sel = st.columns([7, 3])
        with col_tit: st.markdown("<div class='bloco-header'>🌊 Fluxo de Trabalho e Fila</div>", unsafe_allow_html=True)
        
        horas_conf = df_base_armz[df_base_armz['Data_Conf'] == data_sel]['Hora_Conf'].dropna().unique()
        horas_armz = df_producao_equipe['Hora_Armz'].dropna().unique()
        todas_horas = sorted(list(set(list(horas_conf) + list(horas_armz))))
        
        with col_sel:
            st.markdown("<br>", unsafe_allow_html=True)
            hora_manual = st.selectbox("🖱️ Abrir Raio-X da Hora:", ["Selecione..."] + todas_horas)

        dados_grafico = []
        for hora in todas_horas:
            conf_hora = df_base_armz[(df_base_armz['Data_Conf'] == data_sel) & (df_base_armz['Hora_Conf'] == hora)]['NU_ETIQUETA'].nunique()
            armz_hora = df_producao_equipe[df_producao_equipe['Hora_Armz'] == hora]['NU_ETIQUETA'].nunique()
            
            if modo_visao == "Líquida (Apenas do Dia)":
                entrou = df_base_armz[(df_base_armz['Data_Conf'] == data_sel) & (df_base_armz['Hora_Conf'] <= hora)]['NU_ETIQUETA'].nunique()
                saiu = df_base_armz[(df_base_armz['Data_Conf'] == data_sel) & (df_base_armz['SITUACAO'] == '25') & (df_base_armz['Data_Armz'] == data_sel) & (df_base_armz['Hora_Armz'] <= hora)]['NU_ETIQUETA'].nunique()
                pendencias = entrou - saiu
            else:
                entrou = df_armz_filtrado[(df_armz_filtrado['CONFERENTE'].isin(conf_sel)) & ((df_armz_filtrado['Data_Conf'] < data_sel) | ((df_armz_filtrado['Data_Conf'] == data_sel) & (df_armz_filtrado['Hora_Conf'] <= hora)))]['NU_ETIQUETA'].nunique()
                saiu = df_armz_filtrado[(df_armz_filtrado['CONFERENTE'].isin(conf_sel)) & (df_armz_filtrado['SITUACAO'] == '25') & ((df_armz_filtrado['Data_Armz'] < data_sel) | ((df_armz_filtrado['Data_Armz'] == data_sel) & (df_armz_filtrado['Hora_Armz'] <= hora)))]['NU_ETIQUETA'].nunique()
                pendencias = entrou - saiu
                
            dados_grafico.append({'Hora': hora, 'Armazenados': armz_hora, 'Conferidos': conf_hora, 'Pendências': max(0, pendencias)})
            
        df_fluxo = pd.DataFrame(dados_grafico)

        if not df_fluxo.empty:
            fig_fluxo = go.Figure()
            fig_fluxo.add_trace(go.Bar(x=df_fluxo['Hora'], y=df_fluxo['Armazenados'], name='Armazenados', marker_color='#0086FF', text=df_fluxo['Armazenados'], textposition='auto'))
            fig_fluxo.add_trace(go.Bar(x=df_fluxo['Hora'], y=df_fluxo['Conferidos'], name='Conferidos', marker_color='#9d26ff', text=df_fluxo['Conferidos'], textposition='outside'))
            fig_fluxo.add_trace(go.Scatter(x=df_fluxo['Hora'], y=df_fluxo['Pendências'], name='Pendências', mode='lines+markers+text', line=dict(color='#E74C3C', width=3), yaxis='y2', text=df_fluxo['Pendências'], textposition='top center'))
            
            fig_fluxo.update_layout(plot_bgcolor='rgba(0,0,0,0)', barmode='group', legend=dict(orientation="h", y=1.15, x=0.5, xanchor='center'), yaxis2=dict(overlaying='y', side='right', showgrid=False), hovermode="x unified")
            
            ev = st.plotly_chart(fig_fluxo, use_container_width=True, on_select="rerun")
            if hora_manual != "Selecione...": popup_detalhe_hora(hora_manual, df_base_armz, data_sel)
            elif isinstance(ev, dict) and "selection" in ev and ev["selection"].get("points"):
                popup_detalhe_hora(ev["selection"]["points"][0].get("x"), df_base_armz, data_sel)
    else: st.warning("Sem dados de Armazenagem.")

# -------------------------------------------------------------------------
# ABA 2: CONFERÊNCIA (METAS PREDITIVAS E AUTO-SAVE)
# -------------------------------------------------------------------------
with tab2:
    if not df_hist_conf.empty and not df_hoje_conf.empty:
        st.title("🎯 Torre de Conferência e Previsão")
        st.caption("Cálculo preditivo inteligente: Busca no histórico as cargas com o mesmo perfil.")
        
        taxa_global_cd = df_hist_conf['TMP APC'].sum() / df_hist_conf['PEÇAS'].sum() if df_hist_conf['PEÇAS'].sum() > 0 else 1.0

        def calcular_meta_inteligente(row, df_historico):
            forn = str(row.get('ORIGEM', '')).strip().upper()
            linha = str(row.get('CATEGORIA', '')).strip().upper()
            pecas = row.get('PEÇAS', 0)
            sku = row.get('SKU', 0)
            
            min_pecas, max_pecas = pecas * 0.7, pecas * 1.3
            min_sku, max_sku = min(sku * 0.7, sku - 2), max(sku * 1.3, sku + 2)

            df_base_exata = df_historico[(df_historico['FORNECEDOR'].str.upper() == forn) & (df_historico['LINHA'].str.upper() == linha)]
            if not df_base_exata.empty:
                df_gemeas = df_base_exata[(df_base_exata['PEÇAS'] >= min_pecas) & (df_base_exata['PEÇAS'] <= max_pecas) & (df_base_exata['SKU'] >= min_sku) & (df_base_exata['SKU'] <= max_sku)]
                if not df_gemeas.empty: return df_gemeas['TMP APC'].mean()
                df_primas = df_base_exata[(df_base_exata['PEÇAS'] >= min_pecas) & (df_base_exata['PEÇAS'] <= max_pecas)]
                if not df_primas.empty: return df_primas['TMP APC'].mean()
                if df_base_exata['PEÇAS'].sum() > 0: return pecas * (df_base_exata['TMP APC'].sum() / df_base_exata['PEÇAS'].sum())

            df_base_categoria = df_historico[df_historico['LINHA'].str.upper() == linha]
            if not df_base_categoria.empty:
                df_gemeas_cat = df_base_categoria[(df_base_categoria['PEÇAS'] >= min_pecas) & (df_base_categoria['PEÇAS'] <= max_pecas) & (df_base_categoria['SKU'] >= min_sku) & (df_base_categoria['SKU'] <= max_sku)]
                if not df_gemeas_cat.empty: return df_gemeas_cat['TMP APC'].mean()
                df_primas_cat = df_base_categoria[(df_base_categoria['PEÇAS'] >= min_pecas) & (df_base_categoria['PEÇAS'] <= max_pecas)]
                if not df_primas_cat.empty: return df_primas_cat['TMP APC'].mean()
                if df_base_categoria['PEÇAS'].sum() > 0: return pecas * (df_base_categoria['TMP APC'].sum() / df_base_categoria['PEÇAS'].sum())

            return pecas * taxa_global_cd

        df_hoje_conf['DURAÇÃO_REAL_MIN'] = df_hoje_conf['DURAÇÃO CARGA'].apply(time_to_mins)
        df_hoje_conf['STATUS_FISICO'] = df_hoje_conf['STATUS_FISICO'].str.strip().str.upper()
        df_hoje_conf['META_TEMPO_MIN'] = df_hoje_conf.apply(lambda row: calcular_meta_inteligente(row, df_hist_conf), axis=1)
        
        agora = pd.Timestamp.now(tz='America/Sao_Paulo')
        def calcular_previsao(row):
            status = row['STATUS_FISICO']
            if status == 'OK': return "✅ Finalizado"
            restante = row['META_TEMPO_MIN'] - row['DURAÇÃO_REAL_MIN']
            if restante < 0: return "⚠️ Já Estourou"
            return (agora + pd.Timedelta(minutes=restante)).strftime("%H:%M")
            
        def calcular_situacao_meta(row):
            status = row['STATUS_FISICO']
            if status == 'OK': return "✅ No Prazo" if row['DURAÇÃO_REAL_MIN'] <= row['META_TEMPO_MIN'] else "🔴 Atrasou (Finalizado)"
            else:
                if row['DURAÇÃO_REAL_MIN'] > row['META_TEMPO_MIN']: return "🔴 Atrasado (Em Processo)"
                elif status == 'EM PROCESSO': return "⏳ No Ritmo"
                else: return "⏸️ Aguardando Início"
            
        df_hoje_conf['PREVISÃO FIM'] = df_hoje_conf.apply(calcular_previsao, axis=1)
        df_hoje_conf['SITUAÇÃO META'] = df_hoje_conf.apply(calcular_situacao_meta, axis=1)
        
        c1, c2, c3, c4 = st.columns(4)
        cargas_totais = len(df_hoje_conf)
        cargas_ok = df_hoje_conf[df_hoje_conf['STATUS_FISICO'] == 'OK'].shape[0]
        cargas_fila = df_hoje_conf[df_hoje_conf['STATUS_FISICO'].isin(['EM DOCA', 'P-EXTERNO'])].shape[0]
        acertos = df_hoje_conf[df_hoje_conf['SITUAÇÃO META'].isin(['✅ No Prazo', '⏳ No Ritmo', '⏸️ Aguardando Início'])].shape[0]
        perc_acerto = (acertos / cargas_totais) * 100 if cargas_totais > 0 else 0
        
        with c1: exibir_kpi("Total Agendas", cargas_totais, "Na grade de hoje", "#9B59B6")
        with c2: exibir_kpi("Cargas Finalizadas", cargas_ok, "Status 'OK'", "#0086FF")
        with c3: exibir_kpi("Fila Física", cargas_fila, "Doca ou P-Externo", "#FF9800")
        with c4: exibir_kpi("Saúde das Metas", f"{perc_acerto:.1f}%", "Aderência ao tempo", "#4CAF50" if perc_acerto > 80 else "#F44336")
        
        st.markdown("<div class='bloco-header'>📊 Despacho de Cargas e Previsão de Fim</div>", unsafe_allow_html=True)
        
        df_tabela = df_hoje_conf[['AGENDA', 'CONFERENTE', 'CATEGORIA', 'STATUS_FISICO', 'PEÇAS', 'SKU', 'META_TEMPO_MIN', 'DURAÇÃO_REAL_MIN', 'PREVISÃO FIM', 'SITUAÇÃO META']].copy()
        df_tabela['META (Tempo)'] = df_tabela['META_TEMPO_MIN'].apply(mins_to_text)
        df_tabela['GASTO (Tempo)'] = df_tabela['DURAÇÃO_REAL_MIN'].apply(mins_to_text)
        df_tabela['PEÇAS'] = df_tabela['PEÇAS'].apply(lambda x: f"{int(x):,}".replace(",", "."))
        df_tabela['SKU'] = df_tabela['SKU'].apply(lambda x: f"{int(x)}")
        df_tabela = df_tabela[['AGENDA', 'CONFERENTE', 'CATEGORIA', 'STATUS_FISICO', 'PEÇAS', 'SKU', 'META (Tempo)', 'GASTO (Tempo)', 'PREVISÃO FIM', 'SITUAÇÃO META']]
        
        def cor_status(val):
            if '✅' in str(val): return 'color: #155724; background-color: #d4edda; font-weight: bold;'
            if '🔴' in str(val) or '⚠️' in str(val): return 'color: #721c24; background-color: #f8d7da; font-weight: bold;'
            if '⏳' in str(val): return 'color: #856404; background-color: #fff3cd; font-weight: bold;'
            return ''

        st.dataframe(df_tabela.style.applymap(cor_status, subset=['SITUAÇÃO META', 'PREVISÃO FIM']), use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("### 🤖 Piloto Automático de Fechamento")
        
        hora_atual = agora.strftime("%H:%M")
        data_hoje_str = agora.strftime('%d/%m/%Y')
        
        ja_salvou_hoje = False
        if not df_fechamento.empty and data_hoje_str in df_fechamento['DATA'].values:
            ja_salvou_hoje = True
            
        if ja_salvou_hoje:
            st.success(f"✅ O fechamento de hoje ({data_hoje_str}) já foi gravado no cofre automaticamente!")
        else:
            if hora_atual >= "17:20":
                st.info(f"🕒 Passou das 17:20! O Robô está iniciando a gravação do turno...")
                df_para_salvar = df_hoje_conf[(df_hoje_conf['STATUS_FISICO'] == 'OK') & (df_hoje_conf['DURAÇÃO_REAL_MIN'] > 0)].copy()
                
                if not df_para_salvar.empty:
                    df_export = pd.DataFrame({
                        'DATA': data_hoje_str, 'AGENDA': df_para_salvar['AGENDA'], 'CONFERENTE': df_para_salvar['CONFERENTE'],
                        'CATEGORIA': df_para_salvar['CATEGORIA'], 'PEÇAS': df_para_salvar['PEÇAS'],
                        'META MINUTOS': df_para_salvar['META_TEMPO_MIN'].round(2), 'REALIZADO MINUTOS': df_para_salvar['DURAÇÃO_REAL_MIN'].round(2),
                        'RESULTADO': df_para_salvar['SITUAÇÃO META'].apply(lambda x: 'NO PRAZO' if '✅' in x else 'ATRASADO')
                    })
                    with st.spinner("Gravando no Cofre de Dados..."):
                        sucesso = salvar_historico_fechamento(df_export)
                        if sucesso:
                            st.success(f"🎉 Robô salvou {len(df_export)} cargas no histórico com sucesso!")
                            st.cache_data.clear()
                            st.rerun()
                else:
                    st.warning("Deu a hora do fechamento, mas não havia nenhuma carga com status 'OK' e tempo registrado.")
            else:
                st.info(f"⏳ O robô está aguardando dar **17:20** para salvar os resultados. (Hora atual: {hora_atual})")
                
                if st.button("⚠️ Forçar Gravação Agora", type="secondary"):
                    df_para_salvar = df_hoje_conf[(df_hoje_conf['STATUS_FISICO'] == 'OK') & (df_hoje_conf['DURAÇÃO_REAL_MIN'] > 0)].copy()
                    if not df_para_salvar.empty:
                        df_export = pd.DataFrame({
                            'DATA': data_hoje_str, 'AGENDA': df_para_salvar['AGENDA'], 'CONFERENTE': df_para_salvar['CONFERENTE'],
                            'CATEGORIA': df_para_salvar['CATEGORIA'], 'PEÇAS': df_para_salvar['PEÇAS'],
                            'META MINUTOS': df_para_salvar['META_TEMPO_MIN'].round(2), 'REALIZADO MINUTOS': df_para_salvar['DURAÇÃO_REAL_MIN'].round(2),
                            'RESULTADO': df_para_salvar['SITUAÇÃO META'].apply(lambda x: 'NO PRAZO' if '✅' in x else 'ATRASADO')
                        })
                        sucesso = salvar_historico_fechamento(df_export)
                        if sucesso:
                            st.cache_data.clear()
                            st.rerun()
                    else:
                        st.warning("Nenhuma carga finalizada para salvar.")
    else:
        st.warning("⚠️ Planilhas de Conferência não encontradas ou vazias.")

# -------------------------------------------------------------------------
# ABA 3: RANKING DE CONFERENTES (HISTÓRICO BLINDADO MATEMATICAMENTE)
# -------------------------------------------------------------------------
with tab3:
    st.title("🏅 Desempenho Histórico: Equipe de Conferência")
    
    if not df_fechamento.empty:
        datas_disponiveis = df_fechamento['DATA'].unique()
        data_hist_sel = st.multiselect("Filtrar por Data do Histórico:", options=datas_disponiveis, default=datas_disponiveis)
        
        df_f = df_fechamento[df_fechamento['DATA'].isin(data_hist_sel)].copy()
        
        if not df_f.empty:
            # 1. Calcula a diferença de tempo exatamente pelo número salvo
            df_f['Desvio (Minutos)'] = df_f['REALIZADO MINUTOS'] - df_f['META MINUTOS']
            
            # 2. Ignora o texto gravado e RECALCULA o resultado baseado puramente na matemática!
            df_f['STATUS_REAL'] = df_f['Desvio (Minutos)'].apply(lambda x: 'ATRASADO' if x > 0 else 'NO PRAZO')
            
            # 3. Agrupa os dados
            ranking = df_f.groupby('CONFERENTE').agg(
                Cargas_Feitas=('AGENDA', 'count'),
                Atrasos=('STATUS_REAL', lambda x: (x == 'ATRASADO').sum()),
                No_Prazo=('STATUS_REAL', lambda x: (x == 'NO PRAZO').sum()),
                Tempo_Medio_Desvio=('Desvio (Minutos)', 'mean')
            ).reset_index()
            
            # 4. Calcula o % de Acerto
            ranking['% de Acerto'] = (ranking['No_Prazo'] / ranking['Cargas_Feitas']) * 100
            ranking = ranking.sort_values('% de Acerto', ascending=False)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### 🏆 Top Conferentes (% de Metas Batidas)")
                fig_bar = px.bar(ranking, x='CONFERENTE', y='% de Acerto', text_auto='.1f', 
                                 color='% de Acerto', color_continuous_scale='Greens',
                                 labels={'% de Acerto': 'Taxa de Acerto (%)'})
                fig_bar.update_layout(yaxis=dict(range=[0, 100]), coloraxis_showscale=False)
                st.plotly_chart(fig_bar, use_container_width=True)
                
            with col2:
                st.markdown("#### ⏳ Balanço de Tempo (Gargalo vs Agilidade)")
                st.caption("Barras vermelhas indicam tempo médio estourado. Verdes indicam agilidade (tempo salvo).")
                
                ranking_desvio = ranking.sort_values('Tempo_Medio_Desvio', ascending=False)
                cores = ['#F44336' if val > 0 else '#4CAF50' for val in ranking_desvio['Tempo_Medio_Desvio']]
                
                fig_desv = go.Figure(go.Bar(
                    x=ranking_desvio['CONFERENTE'], y=ranking_desvio['Tempo_Medio_Desvio'], 
                    marker_color=cores, text=ranking_desvio['Tempo_Medio_Desvio'].round(1), textposition='auto'
                ))
                fig_desv.update_layout(yaxis_title="Minutos (Média de Desvio)")
                st.plotly_chart(fig_desv, use_container_width=True)
                
            st.markdown("#### 📋 Detalhamento da Equipe")
            st.dataframe(ranking[['CONFERENTE', 'Cargas_Feitas', 'No_Prazo', 'Atrasos', '% de Acerto', 'Tempo_Medio_Desvio']].style.format({
                '% de Acerto': '{:.1f}%',
                'Tempo_Medio_Desvio': '{:.1f} min'
            }), use_container_width=True, hide_index=True)
            
        else:
            st.info("Nenhuma data selecionada.")
    else:
        st.info("📭 Nenhum dado de fechamento encontrado. Lembre-se de criar a aba 'FECHAMENTO' na planilha do Google Sheets.")
