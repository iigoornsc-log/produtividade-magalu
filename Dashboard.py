import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
import json

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

# Converter formato 00:00:00 para Minutos Totais
def time_to_mins(t_str):
    if pd.isna(t_str) or str(t_str).strip() == '': return 0
    try:
        partes = str(t_str).split(':')
        if len(partes) == 3:
            return int(partes[0]) * 60 + int(partes[1]) + float(partes[2]) / 60.0
        return 0
    except: return 0

# Formatar minutos de volta para texto (Ex: 1h 25m)
def mins_to_text(mins):
    if pd.isna(mins) or mins == 0: return "0m"
    h = int(mins // 60)
    m = int(mins % 60)
    if h > 0: return f"{h}h {m}m"
    return f"{m}m"

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
        df['QT_PRODUTO'] = pd.to_numeric(df['QT_PRODUTO'], errors='coerce').fillna(0)
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
        # Acha a aba independente de letras maiúsculas ou espaços sobrando
        aba_hist = next((aba for aba in todas_abas if "BASE DE DADOS" in aba.title.upper()), None)
        if not aba_hist:
            raise ValueError("Não encontrei a aba 'Base de Dados' na planilha 2.")
            
        # Puxa estritamente das colunas Q até W como você pediu!
        data_hist = aba_hist.get("Q:W")
        if not data_hist:
            raise ValueError("Colunas Q:W vazias na aba Base de Dados.")
            
        df_hist = pd.DataFrame(data_hist[1:], columns=data_hist[0])
        df_hist.columns = df_hist.columns.str.strip().str.upper()
        
        if 'TMP APC' in df_hist.columns:
            df_hist['TMP APC'] = pd.to_numeric(df_hist['TMP APC'], errors='coerce').fillna(0)
        if 'PEÇAS' in df_hist.columns:
            df_hist['PEÇAS'] = pd.to_numeric(df_hist['PEÇAS'], errors='coerce').fillna(0)
            
        # --- 2. PLANILHA DIA ATUAL ---
        aba_hoje = next((aba for aba in todas_abas if "DIA ATUAL" in aba.title.upper()), None)
        if not aba_hoje:
            raise ValueError("Não encontrei a aba 'Dia Atual' na planilha 2.")
            
        # Puxa estritamente das colunas A até H
        data_hoje = aba_hoje.get("A:H")
        if not data_hoje:
            raise ValueError("Colunas A:H vazias na aba Dia Atual.")
            
        df_hoje = pd.DataFrame(data_hoje[1:], columns=data_hoje[0])
        df_hoje.columns = df_hoje.columns.str.strip().str.upper()
        
        if 'PEÇAS' in df_hoje.columns:
            df_hoje['PEÇAS'] = pd.to_numeric(df_hoje['PEÇAS'], errors='coerce').fillna(0)
        if 'PÇS PENDENTES' in df_hoje.columns:
            df_hoje['PÇS PENDENTES'] = pd.to_numeric(df_hoje['PÇS PENDENTES'], errors='coerce').fillna(0)
        if 'DURAÇÃO CARGA' in df_hoje.columns:
            df_hoje['DURAÇÃO CARGA'] = df_hoje['DURAÇÃO CARGA'].astype(str).str.strip()
            
        return df_hist, df_hoje
        
    except Exception as e:
        st.error(f"Erro Conferência: {e}")
        return pd.DataFrame(), pd.DataFrame()

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
df_hist_conf, df_hoje_conf = carregar_dados_conferencia()

# Criando o Sistema de Abas (Tabs)
tab1, tab2 = st.tabs(["📦 Torre de Armazenagem (Doca)", "🔎 Torre de Conferência (Metas)"])

# -------------------------------------------------------------------------
# ABA 1: ARMAZENAGEM (O SEU DASHBOARD ATUAL PERFEITO)
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
            saldo_inicial = 0
        else:
            df_base_armz = pd.concat([df_hoje_c, df_hoje_a]).drop_duplicates(subset=['NU_ETIQUETA'])
            mask_heranca = (df_armz_filtrado['Data_Conf'] < data_sel) & ((df_armz_filtrado['Data_Armz'] >= data_sel) | (df_armz_filtrado['SITUACAO'] == '23'))
            saldo_inicial = df_armz_filtrado[mask_heranca]['NU_ETIQUETA'].nunique()

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
        qtd_pecas_armz = df_producao_equipe['QT_PRODUTO'].sum()
        qtd_pendentes_doca = df_base_armz[df_base_armz['SITUACAO'] == '23']['NU_ETIQUETA'].nunique()
        if modo_visao == "Global (Incluir Herança)":
            qtd_pendentes_doca += df_armz_filtrado[(df_armz_filtrado['Data_Conf'] < data_sel) & (df_armz_filtrado['SITUACAO'] == '23') & (df_armz_filtrado['CONFERENTE'].isin(conf_sel))]['NU_ETIQUETA'].nunique()
            
        espera_valida = df_producao_equipe[df_producao_equipe['Tempo_Espera_Minutos'] > 0]['Tempo_Espera_Minutos']
        sla_medio = espera_valida.mean() if not espera_valida.empty else 0
        txt_sla = f"{int(sla_medio // 60)}h {int(sla_medio % 60)}m"
        
        with c1: exibir_kpi("Armazenados (Sit. 25)", f"{qtd_etiquetas_armz:,.0f}", "Equipe selecionada", "#0086FF")
        with c2: exibir_kpi("Pendências (Sit. 23)", f"{qtd_pendentes_doca:,.0f}", "Fila total da Doca", "#E74C3C")
        with c3: exibir_kpi("SLA Médio Doca", txt_sla, "Tempo em espera", "#F44336" if sla_medio > 120 else "#4CAF50")
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
# ABA 2: CONFERÊNCIA (O SEU NOVO MOTOR DE METAS PREDITIVAS)
# -------------------------------------------------------------------------
with tab2:
    if not df_hist_conf.empty and not df_hoje_conf.empty:
        st.title("🎯 Produtividade e Metas de Conferência")
        st.caption("Cálculo preditivo de metas baseado no histórico do fornecedor e categoria.")
        
        # 1. Aprender com o Histórico (Machine Learning Básico)
        taxas_hist = df_hist_conf.groupby(['FORNECEDOR', 'LINHA']).agg({'TMP APC':'sum', 'PEÇAS':'sum'}).reset_index()
        taxas_hist = taxas_hist[taxas_hist['PEÇAS'] > 0]
        taxas_hist['VELOCIDADE_MIN_PECA'] = taxas_hist['TMP APC'] / taxas_hist['PEÇAS']
        
        taxa_global_cd = df_hist_conf['TMP APC'].sum() / df_hist_conf['PEÇAS'].sum() if df_hist_conf['PEÇAS'].sum() > 0 else 1.0

        # 2. Aplicar Metas no Dia Atual
        df_hoje_conf['DURAÇÃO_REAL_MIN'] = df_hoje_conf['DURAÇÃO CARGA'].apply(time_to_mins)
        
        # Cruzar a carga de hoje com a velocidade aprendida
        df_analise = pd.merge(df_hoje_conf, taxas_hist[['FORNECEDOR', 'LINHA', 'VELOCIDADE_MIN_PECA']], 
                              left_on=['ORIGEM', 'CATEGORIA'], right_on=['FORNECEDOR', 'LINHA'], how='left')
        
        # Se for um fornecedor novo, usa a média global do CD
        df_analise['VELOCIDADE_MIN_PECA'] = df_analise['VELOCIDADE_MIN_PECA'].fillna(taxa_global_cd)
        
        # A Mágica: Calcular a Meta de Tempo
        df_analise['META_TEMPO_MIN'] = df_analise['PEÇAS'] * df_analise['VELOCIDADE_MIN_PECA']
        
        # Calculando o Status
        def calcular_status(row):
            if row['DURAÇÃO_REAL_MIN'] == 0: return "⏳ Em Andamento"
            if row['DURAÇÃO_REAL_MIN'] <= row['META_TEMPO_MIN']: return "✅ No Prazo"
            return "🔴 Atrasado"
            
        df_analise['STATUS'] = df_analise.apply(calcular_status, axis=1)
        
        # KPIs de Conferência
        c1, c2, c3, c4 = st.columns(4)
        cargas_concluidas = df_analise[df_analise['DURAÇÃO_REAL_MIN'] > 0].shape[0]
        cargas_atrasadas = df_analise[df_analise['STATUS'] == "🔴 Atrasado"].shape[0]
        perc_acerto = 100 - ((cargas_atrasadas / cargas_concluidas) * 100) if cargas_concluidas > 0 else 0
        
        with c1: exibir_kpi("Agendas do Dia", len(df_analise), "Atribuídas a conferentes", "#9B59B6")
        with c2: exibir_kpi("Concluídas", cargas_concluidas, "Finalizadas", "#0086FF")
        with c3: exibir_kpi("Aderência à Meta", f"{perc_acerto:.1f}%", "Cargas dentro do tempo", "#4CAF50" if perc_acerto > 80 else "#F44336")
        with c4: exibir_kpi("Conferentes Ativos", df_analise['CONFERENTE'].nunique(), "Equipe na doca", "#FF9800")
        
        st.markdown("<div class='bloco-header'>📊 Tabela de Produtividade vs Meta</div>", unsafe_allow_html=True)
        
        # Preparar tabela bonita
        df_tabela = df_analise[['AGENDA', 'CONFERENTE', 'ORIGEM', 'CATEGORIA', 'PEÇAS', 'META_TEMPO_MIN', 'DURAÇÃO_REAL_MIN', 'STATUS']].copy()
        df_tabela['META_TEMPO'] = df_tabela['META_TEMPO_MIN'].apply(mins_to_text)
        df_tabela['TEMPO_REAL'] = df_tabela['DURAÇÃO_REAL_MIN'].apply(mins_to_text)
        
        # Reordenar para esconder os minutos brutos
        df_tabela = df_tabela[['AGENDA', 'CONFERENTE', 'ORIGEM', 'CATEGORIA', 'PEÇAS', 'META_TEMPO', 'TEMPO_REAL', 'STATUS']]
        
        st.dataframe(df_tabela, use_container_width=True, hide_index=True)
        
        # Gráfico de Dispersão (Desvio de Meta)
        st.markdown("<div class='bloco-header'>⚖️ Análise de Desvio: Meta vs Realizado</div>", unsafe_allow_html=True)
        df_grafico_conf = df_analise[df_analise['DURAÇÃO_REAL_MIN'] > 0]
        
        if not df_grafico_conf.empty:
            fig_disp = px.scatter(df_grafico_conf, x="META_TEMPO_MIN", y="DURAÇÃO_REAL_MIN", 
                                  color="STATUS", hover_data=['AGENDA', 'CONFERENTE'],
                                  color_discrete_map={"✅ No Prazo": "#4CAF50", "🔴 Atrasado": "#F44336"},
                                  labels={"META_TEMPO_MIN": "Meta de Tempo (Minutos)", "DURAÇÃO_REAL_MIN": "Tempo Gasto Real (Minutos)"})
            # Adiciona a linha de tendência "Perfeita"
            fig_disp.add_shape(type="line", line=dict(dash='dash', color='gray'),
                               x0=0, y0=0, x1=df_grafico_conf['META_TEMPO_MIN'].max(), y1=df_grafico_conf['META_TEMPO_MIN'].max())
            st.plotly_chart(fig_disp, use_container_width=True)
        
    else:
        st.warning("⚠️ Planilhas de Conferência não encontradas ou vazias.")
