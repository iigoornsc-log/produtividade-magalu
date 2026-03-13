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
st.set_page_config(page_title="Torre de Controle | Armazenagem", page_icon="⚡️", layout="wide")

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

# =========================================================================
# 2. MOTOR DE DADOS
# =========================================================================
@st.cache_data(ttl=300)
def carregar_dados():
    try:
        cred_dict = json.loads(st.secrets["google_json"])
        creds = Credentials.from_service_account_info(cred_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        client = gspread.authorize(creds)
        
        sh = client.open_by_key('1F4Qs5xGPMjgWSO6giHSwFfDf5F-mlv1RuT4riEVU0I0')
        ws = sh.worksheet("ACOMPANHAMENTO GERAL")
        data = ws.get_all_values()
        
        df = pd.DataFrame(data[1:], columns=data[0])
        
        df['QT_PRODUTO'] = pd.to_numeric(df['QT_PRODUTO'], errors='coerce').fillna(0)
        
        df['DT_CONFERENCIA'] = pd.to_datetime(df['DT_CONFERENCIA'], errors='coerce', dayfirst=True, format='mixed') 
        df['DT_ARMAZENAGEM'] = pd.to_datetime(df['DT_ARMAZENAGEM'], errors='coerce', dayfirst=True, format='mixed')
        
        df['AGENDA'] = df['AGENDA'].astype(str).str.strip().str.upper()
        df['FORNECEDOR'] = df['FORNECEDOR'].astype(str).str.strip().str.upper()
        df['OPERADOR'] = df['OPERADOR'].astype(str).str.strip().str.upper()
        
        df['Data_Conf'] = df['DT_CONFERENCIA'].dt.date
        df['Hora_Conf'] = df['DT_CONFERENCIA'].dt.strftime('%H:00')
        
        df['Data_Armz'] = df['DT_ARMAZENAGEM'].dt.date
        df['Hora_Armz'] = df['DT_ARMAZENAGEM'].dt.strftime('%H:00')
        
        df['Tempo_Espera_Minutos'] = (df['DT_ARMAZENAGEM'] - df['DT_CONFERENCIA']).dt.total_seconds() / 60.0
        
        return df
    except Exception as e:
        st.error(f"Erro na conexão: {e}")
        return pd.DataFrame()

df_bruto = carregar_dados()

if not df_bruto.empty:
    # -------------------------------------------------------------------------
    # PAINEL LATERAL: FILTROS E CONTROLES
    # -------------------------------------------------------------------------
    st.sidebar.image("https://magalog.com.br/opengraph-image.jpg?fdd536e7d35ec9da", width=250)
    st.sidebar.markdown("### 🎛️ Controles da Operação")
    
    data_max = df_bruto['Data_Conf'].dropna().max()
    data_sel = st.sidebar.date_input("🗓️ Data de Análise", data_max)
    
    modo_visao = st.sidebar.radio(
        "🔎 Modo de Análise",
        ["Líquida (Apenas do Dia)", "Global (Incluir Herança)"],
        help="Líquida: Foca estritamente no que entrou hoje. Global: Puxa todo o backlog de ontem que foi armazenado hoje e a fila real."
    )
    
    # LÓGICA DO MODO DE VISÃO
    df_hoje_conf = df_bruto[df_bruto['Data_Conf'] == data_sel]
    df_hoje_armz = df_bruto[df_bruto['Data_Armz'] == data_sel]
    
    if modo_visao == "Líquida (Apenas do Dia)":
        df_base = df_hoje_conf.copy()
        saldo_inicial = 0
    else:
        # Pega tudo que tocou no dia de hoje (Conferido hoje ou Armazenado hoje)
        df_base = pd.concat([df_hoje_conf, df_hoje_armz]).drop_duplicates(subset=['NU_ETIQUETA'])
        # Calcula a herança exata (Conferido antes de hoje e que amanheceu pendente)
        mask_heranca = (df_bruto['Data_Conf'] < data_sel) & ((df_bruto['Data_Armz'] >= data_sel) | (df_bruto['Data_Armz'].isnull()))
        saldo_inicial = df_bruto[mask_heranca]['NU_ETIQUETA'].nunique()

    # FILTRO DE EQUIPE (Múltipla Seleção)
    fantasmas = ['', 'NAN', 'NONE', 'NULL']
    operadores_validos = sorted([op for op in df_base['OPERADOR'].unique() if pd.notna(op) and op not in fantasmas])
    
    op_sel = st.sidebar.multiselect("👥 Filtrar Equipe (Remova intrusos):", options=operadores_validos, default=operadores_validos)
    
    # -------------------------------------------------------------------------
    # CONSTRUÇÃO DO DASHBOARD
    # -------------------------------------------------------------------------
    st.title(f"🚀 Torre de Controle | {data_sel.strftime('%d/%m/%Y')}")
    st.caption(f"Visualizando: **{modo_visao}**")
    
    # df_producao foca APENAS no que a equipe selecionada fez hoje
    df_producao = df_base[(df_base['Data_Armz'] == data_sel) & (df_base['OPERADOR'].isin(op_sel))]
    
    # BLOCO 1: KPIs
    c1, c2, c3, c4 = st.columns(4)
    
    qtd_etiquetas = df_producao['NU_ETIQUETA'].nunique()
    qtd_pecas = df_producao['QT_PRODUTO'].sum()
    espera_valida = df_producao[df_producao['Tempo_Espera_Minutos'] > 0]['Tempo_Espera_Minutos']
    sla_medio = espera_valida.mean() if not espera_valida.empty else 0
    txt_sla = f"{int(sla_medio // 60)}h {int(sla_medio % 60)}m"
    
    texto_op_kpi = "Equipe Total" if len(op_sel) == len(operadores_validos) else (op_sel[0] if len(op_sel) == 1 else f"{len(op_sel)} Operadores")
    
    with c1: exibir_kpi("Armazenados (Equipe)", f"{qtd_etiquetas:,.0f}".replace(',','.'), "Etiquetas bipadas", "#0086FF")
    with c2: exibir_kpi("Peças Processadas", f"{qtd_pecas:,.0f}".replace(',','.'), "Volume físico guardado", "#9B59B6")
    with c3: exibir_kpi("SLA Médio Doca", txt_sla, "Tempo em espera", "#F44336" if sla_medio > 120 else "#4CAF50")
    with c4: exibir_kpi("Filtro Ativo", texto_op_kpi, "Pessoas analisadas", "#FF9800")

    # BLOCO 2: FLUXO E PENDÊNCIAS (REALIDADE DA DOCA)
    st.markdown("<div class='bloco-header'>🌊 Fluxo de Trabalho e Pendências Acumuladas</div>", unsafe_allow_html=True)
    
    # Conferidos da Doca (Ignora o filtro de operador, pois demanda é da doca inteira)
    df_in = df_base[df_base['Data_Conf'] == data_sel].groupby('Hora_Conf')['NU_ETIQUETA'].nunique().reset_index(name='Conferidos')
    df_in.rename(columns={'Hora_Conf': 'Hora'}, inplace=True)
    
    # Armazenados Produzidos pela Equipe (Para a barra Azul)
    df_out_equipe = df_producao.groupby('Hora_Armz')['NU_ETIQUETA'].nunique().reset_index(name='Armazenados')
    df_out_equipe.rename(columns={'Hora_Armz': 'Hora'}, inplace=True)
    
    # Armazenados Totais do CD (Para calcular a pendência real sem inflar a doca)
    df_out_real_cd = df_base[df_base['Data_Armz'] == data_sel].groupby('Hora_Armz')['NU_ETIQUETA'].nunique().reset_index(name='Armz_Real')
    df_out_real_cd.rename(columns={'Hora_Armz': 'Hora'}, inplace=True)
    
    # Mesclagem Mestra
    df_fluxo = pd.merge(df_in, df_out_equipe, on='Hora', how='outer')
    df_fluxo = pd.merge(df_fluxo, df_out_real_cd, on='Hora', how='left').fillna(0).sort_values('Hora')
    
    df_fluxo['Acum_Conf'] = df_fluxo['Conferidos'].cumsum()
    df_fluxo['Acum_Armz_Real'] = df_fluxo['Armz_Real'].cumsum()
    
    # Cálculo Final da Pendência (Saldo Herança + Demandas Hoje - Saídas Reais Hoje)
    df_fluxo['Pendências'] = saldo_inicial + df_fluxo['Acum_Conf'] - df_fluxo['Acum_Armz_Real']
    df_fluxo['Pendências'] = df_fluxo['Pendências'].apply(lambda x: x if x > 0 else 0)

    max_y = df_fluxo[['Armazenados', 'Conferidos']].max().max() if not df_fluxo.empty else 10
    teto_grafico = max_y * 1.2 if max_y > 0 else 10

    fig_fluxo = go.Figure()
    
    fig_fluxo.add_trace(go.Bar(
        x=df_fluxo['Hora'], y=df_fluxo['Armazenados'], name='Armazenados (Equipe Filtrada)', 
        marker_color='#0086FF', text=df_fluxo['Armazenados'], textposition='auto', textfont=dict(color='white')
    ))
    
    fig_fluxo.add_trace(go.Bar(
        x=df_fluxo['Hora'], y=df_fluxo['Conferidos'], name='Conferidos (Demanda CD)', 
        marker_color='#9d26ff', text=df_fluxo['Conferidos'], textposition='outside', textfont=dict(color='#9d26ff')
    ))
    
    fig_fluxo.add_trace(go.Scatter(
        x=df_fluxo['Hora'], y=df_fluxo['Pendências'], name='Pendências (Fila Real da Doca)', mode='lines+markers+text', 
        line=dict(color='#E74C3C', width=3), yaxis='y2', text=df_fluxo['Pendências'], textposition='top center', textfont=dict(color='#E74C3C', weight='bold')
    ))
    
    fig_fluxo.update_layout(
        plot_bgcolor='rgba(0,0,0,0)', barmode='group',
        legend=dict(orientation="h", y=1.15, x=0.5, xanchor='center'),
        yaxis=dict(title="Qtd Etiquetas", showgrid=True, gridcolor='#F1F3F5', range=[0, teto_grafico]), 
        yaxis2=dict(title="Fila Acumulada", overlaying='y', side='right', showgrid=False),
        hovermode="x unified"
    )
    st.plotly_chart(fig_fluxo, use_container_width=True)

    # BLOCO 3: OPERADORES
    if len(op_sel) > 0:
        st.markdown("<div class='bloco-header'>👥 Performance da Equipe Selecionada</div>", unsafe_allow_html=True)
        col_rank, col_heat = st.columns([4, 6])
        
        with col_rank:
            st.markdown("##### 🏆 Ranking de Armazenagem")
            rank_op = df_producao.groupby('OPERADOR').agg({'NU_ETIQUETA': 'nunique', 'Hora_Armz': 'nunique'}).reset_index()
            rank_op.columns = ['Operador', 'Etiquetas', 'Horas']
            rank_op['Etq/Hora'] = (rank_op['Etiquetas'] / rank_op['Horas']).round(1)
            st.dataframe(rank_op.sort_values('Etiquetas', ascending=False)[['Operador', 'Etiquetas', 'Etq/Hora']], use_container_width=True, hide_index=True, height=350)
        
        with col_heat:
            st.markdown("##### 🔥 Calor de Produtividade (Armazenados/Hora)")
            df_heat = df_producao.groupby(['OPERADOR', 'Hora_Armz'])['NU_ETIQUETA'].nunique().reset_index()
            fig_heat = px.density_heatmap(df_heat, x="Hora_Armz", y="OPERADOR", z="NU_ETIQUETA", color_continuous_scale="Blues", text_auto=True)
            fig_heat.update_layout(plot_bgcolor='rgba(0,0,0,0)', yaxis=dict(title=""), xaxis_title="Hora", coloraxis_showscale=False)
            st.plotly_chart(fig_heat, use_container_width=True)
    else:
        st.warning("⚠️ Selecione ao menos um operador na barra lateral para ver os dados da equipe.")

else:
    st.error("⚠️ Dados não encontrados para a data selecionada.")
