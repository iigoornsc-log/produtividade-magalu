import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
import datetime
import time
from datetime import date

# ==========================================================
# 1. CONFIGURAÇÃO DA PÁGINA E CSS (THEME MAGALOG CORPORATIVO)
# ==========================================================
st.set_page_config(page_title="MAGALOG | Torre de produtividade recebimento", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    /* 1. FONTE PREMIUM E ÍCONES GOOGLE MATERIAL */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
    @import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@24,600,1,0');
    
    * { font-family: 'Inter', sans-serif !important; }

    /* Classe para alinhar os ícones no HTML perfeitamente com o texto */
    .icon-MAGALOG {
        font-family: 'Material Symbols Rounded' !important;
        font-variation-settings: 'FILL' 1, 'wght' 600, 'GRAD' 0, 'opsz' 24;
        font-weight: normal;
        font-style: normal;
        letter-spacing: normal;
        text-transform: none;
        white-space: nowrap;
        direction: ltr;
        -webkit-font-smoothing: antialiased;
        vertical-align: middle;
        display: inline-block;
        line-height: 1;
        font-size: inherit;
    }

    /* --- CORREÇÃO DOS ÍCONES DA SIDEBAR E SISTEMA --- */
    .material-icons, .material-symbols-rounded, [data-testid="stSidebarCollapseButton"] * {
        font-family: 'Material Symbols Rounded', 'Material Icons', sans-serif !important;
    }

    /* 2. ANIMAÇÃO RGB LUIZALABS */
    @keyframes Glow {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* Linha Tech Animada no Topo da Tela */
    .stApp::before {
        content: ""; position: fixed; top: 0; left: 0; right: 0; height: 5px;
        background: linear-gradient(90deg, #0086FF, #FF007F, #00C853, #0086FF, #FF007F);
        background-size: 300% 300%;
        animation: MAGALOGGlow 6s linear infinite;
        z-index: 999999;
    }

    /* Fundo da Aplicação (Soft com micro-gradiente) */
    .stApp {
        background-color: #F0F4F8;
        background-image: radial-gradient(circle at 100% 0%, #E2EDF8 0%, transparent 40%);
    }

    /* 3. TÍTULOS COM DEGRADÊ METÁLICO */
    .MAGALOG-page-title { 
        background: linear-gradient(135deg, #0086FF 0%, #001A57 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 32px; font-weight: 900; letter-spacing: -1px; margin-bottom: 2px;
    }
    .MAGALOG-page-subtitle { color: #64748B; font-size: 15px; font-weight: 500; margin-bottom: 25px; }

    /* 4. ABAS (TABS) CORPORATIVAS */
    [data-baseweb="tab-list"] {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 6px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.04);
        gap: 8px;
        margin-bottom: 25px;
    }
    button[data-baseweb="tab"] {
        background-color: transparent !important;
        border-radius: 8px !important;
        padding: 12px 24px !important;
        color: #64748B !important;
        font-weight: 600 !important;
        border: none !important;
        transition: all 0.3s ease !important;
    }
    button[data-baseweb="tab"]:hover {
        background-color: #F1F5F9 !important;
        color: #0F172A !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #0086FF !important;
        color: #FFFFFF !important;
        box-shadow: 0 4px 15px rgba(0,134,255,0.35) !important;
    }
    [data-baseweb="tab-border"] { display: none; }

    /* 5. RIBBONS ANIMADOS */
    .MAGALOG-ribbon {
        background: linear-gradient(90deg, #0086FF, #005BFF, #FF007F, #0086FF);
        background-size: 300% 300%;
        animation: MAGALOGGlow 8s ease infinite;
        color: #FFFFFF; padding: 8px 24px; font-size: 13px; font-weight: 700;
        border-radius: 0px 8px 8px 0px; margin-bottom: 15px; margin-top: 10px;
        position: relative; left: -1rem; box-shadow: 0 4px 15px rgba(0,134,255,0.3);
        text-transform: uppercase; letter-spacing: 1px;
    }

    /* 6. CARDS GLASSMORPHISM */
    .MAGALOG-card, div[data-testid="stVerticalBlock"] > div > div[data-testid="stVerticalBlockBorderWrapper"] {
        background: rgba(255, 255, 255, 0.95) !important;
        backdrop-filter: blur(10px) !important;
        border: 1px solid rgba(255,255,255,0.6) !important;
        border-radius: 16px !important;
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.03) !important;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        padding: 20px !important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"]:hover {
        transform: translateY(-3px);
        box-shadow: 0 12px 40px rgba(0, 134, 255, 0.08) !important;
    }

    /* DATE INPUT SIDEBAR */
    section[data-testid="stSidebar"] div[data-baseweb="input"] > div {
        background: #FFFFFF !important; border-radius: 12px !important;
        border: 1px solid #E2E8F0 !important; min-height: 48px !important; transition: all 0.3s ease !important;
    }
    section[data-testid="stSidebar"] div[data-baseweb="input"] > div:focus-within {
        border-color: #0086FF !important; box-shadow: 0 0 0 3px rgba(0,134,255,0.15) !important;
    }
    section[data-testid="stSidebar"] input { color: #0F172A !important; font-weight: 600 !important; }
    section[data-testid="stSidebar"] svg { color: #0086FF !important; }

    /* CHIPS MULTISELECT */
    span[data-baseweb="tag"] {
        background-color: #E6F2FF !important; color: #0086FF !important;
        border-radius: 8px !important; border: 1px solid #BAE6FD !important;
        font-weight: 700 !important; padding: 6px 12px !important; margin: 4px 4px 4px 0px !important;
    }
    span[data-baseweb="tag"] svg { fill: #0086FF !important; } 

    /* BOTÃO SIDEBAR */
    section[data-testid="stSidebar"] .stButton>button {
        background: linear-gradient(135deg, #0086FF 0%, #005BFF 100%); color: #FFFFFF;
        border: none; border-radius: 12px; font-weight: 700; font-size: 14px;
        padding: 0.8rem 1rem; box-shadow: 0 6px 20px rgba(0,134,255,0.25); transition: all 0.3s ease;
    }
    section[data-testid="stSidebar"] .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 10px 25px rgba(0,134,255,0.4); }
    section[data-testid="stSidebar"] .stButton>button:active { transform: scale(0.97); }
    
    /* BOTÃO PRIMARY (Verde) */
    button[kind="primary"] {
        background: linear-gradient(135deg, #00C853 0%, #009624 100%) !important; border: none !important; color: white !important; font-weight: 800 !important;
        border-radius: 10px !important; box-shadow: 0 6px 20px rgba(0,200,83,0.3) !important; transition: all 0.3s ease !important;
    }
    button[kind="primary"]:hover { box-shadow: 0 8px 25px rgba(0,200,83,0.5) !important; transform: translateY(-2px) !important; }

    /* SCROLLBAR */
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 10px; }
    ::-webkit-scrollbar-thumb:hover { background: #94A3B8; }

    /* KPI CARDS */
    .kpi-card { 
        background: #FFFFFF; border-radius: 16px; padding: 20px 15px; 
        border: 1px solid #F1F5F9; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.03);
        margin-bottom: 12px; display: flex; flex-direction: column; 
        align-items: center; justify-content: center; transition: all 0.3s ease;
    }
    .kpi-card:hover { transform: translateY(-4px); box-shadow: 0 12px 30px rgba(0, 134, 255, 0.08); }
    .kpi-title { color: #64748B; font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;}
    .kpi-value { color: #0F172A; font-size: 24px; font-weight: 900; letter-spacing: -0.5px; }

    /* SIDEBAR E DASHBOARD PREMIUM */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #082a63 0%, #00153d 100%) !important;
        border-right: 1px solid rgba(255,255,255,0.08);
    }
    section[data-testid="stSidebar"] * { color: #EAF2FF; }
    section[data-testid="stSidebar"] .stRadio > div { gap: 8px; }
    section[data-testid="stSidebar"] label[data-baseweb="radio"] {
        background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.08);
        border-radius: 14px; padding: 10px 12px; margin-bottom: 8px; transition: all 0.25s ease;
    }
    section[data-testid="stSidebar"] label[data-baseweb="radio"]:hover {
        background: rgba(255,255,255,0.12); border-color: rgba(85,170,255,0.55);
    }
    .MAGALOG-shell {
        background: linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(247,250,255,0.98) 100%);
        border: 1px solid rgba(255,255,255,0.7); border-radius: 24px; box-shadow: 0 20px 60px rgba(15, 23, 42, 0.08);
        padding: 28px; margin-bottom: 22px;
    }
    .MAGALOG-hero {
        position: relative; overflow: hidden; background: linear-gradient(135deg, #0A4FB3 0%, #062B76 45%, #0D1836 100%);
        border-radius: 28px; padding: 34px 34px 26px 34px; color: #FFFFFF;
        box-shadow: 0 24px 70px rgba(0, 74, 173, 0.25); margin-bottom: 24px;
    }
    .MAGALOG-hero::after {
        content: ""; position: absolute; inset: auto -80px -90px auto; width: 320px; height: 320px;
        background: radial-gradient(circle, rgba(0,255,255,0.22) 0%, rgba(0,255,255,0) 68%);
    }
    .MAGALOG-hero::before {
        content: ""; position: absolute; inset: 0; background: linear-gradient(90deg, rgba(255,255,255,0.05), rgba(255,255,255,0)); pointer-events: none;
    }
    .MAGALOG-hero-badge {
        display: inline-block; background: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.15);
        border-radius: 999px; padding: 8px 14px; font-size: 12px; font-weight: 800; letter-spacing: .08em;
        margin-bottom: 16px; text-transform: uppercase; backdrop-filter: blur(8px);
    }
    .MAGALOG-hero-title {
        font-size: 40px; font-weight: 900; line-height: 1.02; letter-spacing: -1.2px;
        margin-bottom: 10px; position: relative; z-index: 2;
    }
    .MAGALOG-hero-subtitle {
        color: rgba(255,255,255,0.82); font-size: 16px; max-width: 860px; position: relative; z-index: 2;
    }
    .MAGALOG-grid {
        display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin: 22px 0;
    }
    .MAGALOG-feature-card {
        background: linear-gradient(180deg, #ffffff 0%, #f6f9ff 100%); border: 1px solid #E5EEF9;
        border-radius: 22px; padding: 22px; box-shadow: 0 12px 35px rgba(15,23,42,0.06); min-height: 150px;
    }
    .MAGALOG-feature-icon {
        width: 52px; height: 52px; border-radius: 16px; display:flex; align-items:center;
        justify-content:center; font-size: 24px; margin-bottom: 14px;
        background: linear-gradient(135deg, #E9F3FF 0%, #DDEBFF 100%); color: #0A4FB3;
    }
    .MAGALOG-feature-title { color: #0F172A; font-size: 18px; font-weight: 800; margin-bottom: 6px; }
    .MAGALOG-feature-text { color: #64748B; font-size: 14px; line-height: 1.5; }
    .MAGALOG-info-strip {
        background: linear-gradient(90deg, #0A57C9 0%, #0094FF 55%, #FF6B3D 100%); border-radius: 18px;
        padding: 16px 20px; color: #fff; font-weight: 800; letter-spacing: .04em; text-transform: uppercase;
        box-shadow: 0 14px 40px rgba(0,134,255,0.18); margin-top: 10px;
    }
    .MAGALOG-mini-card {
        background: rgba(255,255,255,0.94); border: 1px solid #E6EDF7; border-radius: 18px;
        padding: 18px; box-shadow: 0 10px 30px rgba(15,23,42,0.05); height: 100%;
    }
    .MAGALOG-mini-label { color: #64748B; font-size: 11px; text-transform: uppercase; font-weight: 800; letter-spacing: .06em; margin-bottom: 6px; }
    .MAGALOG-mini-value { color: #0F172A; font-size: 28px; font-weight: 900; letter-spacing: -0.8px; margin-bottom: 4px; }
    .MAGALOG-mini-desc { color: #64748B; font-size: 13px; }
    </style>
""", unsafe_allow_html=True)

def exibir_kpi(titulo, valor, subtitulo="", cor="#0086FF"):
    st.markdown(f"""
    <div class="kpi-card" style="border-top: 4px solid {cor};">
        <div class="kpi-title">{titulo}</div>
        <div class="kpi-value">{valor}</div>
        <div style="font-size: 11px; color: {cor}; opacity: 0.8; font-weight: 600; margin-top: 4px;">{subtitulo}</div>
    </div>
    """, unsafe_allow_html=True)

# --- FUNÇÕES DE LIMPEZA E BACK-END ---
def limpa_texto(valor):
    if pd.isna(valor): return ""
    return str(valor).strip().upper()

def limpa_agenda(valor):
    if pd.isna(valor) or str(valor).strip() in ['', 'NAN', 'NULL', 'NONE']: return ""
    v = str(valor).strip().upper()
    if v.endswith('.0'): v = v[:-2] 
    return v

def limpa_numero_br(valor):
    if pd.isna(valor) or str(valor).strip() in ['', 'NAN', 'NULL', 'NONE']: return 0
    v = str(valor).strip()
    if ',' in v: v = v.replace('.', '').replace(',', '.')
    else: v = v.replace('.', '')
    try: return float(v)
    except: return 0

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

# ==========================================================
# 2. MOTOR DE DADOS MULTI-PLANILHAS
# ==========================================================
@st.cache_data(ttl=3) 
def ler_cofre_vivo():
    try:
        cred_dict = json.loads(st.secrets["google_json"])
        creds = Credentials.from_service_account_info(cred_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        client = gspread.authorize(creds)
        sh2 = client.open_by_key('1bj5vIu8LOIWqaW5evogwQeyrJd9yj1iQkXHbJKvTeks')
        aba_fechamento = sh2.worksheet("FECHAMENTO")
        data_fech = aba_fechamento.get_all_values()
        if len(data_fech) > 1:
            df = pd.DataFrame(data_fech[1:], columns=data_fech[0])
            df.columns = df.columns.str.strip().str.upper()
            if 'DATA' in df.columns: df['DATA'] = df['DATA'].apply(limpa_texto)
            if 'AGENDA' in df.columns: df['AGENDA'] = df['AGENDA'].apply(limpa_agenda)
            if 'META MINUTOS' in df.columns: df['META MINUTOS'] = df['META MINUTOS'].apply(limpa_numero_br)
            if 'REALIZADO MINUTOS' in df.columns: df['REALIZADO MINUTOS'] = df['REALIZADO MINUTOS'].apply(limpa_numero_br)
            return df
        return pd.DataFrame()
    except:
        return pd.DataFrame()

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
        df['NU_ETIQUETA'] = df['NU_ETIQUETA'].apply(limpa_texto)
        df['AGENDA'] = df['AGENDA'].apply(limpa_agenda)
        df['PRODUTO'] = df['PRODUTO'].apply(limpa_texto)
        df['QT_PRODUTO'] = df['QT_PRODUTO'].apply(limpa_numero_br)
        df['SITUACAO'] = df['SITUACAO'].apply(limpa_texto)
        df['OPERADOR'] = df['OPERADOR'].apply(limpa_texto)
        df['CONFERENTE'] = df['CONFERENTE'].apply(limpa_texto)
        
        # Datas reais do que aconteceu
        df['DT_CONFERENCIA_CALC'] = pd.to_datetime(df['DT_CONFERENCIA'], errors='coerce') 
        df['DT_ARMAZENAGEM_CALC'] = pd.to_datetime(df['DT_ARMAZENAGEM'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
        
        df['Data_Conf'] = df['DT_CONFERENCIA_CALC'].dt.date
        df['Data_Armz'] = df['DT_ARMAZENAGEM_CALC'].dt.date 
        
        def formata_hora(h):
            if pd.isna(h) or str(h).strip() in ['', 'NAN', 'NULL', 'NONE']: return None
            try: return f"{int(float(h)):02d}:00"
            except: return None
            
        df['Hora_Conf'] = df['HORA CONF'].apply(formata_hora)
        df['Hora_Armz'] = df['HORA ARMZ'].apply(formata_hora)
        df['Tempo_Espera_Minutos'] = (df['DT_ARMAZENAGEM_CALC'] - df['DT_CONFERENCIA_CALC']).dt.total_seconds() / 60.0
        
        df['Data_Ref'] = pd.to_datetime(df['DATA'], format='%d/%m/%Y', errors='coerce').dt.date
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
        
        aba_hist = next((aba for aba in todas_abas if "BASE DE DADOS" in aba.title.upper()), None)
        if aba_hist:
            data_hist = aba_hist.get("Q:W")
            if data_hist and len(data_hist) > 0:
                df_hist = pd.DataFrame(data_hist[1:], columns=data_hist[0])
                df_hist.columns = df_hist.columns.str.strip().str.upper()
                if 'TMP APC' in df_hist.columns: df_hist['TMP APC'] = df_hist['TMP APC'].apply(limpa_numero_br)
                if 'PEÇAS' in df_hist.columns: df_hist['PEÇAS'] = df_hist['PEÇAS'].apply(limpa_numero_br)
                if 'SKU' in df_hist.columns: df_hist['SKU'] = df_hist['SKU'].apply(limpa_numero_br)
            else: df_hist = pd.DataFrame()
        else: df_hist = pd.DataFrame()
            
        aba_hoje = next((aba for aba in todas_abas if "DIA ATUAL" in aba.title.upper()), None)
        if aba_hoje:
            data_hoje = aba_hoje.get("A:I")
            if data_hoje and len(data_hoje) > 0:
                df_hoje = pd.DataFrame(data_hoje[1:], columns=data_hoje[0])
                df_hoje.columns = df_hoje.columns.str.strip().str.upper()
                if 'AGENDA' in df_hoje.columns: df_hoje['AGENDA'] = df_hoje['AGENDA'].apply(limpa_agenda)
                if 'STATUS' in df_hoje.columns: df_hoje.rename(columns={'STATUS': 'STATUS_FISICO'}, inplace=True)
                if 'PEÇAS' in df_hoje.columns: df_hoje['PEÇAS'] = df_hoje['PEÇAS'].apply(limpa_numero_br)
                if 'SKU' in df_hoje.columns: df_hoje['SKU'] = df_hoje['SKU'].apply(limpa_numero_br)
                if 'DURAÇÃO CARGA' in df_hoje.columns: df_hoje['DURAÇÃO CARGA'] = df_hoje['DURAÇÃO CARGA'].astype(str).str.strip()
            else: df_hoje = pd.DataFrame()
        else: df_hoje = pd.DataFrame()
            
        return df_hist, df_hoje
    except Exception as e:
        st.error(f"Erro detalhado na conexão da Conferência: {e}")
        return pd.DataFrame(), pd.DataFrame()

@st.dialog("RAIO-X DA HORA: DETALHAMENTO", width="large")
def popup_detalhe_hora(hora, df_base, data_sel):
    df_conferido = df_base[(df_base['Hora_Conf'] == hora) & (df_base['Data_Conf'] == data_sel)].copy()
    df_armazenado = df_base[(df_base['Hora_Armz'] == hora) & (df_base['Data_Armz'] == data_sel) & (df_base['SITUACAO'].isin(['25', 'AGRUPADO', 'NORMAL']))].copy()
    df_hora = pd.concat([df_conferido, df_armazenado]).drop_duplicates(subset=['NU_ETIQUETA'])
    
    if df_hora.empty:
        st.warning(f"Nenhuma movimentação às {hora}.")
        return
        
    st.markdown(f"### <span class='icon-MAGALOG'>schedule</span> Resumo das **{hora}**", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Conferidas Aqui", df_conferido['NU_ETIQUETA'].nunique())
    c2.metric("Armazenadas Aqui", df_armazenado['NU_ETIQUETA'].nunique())
    c3.metric("Peças Movimentadas", f"{df_hora['QT_PRODUTO'].sum():,.0f}".replace(',','.'))
    c4.metric("Agendas Envolvidas", df_hora['AGENDA'].nunique())
    
    df_exibicao = df_hora[['NU_ETIQUETA', 'SITUACAO', 'PRODUTO', 'CONFERENTE', 'OPERADOR', 'Hora_Conf', 'Hora_Armz']].copy()
    st.dataframe(df_exibicao, use_container_width=True, hide_index=True)

# ==========================================================
# 3. INTERFACE E ABAS PRINCIPAIS (COM SIDEBAR MAGALOG)
# ==========================================================
st.sidebar.markdown("""
    <div style="padding: 10px 8px 4px 8px; margin-bottom: 14px;">
        <div style="background: linear-gradient(135deg, rgba(255,255,255,0.12), rgba(255,255,255,0.05)); border:1px solid rgba(255,255,255,0.10); border-radius: 22px; padding: 18px 16px; box-shadow: inset 0 1px 0 rgba(255,255,255,0.08);">
            <div style="font-size: 13px; font-weight: 800; letter-spacing:.08em; text-transform: uppercase; color:#9CC8FF; margin-bottom:8px;">MAGALOG</div>
            <div style="font-size: 26px; font-weight: 900; line-height:1.0; color:#FFFFFF; margin-bottom:8px;">Produtividade recebimento</div>
            <div style="font-size: 13px; color:rgba(255,255,255,0.72);">Operação, equipe e performance logística.</div>
            <div style="height: 8px; margin-top:14px; border-radius: 999px; background: linear-gradient(90deg, #0086FF, #00D2FF, #FF8A3D, #FF4D6D); background-size:300% 300%; animation: Glow 7s linear infinite;"></div>
        </div>
    </div>
""", unsafe_allow_html=True)

st.sidebar.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
if st.sidebar.button("Sincronizar Data", type="secondary", use_container_width=True):
    with st.spinner("Puxando dados da base..."):
        st.cache_data.clear()
        st.rerun()

st.markdown("""
    <div class="MAGALOG-hero">
        <div class="MAGALOG-hero-badge">Plataforma Operacional</div>
        <div class="MAGALOG-hero-title">Central de Operações Logísticas</div>
        <div class="MAGALOG-hero-subtitle">Acompanhamento preditivo inteligente de conferência, ranking de equipes e controle da doca de armazenagem.</div>
    </div>
""", unsafe_allow_html=True)

df_armz = carregar_dados_armazenagem()
df_hist_conf, df_hoje_conf = carregar_dados_conferencia()
df_fechamento = ler_cofre_vivo()

# CORREÇÃO 1: Remoção de todos os Emojis nas abas. Texto limpo.
tab1, tab2, tab3 = st.tabs(["Torre de Armazenagem (Doca)", "Torre de Conferência (Metas)", "Desempenho da Equipe"])

# -------------------------------------------------------------------------
# ABA 1: ARMAZENAGEM (LÓGICA DE RANKING ABSOLUTO)
# -------------------------------------------------------------------------
with tab1:
    if not df_armz.empty:
        df_armz_filtrado = df_armz[df_armz['SITUACAO'].isin(['23', '25'])]
        
        st.sidebar.markdown("### <span class='icon-MAGALOG'>tune</span> Filtros de Análise", unsafe_allow_html=True)
        
        data_max = df_armz['Data_Conf'].dropna().max()
        data_sel = st.sidebar.date_input("Data de Análise (Armaz.)", data_max)
        
        modo_visao = st.sidebar.radio("Modo de Análise da Doca", ["Líquida (Apenas do Dia)", "Global (Incluir Herança)"])
        
        fantasmas = ['', 'NAN', 'NONE', 'NULL']
        
        df_hoje_c = df_armz[df_armz['Data_Conf'] == data_sel]
        conferentes_validos = sorted([c for c in df_hoje_c['CONFERENTE'].unique() if pd.notna(c) and c not in fantasmas])
        conf_sel = st.sidebar.multiselect("Equipe de Conferência:", options=conferentes_validos, default=conferentes_validos)
        
        df_pura_armz = df_armz[df_armz['Data_Armz'] == data_sel]
        operadores_validos = sorted([op for op in df_pura_armz['OPERADOR'].unique() if pd.notna(op) and op not in fantasmas])
        op_sel = st.sidebar.multiselect("Equipe de Armazenagem:", options=operadores_validos, default=operadores_validos)
        
        df_producao_real = df_pura_armz[df_pura_armz['OPERADOR'].isin(op_sel)]
        
        c1, c2, c3, c4 = st.columns(4)
        qtd_etiquetas_armz = df_producao_real['NU_ETIQUETA'].nunique()
        
        if modo_visao == "Líquida (Apenas do Dia)":
            df_fila = df_hoje_c[(df_hoje_c['SITUACAO'] == '23') & (df_hoje_c['CONFERENTE'].isin(conf_sel))]
        else:
            df_fila = df_armz[(df_armz['SITUACAO'] == '23') & (df_armz['CONFERENTE'].isin(conf_sel)) & (df_armz['Data_Conf'] <= data_sel)]
            
        qtd_pendentes_doca = df_fila['NU_ETIQUETA'].nunique()
            
        espera_valida = df_producao_real[df_producao_real['Tempo_Espera_Minutos'] > 0]['Tempo_Espera_Minutos']
        sla_medio = espera_valida.mean() if not espera_valida.empty else 0
        
        with c1: exibir_kpi("Armazenados", f"{qtd_etiquetas_armz:,.0f}", "Etiquetas Guardadas na Data", "#0086FF")
        with c2: exibir_kpi("Fila da Doca", f"{qtd_pendentes_doca:,.0f}", "Pendentes de Armazenamento", "#E74C3C")
        with c3: exibir_kpi("Tempo de Espera", mins_to_text(sla_medio), "SLA Médio na Doca", "#F44336" if sla_medio > 120 else "#10B981")
        with c4: exibir_kpi("Operadores", str(len(op_sel)), "Quantidade no Dia", "#F59E0B")

        col_tit, col_sel = st.columns([7, 3])
        
        with col_tit: st.markdown("<h4 style='color: #334155; margin-bottom: 15px; margin-top: 40px;'><span class='icon-MAGALOG'>insights</span> Fluxo de Trabalho e Capacidade</h4>", unsafe_allow_html=True)
        
        horas_conf = df_hoje_c['Hora_Conf'].dropna().unique()
        horas_armz = df_producao_real['Hora_Armz'].dropna().unique()
        todas_horas = sorted(list(set(list(horas_conf) + list(horas_armz))))
        
        with col_sel:
            st.markdown("<br>", unsafe_allow_html=True)
            hora_manual = st.selectbox("Inspecionar Hora:", ["Selecione..."] + todas_horas)

        dados_grafico = []
        for hora in todas_horas:
            conf_hora = df_hoje_c[(df_hoje_c['Hora_Conf'] == hora) & (df_hoje_c['CONFERENTE'].isin(conf_sel))]['NU_ETIQUETA'].nunique()
            armz_hora = df_producao_real[df_producao_real['Hora_Armz'] == hora]['NU_ETIQUETA'].nunique()
            
            if modo_visao == "Líquida (Apenas do Dia)":
                entrou = df_hoje_c[(df_hoje_c['Hora_Conf'] <= hora) & (df_hoje_c['CONFERENTE'].isin(conf_sel))]['NU_ETIQUETA'].nunique()
                saiu = df_producao_real[(df_producao_real['Hora_Armz'] <= hora) & (df_producao_real['Data_Conf'] == data_sel)]['NU_ETIQUETA'].nunique()
                pendencias = max(0, entrou - saiu)
            else:
                entrou = df_armz[(df_armz['CONFERENTE'].isin(conf_sel)) & ((df_armz['Data_Conf'] < data_sel) | ((df_armz['Data_Conf'] == data_sel) & (df_armz['Hora_Conf'] <= hora)))]['NU_ETIQUETA'].nunique()
                saiu = df_armz[(df_armz['CONFERENTE'].isin(conf_sel)) & (df_armz['Data_Armz'].notna()) & ((df_armz['Data_Armz'] < data_sel) | ((df_armz['Data_Armz'] == data_sel) & (df_armz['Hora_Armz'] <= hora)))]['NU_ETIQUETA'].nunique()
                pendencias = max(0, entrou - saiu)
                
            dados_grafico.append({'Hora': hora, 'Armazenados': armz_hora, 'Conferidos': conf_hora, 'Pendências': pendencias})
            
        df_fluxo = pd.DataFrame(dados_grafico)
        if not df_fluxo.empty:
            st.markdown('<div class="MAGALOG-card">', unsafe_allow_html=True)
            fig_fluxo = go.Figure()
            fig_fluxo.add_trace(go.Bar(x=df_fluxo['Hora'], y=df_fluxo['Armazenados'], name='Armazenados', marker_color='#0086FF', text=df_fluxo['Armazenados'], textposition='auto'))
            fig_fluxo.add_trace(go.Bar(x=df_fluxo['Hora'], y=df_fluxo['Conferidos'], name='Conferidos', marker_color='#CBD5E1', text=df_fluxo['Conferidos'], textposition='outside'))
            fig_fluxo.add_trace(go.Scatter(x=df_fluxo['Hora'], y=df_fluxo['Pendências'], name='Fila Acumulada', mode='lines+markers+text', line=dict(color='#FF3366', width=3), yaxis='y2', text=df_fluxo['Pendências'], textposition='top center'))
            
            fig_fluxo.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', 
                paper_bgcolor='rgba(0,0,0,0)',
                font_family="Inter",
                barmode='group', 
                legend=dict(orientation="h", y=1.15, x=0.5, xanchor='center'), 
                yaxis2=dict(overlaying='y', side='right', showgrid=False), 
                hovermode="x unified",
                xaxis=dict(showgrid=False),
                yaxis=dict(gridcolor='#F1F5F9')
            )
            
            ev = st.plotly_chart(fig_fluxo, use_container_width=True, on_select="rerun")
            st.markdown('</div>', unsafe_allow_html=True)
            
            if hora_manual != "Selecione...": popup_detalhe_hora(hora_manual, df_armz, data_sel)
            elif isinstance(ev, dict) and "selection" in ev and ev["selection"].get("points"):
                popup_detalhe_hora(ev["selection"]["points"][0].get("x"), df_armz, data_sel)
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<h4 style='color: #334155; margin-bottom: 15px;'><span class='icon-MAGALOG'>social_leaderboard</span> Placar de Produtividade: Operadores</h4>", unsafe_allow_html=True)
        
        if not df_producao_real.empty:
            df_cadencia = df_producao_real.sort_values(['OPERADOR', 'DT_ARMAZENAGEM_CALC']).copy()
            df_cadencia['DIFF_MINS'] = df_cadencia.groupby('OPERADOR')['DT_ARMAZENAGEM_CALC'].diff().dt.total_seconds() / 60.0
            
            df_cadencia_limpo = df_cadencia[(df_cadencia['DIFF_MINS'] > 0) & (df_cadencia['DIFF_MINS'] < 45)]
            
            rank_op = df_producao_real.groupby('OPERADOR').agg(
                Etiquetas_Armazenadas=('NU_ETIQUETA', 'nunique'),
                Horas_Trabalhadas=('Hora_Armz', 'nunique'),
                SLA_Medio_Doca=('Tempo_Espera_Minutos', 'mean')
            ).reset_index()
            
            media_cadencia = df_cadencia_limpo.groupby('OPERADOR')['DIFF_MINS'].mean().reset_index()
            media_cadencia.columns = ['OPERADOR', 'Intervalo_Medio']
            
            rank_op = pd.merge(rank_op, media_cadencia, on='OPERADOR', how='left')
            
            rank_op['Média (Etq/Hora)'] = (rank_op['Etiquetas_Armazenadas'] / rank_op['Horas_Trabalhadas'].replace(0, 1)).round(1)
            rank_op['Tempo Médio Doca'] = rank_op['SLA_Medio_Doca'].apply(mins_to_text)
            rank_op['Ritmo_Bipagem'] = rank_op['Intervalo_Medio'].apply(lambda x: f"{x:.1f} min" if pd.notna(x) else "-")
            
            rank_op = rank_op.sort_values(['Etiquetas_Armazenadas', 'Média (Etq/Hora)'], ascending=[False, False]).reset_index(drop=True)
            
            st.markdown("""
            <style>
            .lb-wrapper { background: rgba(255, 255, 255, 0.85); border: 1px solid rgba(255, 255, 255, 0.4); box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.07); backdrop-filter: blur(8px); border-radius: 16px; padding: 25px; margin-top: 15px;}
            .lb-header-op { display: grid; grid-template-columns: 0.5fr 2fr 1fr 1fr 1.2fr 1.2fr; color: #64748B; font-weight: 800; font-size: 10px; text-transform: uppercase; padding: 0 15px 12px 15px; border-bottom: 2px solid #F1F5F9; margin-bottom: 12px; align-items: center;}
            .lb-row-op { display: grid; grid-template-columns: 0.5fr 2fr 1fr 1fr 1.2fr 1.2fr; align-items: center; background: #FFFFFF; margin-bottom: 8px; padding: 12px 15px; border-radius: 12px; border: 1px solid #E2E8F0; transition: all 0.2s; border-left: 6px solid transparent;}
            .lb-row-op:hover { transform: translateX(4px); box-shadow: 0 5px 15px rgba(0,0,0,0.05); }
            .lb-gold { background: linear-gradient(90deg, #FFFBEB 0%, #FFFFFF 30%); border-left-color: #F59E0B; border-color: #FEF3C7;}
            .lb-silver { background: linear-gradient(90deg, #F8FAFC 0%, #FFFFFF 30%); border-left-color: #94A3B8; border-color: #E2E8F0;}
            .lb-bronze { background: linear-gradient(90deg, #FFF7ED 0%, #FFFFFF 30%); border-left-color: #B45309; border-color: #FFEDD5;}
            .lb-danger { background: linear-gradient(90deg, #FEF2F2 0%, #FFFFFF 30%); border-left-color: #EF4444; border-color: #FEE2E2;}
            .lb-rank { font-size: 18px; font-weight: 900; color: #0F172A; }
            .lb-name { font-size: 13px; font-weight: 800; color: #1E293B; display: flex; align-items: center; gap: 8px;}
            .lb-stat { font-size: 13px; font-weight: 600; color: #475569; }
            .lb-highlight { background: #F0F9FF; color: #0284C7; padding: 4px 8px; border-radius: 6px; font-weight: 800; border: 1px solid #BAE6FD; display: inline-block; font-size: 11px;}
            .lb-tooltip { position: relative; display: inline-block; cursor: help; color: #94A3B8; margin-left: 4px; vertical-align: middle;}
            .lb-tooltip .lb-tooltiptext { visibility: hidden; width: 200px; background-color: #0F172A; color: #fff; text-align: center; border-radius: 6px; padding: 8px; position: absolute; z-index: 1; bottom: 125%; left: 50%; margin-left: -100px; opacity: 0; transition: opacity 0.3s; font-size: 11px; font-weight: 400; text-transform: none;}
            .lb-tooltip:hover .lb-tooltiptext { visibility: visible; opacity: 1; }
            </style>
            """, unsafe_allow_html=True)

            html_lb = "<div class='lb-wrapper'>"
            html_lb += "<div class='lb-header-op'>"
            html_lb += "<div>POS</div><div>OPERADOR</div>"
            html_lb += "<div>GUARDADAS</div>"
            html_lb += "<div>VELOC.</div>"
            html_lb += "<div>SLA DOCA</div>"
            html_lb += "<div>Intervalo entre Armazenagens <div class='lb-tooltip'><span class='icon-MAGALOG' style='font-size:14px;'>help</span><div class='lb-tooltiptext'>Tempo médio de espera entre uma armazenagem e outra.</div></div></div>"
            html_lb += "</div>"
            
            total_ops = len(rank_op)
            for idx, row_r in rank_op.iterrows():
                pos_r = idx + 1
                css_c = "lb-gold" if pos_r == 1 else "lb-silver" if pos_r == 2 else "lb-bronze" if pos_r == 3 else "lb-danger" if (pos_r >= total_ops - 1 and total_ops >= 5) else ""
                ic_r = "workspace_premium" if pos_r == 1 else "military_tech" if pos_r <= 3 else "warning" if "danger" in css_c else "person"
                
                cor_ritmo = "color: #DC2626;" if (pd.notna(row_r['Intervalo_Medio']) and row_r['Intervalo_Medio'] > 10) else ""
                html_lb += f"<div class='lb-row-op {css_c}'>"
                html_lb += f"<div class='lb-rank'>{pos_r}º</div>"
                html_lb += f"<div class='lb-name'><span class='icon-MAGALOG' style='font-size:20px;'>{ic_r}</span> {row_r['OPERADOR']}</div>"
                html_lb += f"<div class='lb-stat'><span style='font-weight: 800;'>{int(row_r['Etiquetas_Armazenadas'])}</span></div>"
                html_lb += f"<div class='lb-stat'><span class='lb-highlight'>{row_r['Média (Etq/Hora)']} etq/h</span></div>"
                html_lb += f"<div class='lb-stat'>{row_r['Tempo Médio Doca']}</div>"
                html_lb += f"<div class='lb-stat' style='font-weight: 700; {cor_ritmo}'>{row_r['Ritmo_Bipagem']}</div>"
                html_lb += "</div>"
            
            html_lb += "</div>"
            st.markdown(html_lb, unsafe_allow_html=True)
        else:
            st.info("Nenhuma armazenagem registrada para os filtros selecionados.")

# -------------------------------------------------------------------------
# ABA 2: CONFERÊNCIA (METAS PREDITIVAS E AUTO-SAVE BLINDADO)
# -------------------------------------------------------------------------
with tab2:
    if not df_hist_conf.empty and not df_hoje_conf.empty:
        st.markdown('<div class="MAGALOG-ribbon">Inteligência Preditiva</div>', unsafe_allow_html=True)
        st.caption("O algoritmo localiza cargas irmãs no histórico para gerar a meta mais justa possível, considerando também tempo de setup.")
        
        def calcular_meta_inteligente(row, df_historico):
            forn = str(row.get('ORIGEM', '')).strip().upper()
            linha = str(row.get('CATEGORIA', '')).strip().upper()
            pecas = row.get('PEÇAS', 0)
            sku = row.get('SKU', 0)
            
            TEMPO_SETUP = 15
            
            min_pecas, max_pecas = pecas * 0.7, pecas * 1.3
            min_sku, max_sku = min(sku * 0.7, sku - 2), max(sku * 1.3, sku + 2)
            df_hist_limpo = df_historico[(df_historico['TMP APC'] > 5) & (df_historico['PEÇAS'] > 0)].copy()
            df_hist_limpo['VELOCIDADE'] = df_hist_limpo['TMP APC'] / df_hist_limpo['PEÇAS']
            df_hist_limpo = df_hist_limpo[df_hist_limpo['VELOCIDADE'] >= 0.05] 
            taxa_global_mediana = df_hist_limpo['VELOCIDADE'].median()
            
            if pd.isna(taxa_global_mediana): taxa_global_mediana = 0.5 
            df_base_exata = df_hist_limpo[(df_hist_limpo['FORNECEDOR'].str.upper() == forn) & (df_hist_limpo['LINHA'].str.upper() == linha)]
            
            if not df_base_exata.empty:
                df_gemeas = df_base_exata[(df_base_exata['PEÇAS'] >= min_pecas) & (df_base_exata['PEÇAS'] <= max_pecas) & (df_base_exata['SKU'] >= min_sku) & (df_base_exata['SKU'] <= max_sku)]
                if not df_gemeas.empty: return df_gemeas['TMP APC'].median() 
                
                df_primas = df_base_exata[(df_base_exata['PEÇAS'] >= min_pecas) & (df_base_exata['PEÇAS'] <= max_pecas)]
                if not df_primas.empty: return df_primas['TMP APC'].median()
                
                vel_mediana = df_base_exata['VELOCIDADE'].median()
                return TEMPO_SETUP + (pecas * vel_mediana)
                
            df_base_categoria = df_hist_limpo[df_hist_limpo['LINHA'].str.upper() == linha]
            if not df_base_categoria.empty:
                df_gemeas_cat = df_base_categoria[(df_base_categoria['PEÇAS'] >= min_pecas) & (df_base_categoria['PEÇAS'] <= max_pecas) & (df_base_categoria['SKU'] >= min_sku) & (df_base_categoria['SKU'] <= max_sku)]
                if not df_gemeas_cat.empty: return df_gemeas_cat['TMP APC'].median()
                
                df_primas_cat = df_base_categoria[(df_base_categoria['PEÇAS'] >= min_pecas) & (df_base_categoria['PEÇAS'] <= max_pecas)]
                if not df_primas_cat.empty: return df_primas_cat['TMP APC'].median()
                
                vel_mediana_cat = df_base_categoria['VELOCIDADE'].median()
                return TEMPO_SETUP + (pecas * vel_mediana_cat)
                
            return TEMPO_SETUP + (pecas * taxa_global_mediana)

        df_hoje_conf['DURAÇÃO_REAL_MIN'] = df_hoje_conf['DURAÇÃO CARGA'].apply(time_to_mins)
        df_hoje_conf['STATUS_FISICO'] = df_hoje_conf['STATUS_FISICO'].str.strip().str.upper()
        df_hoje_conf['META_TEMPO_MIN'] = df_hoje_conf.apply(lambda row: calcular_meta_inteligente(row, df_hist_conf), axis=1)
        
        agora = pd.Timestamp.now(tz='America/Sao_Paulo')
        
        # CORREÇÃO 2: Limpeza dos Emojis nos DataFrames e Lógica
        def calcular_previsao(row):
            status = row['STATUS_FISICO']
            if status == 'OK': return "FINALIZADO"
            restante = row['META_TEMPO_MIN'] - row['DURAÇÃO_REAL_MIN']
            if restante < 0: return "ESTOUROU"
            return (agora + pd.Timedelta(minutes=restante)).strftime("%H:%M")
            
        def calcular_situacao_meta(row):
            status = row['STATUS_FISICO']
            if status == 'OK': return "NO PRAZO" if row['DURAÇÃO_REAL_MIN'] <= row['META_TEMPO_MIN'] else "ATRASADO (FIN)"
            else:
                if row['DURAÇÃO_REAL_MIN'] > row['META_TEMPO_MIN']: return "ESTOURADO"
                elif status == 'EM PROCESSO': return "NO RITMO"
                else: return "AGUARDANDO"
            
        df_hoje_conf['PREVISÃO FIM'] = df_hoje_conf.apply(calcular_previsao, axis=1)
        df_hoje_conf['SITUAÇÃO META'] = df_hoje_conf.apply(calcular_situacao_meta, axis=1)
        
        c1, c2, c3, c4 = st.columns(4)
        cargas_totais = len(df_hoje_conf)
        cargas_ok = df_hoje_conf[df_hoje_conf['STATUS_FISICO'] == 'OK'].shape[0]
        cargas_fila = df_hoje_conf[df_hoje_conf['STATUS_FISICO'].isin(['EM DOCA', 'P-EXTERNO'])].shape[0]
        acertos = df_hoje_conf[df_hoje_conf['SITUAÇÃO META'].isin(['NO PRAZO', 'NO RITMO', 'AGUARDANDO'])].shape[0]
        perc_acerto = (acertos / cargas_totais) * 100 if cargas_totais > 0 else 0
        
        with c1: exibir_kpi("Agendas do Dia", cargas_totais, "Na grade", "#8B5CF6")
        with c2: exibir_kpi("Finalizadas", cargas_ok, "Cargas Entregues", "#0086FF")
        with c3: exibir_kpi("Fila Física", cargas_fila, "Doca ou Pátio", "#F59E0B")
        with c4: exibir_kpi("Saúde das Metas", f"{perc_acerto:.1f}%", "Aderência", "#10B981" if perc_acerto > 80 else "#EF4444")
        
        st.markdown("<h4 style='color: #334155; margin-bottom: 15px; margin-top: 25px;'><span class='icon-MAGALOG'>rocket_launch</span> Despacho de Cargas e Previsão Algorítmica</h4>", unsafe_allow_html=True)
        
        df_tabela = df_hoje_conf[['AGENDA', 'CONFERENTE', 'CATEGORIA', 'STATUS_FISICO', 'PEÇAS', 'SKU', 'META_TEMPO_MIN', 'DURAÇÃO_REAL_MIN', 'PREVISÃO FIM', 'SITUAÇÃO META']].copy()
        df_tabela['META (Tempo)'] = df_tabela['META_TEMPO_MIN'].apply(mins_to_text)
        df_tabela['GASTO (Tempo)'] = df_tabela['DURAÇÃO_REAL_MIN'].apply(mins_to_text)
        df_tabela['PEÇAS'] = df_tabela['PEÇAS'].apply(lambda x: f"{int(x):,}".replace(",", "."))
        df_tabela['SKU'] = df_tabela['SKU'].apply(lambda x: f"{int(x)}")
        df_tabela = df_tabela[['AGENDA', 'CONFERENTE', 'CATEGORIA', 'STATUS_FISICO', 'PEÇAS', 'SKU', 'META (Tempo)', 'GASTO (Tempo)', 'PREVISÃO FIM', 'SITUAÇÃO META']]
        
        def estilizar_tabela(df):
            estilos = pd.DataFrame('', index=df.index, columns=df.columns)
            
            # Formatação baseada na string exata (Sem Emoji)
            cond_verde_meta = df['SITUAÇÃO META'].astype(str).str.contains('NO PRAZO')
            cond_verm_meta = df['SITUAÇÃO META'].astype(str).str.contains('ATRASADO|ESTOURADO')
            cond_amar_meta = df['SITUAÇÃO META'].astype(str).str.contains('NO RITMO')
            
            estilos.loc[cond_verde_meta, 'SITUAÇÃO META'] = 'color: #065F46; background-color: #D1FAE5; font-weight: 600;'
            estilos.loc[cond_verm_meta, 'SITUAÇÃO META'] = 'color: #991B1B; background-color: #FEE2E2; font-weight: 600;'
            estilos.loc[cond_amar_meta, 'SITUAÇÃO META'] = 'color: #92400E; background-color: #FEF3C7; font-weight: 600;'
            
            cond_verde_prev = df['PREVISÃO FIM'].astype(str).str.contains('FINALIZADO')
            cond_verm_prev = df['PREVISÃO FIM'].astype(str).str.contains('ESTOUROU')
            cond_amar_prev = df['PREVISÃO FIM'].astype(str).str.contains(':') # Quando é horário HH:MM
            
            estilos.loc[cond_verde_prev, 'PREVISÃO FIM'] = 'color: #065F46; background-color: #D1FAE5; font-weight: 600;'
            estilos.loc[cond_verm_prev, 'PREVISÃO FIM'] = 'color: #991B1B; background-color: #FEE2E2; font-weight: 600;'
            estilos.loc[cond_amar_prev, 'PREVISÃO FIM'] = 'color: #92400E; background-color: #FEF3C7; font-weight: 600;'
            
            return estilos

        st.dataframe(df_tabela.style.apply(estilizar_tabela, axis=None), use_container_width=True, hide_index=True)

        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("<div class='MAGALOG-card' style='border-left: 4px solid #10B981 !important;'>", unsafe_allow_html=True)
        st.markdown("<h4 style='color: #0F172A; margin-top:0;'><span class='icon-MAGALOG' style='color:#10B981;'>sync</span> Sincronização Contínua (Anti-F5)</h4>", unsafe_allow_html=True)
        
        data_hoje_str = agora.strftime('%d/%m/%Y')
        
        df_hoje_ok = df_hoje_conf[(df_hoje_conf['STATUS_FISICO'] == 'OK') & (df_hoje_conf['DURAÇÃO_REAL_MIN'] > 0)].copy()
        df_hoje_ok = df_hoje_ok.drop_duplicates(subset=['AGENDA'])
        
        agendas_no_cofre = []
        if not df_fechamento.empty and 'AGENDA' in df_fechamento.columns:
            agendas_no_cofre = df_fechamento['AGENDA'].astype(str).tolist()
            
        df_para_salvar = df_hoje_ok[~df_hoje_ok['AGENDA'].astype(str).isin(agendas_no_cofre)].copy()
        
        # CORREÇÃO 3: Troca do st.metric() feio com emoji por Cards HTML customizados
        c_sync1, c_sync2 = st.columns(2)
        with c_sync1:
            st.markdown(f"""
            <div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; padding: 15px;">
                <div style="font-size: 11px; color: #64748B; font-weight: 800; text-transform: uppercase;"><span class='icon-MAGALOG' style='vertical-align: middle; font-size: 16px; margin-right: 4px;'>inventory_2</span> Cargas Registradas (Cofre)</div>
                <div style="font-size: 24px; font-weight: 900; color: #0F172A; margin-top: 4px;">{len(agendas_no_cofre)}</div>
            </div>
            """, unsafe_allow_html=True)
            
        with c_sync2:
            st.markdown(f"""
            <div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; padding: 15px;">
                <div style="font-size: 11px; color: #64748B; font-weight: 800; text-transform: uppercase;"><span class='icon-MAGALOG' style='vertical-align: middle; font-size: 16px; margin-right: 4px; color: #F59E0B;'>notification_important</span> Novas Cargas na Fila</div>
                <div style="font-size: 24px; font-weight: 900; color: #F59E0B; margin-top: 4px;">{len(df_para_salvar)}</div>
            </div>
            """, unsafe_allow_html=True)
        
        if not df_para_salvar.empty:
            st.info(f"Foram encontradas {len(df_para_salvar)} novas cargas finalizadas! Salvando no cofre automaticamente...")
            
            df_export = pd.DataFrame({
                'DATA': data_hoje_str, 'AGENDA': df_para_salvar['AGENDA'], 'CONFERENTE': df_para_salvar['CONFERENTE'],
                'CATEGORIA': df_para_salvar['CATEGORIA'], 'PEÇAS': df_para_salvar['PEÇAS'],
                'META MINUTOS': df_para_salvar['META_TEMPO_MIN'].round(2), 'REALIZADO MINUTOS': df_para_salvar['DURAÇÃO_REAL_MIN'].round(2),
                'RESULTADO': df_para_salvar['SITUAÇÃO META'].apply(lambda x: 'NO PRAZO' if 'NO PRAZO' in x else 'ATRASADO')
            })
            
            with st.spinner("Sincronizando Banco de Dados..."):
                sucesso = salvar_historico_fechamento(df_export)
                if sucesso:
                    st.success("Novas cargas sincronizadas com sucesso!")
                    st.cache_data.clear() 
                    st.rerun() 
        else:
            st.success("O Cofre está 100% sincronizado. Nenhuma carga nova pendente de gravação.")
            
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='MAGALOG-card'><h4 style='color: #DC2626;'><span class='icon-MAGALOG'>warning</span> Planilhas de Conferência desconectadas ou vazias</h4><p style='color: #64748B; font-size: 13px;'>Verifique o arquivo fonte no Google Sheets.</p></div>", unsafe_allow_html=True)

# -------------------------------------------------------------------------
# ABA 3: DESEMPENHO DA EQUIPE
# -------------------------------------------------------------------------
    with tab3:
        st.caption("Acompanhamento histórico de performance, velocidade e aderência às metas da equipe.")
        
        if not df_fechamento.empty:
            datas_disponiveis = sorted(df_fechamento['DATA'].unique(), reverse=True)
            data_hist_sel = st.multiselect("Filtrar Período:", options=datas_disponiveis, default=datas_disponiveis)
            
            df_f = df_fechamento[df_fechamento['DATA'].isin(data_hist_sel)].copy()
            
            if not df_f.empty:
                df_f['Desvio (Minutos)'] = df_f['REALIZADO MINUTOS'] - df_f['META MINUTOS']
                df_f['STATUS_REAL'] = df_f['Desvio (Minutos)'].apply(lambda x: 'ATRASADO' if x > 0 else 'NO PRAZO')
                
                ranking = df_f.groupby('CONFERENTE').agg(
                    Cargas_Feitas=('AGENDA', 'count'),
                    Atrasos=('STATUS_REAL', lambda x: (x == 'ATRASADO').sum()),
                    No_Prazo=('STATUS_REAL', lambda x: (x == 'NO PRAZO').sum()),
                    Tempo_Medio_Desvio=('Desvio (Minutos)', 'mean')
                ).reset_index()
                
                ranking['% de Acerto'] = (ranking['No_Prazo'] / ranking['Cargas_Feitas']) * 100
                ranking = ranking.sort_values('% de Acerto', ascending=False)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("<h4 style='color: #334155; margin-bottom: 15px; margin-top: 15px;'><span class='icon-MAGALOG'>military_tech</span> Top Performers (Aderência)</h4>", unsafe_allow_html=True)
                    st.markdown('<div class="MAGALOG-card">', unsafe_allow_html=True)
                    fig_bar = px.bar(ranking, x='CONFERENTE', y='% de Acerto', text_auto='.1f', 
                                     color='% de Acerto', color_continuous_scale='Blues',
                                     labels={'% de Acerto': 'Taxa de Acerto (%)'})
                    fig_bar.update_layout(yaxis=dict(range=[0, 100]), coloraxis_showscale=False, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(l=0, r=0, t=10, b=0))
                    st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                with col2:
                    st.markdown("<h4 style='color: #334155; margin-bottom: 15px; margin-top: 15px;'><span class='icon-MAGALOG'>balance</span> Balanço de Tempo (Gargalo vs Ganho)</h4>", unsafe_allow_html=True)
                    st.markdown('<div class="MAGALOG-card">', unsafe_allow_html=True)
                    ranking_desvio = ranking.sort_values('Tempo_Medio_Desvio', ascending=False)
                    cores = ['#FF3366' if val > 0 else '#00C853' for val in ranking_desvio['Tempo_Medio_Desvio']]
                    
                    fig_desv = go.Figure(go.Bar(
                        x=ranking_desvio['CONFERENTE'], y=ranking_desvio['Tempo_Medio_Desvio'], 
                        marker_color=cores, text=ranking_desvio['Tempo_Medio_Desvio'].round(1), textposition='auto'
                    ))
                    fig_desv.update_layout(yaxis_title="Minutos (Média de Desvio)", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(l=0, r=0, t=10, b=0))
                    st.plotly_chart(fig_desv, use_container_width=True, config={'displayModeBar': False})
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                st.markdown("<br><h4 style='color: #334155; margin-bottom: 15px;'><span class='icon-MAGALOG'>analytics</span> Detalhamento Analítico Geral</h4>", unsafe_allow_html=True)
                st.dataframe(ranking[['CONFERENTE', 'Cargas_Feitas', 'No_Prazo', 'Atrasos', '% de Acerto', 'Tempo_Medio_Desvio']].style.format({
                    '% de Acerto': '{:.1f}%',
                    'Tempo_Medio_Desvio': '{:.1f} min'
                }), use_container_width=True, hide_index=True)
                
                st.markdown("---")
                
                st.markdown("<h4 style='color: #334155; margin-bottom: 15px;'><span class='icon-MAGALOG'>person_search</span> Investigador de Conferente</h4>", unsafe_allow_html=True)
                
                lista_conferentes = ["Selecione um Conferente..."] + sorted(df_f['CONFERENTE'].unique())
                conferente_alvo = st.selectbox("Escolha quem você quer investigar:", lista_conferentes)
                
                if conferente_alvo != "Selecione um Conferente...":
                    df_individual = df_f[df_f['CONFERENTE'] == conferente_alvo].copy()
                    
                    c_k1, c_k2, c_k3 = st.columns(3)
                    qtd_ok = df_individual[df_individual['STATUS_REAL'] == 'NO PRAZO'].shape[0]
                    qtd_bad = df_individual[df_individual['STATUS_REAL'] == 'ATRASADO'].shape[0]
                    saldo_total_min = df_individual['Desvio (Minutos)'].sum()
                    
                    with c_k1: exibir_kpi("Cargas no Prazo", qtd_ok, "", "#00C853")
                    with c_k2: exibir_kpi("Cargas Estouradas", qtd_bad, "", "#FF3366")
                    with c_k3: exibir_kpi("Balanço Total (Minutos)", f"{saldo_total_min:.1f} min", "", "#0086FF")
                    
                    st.markdown(f"<div style='font-weight: 800; color: #0F172A; margin-bottom: 10px;'>Histórico de agendas de {conferente_alvo}</div>", unsafe_allow_html=True)
                    
                    df_detalhe = df_individual[['DATA', 'AGENDA', 'CATEGORIA', 'PEÇAS', 'META MINUTOS', 'REALIZADO MINUTOS', 'Desvio (Minutos)', 'STATUS_REAL']].copy()
                    df_detalhe['META (Tempo)'] = df_detalhe['META MINUTOS'].apply(mins_to_text)
                    df_detalhe['REAL (Tempo)'] = df_detalhe['REALIZADO MINUTOS'].apply(mins_to_text)
                    df_detalhe['Desvio (Minutos)'] = df_detalhe['Desvio (Minutos)'].round(1)
                    
                    df_detalhe = df_detalhe[['DATA', 'AGENDA', 'CATEGORIA', 'PEÇAS', 'META (Tempo)', 'REAL (Tempo)', 'Desvio (Minutos)', 'STATUS_REAL']]
                    
                    def estilizar_tabela_indiv(df):
                        estilos = pd.DataFrame('', index=df.index, columns=df.columns)
                        
                        cond_verde = df['STATUS_REAL'].astype(str).str.contains('NO PRAZO')
                        cond_verm = df['STATUS_REAL'].astype(str).str.contains('ATRASADO')
                        
                        estilos.loc[cond_verde, 'STATUS_REAL'] = 'color: #065F46; background-color: #D1FAE5; font-weight: 600;'
                        estilos.loc[cond_verm, 'STATUS_REAL'] = 'color: #991B1B; background-color: #FEE2E2; font-weight: 600;'
                        
                        return estilos

                    st.dataframe(df_detalhe.style.apply(estilizar_tabela_indiv, axis=None), use_container_width=True, hide_index=True)
            else:
                st.info("Nenhuma data selecionada.")
        else:
            st.info("Banco de Fechamento Vazio. Os resultados aparecerão após a primeira gravação diária.")
