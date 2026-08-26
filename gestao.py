import io
import os
import smtplib
import urllib.parse
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, date, timedelta
import zoneinfo
import importlib
import ferias
import pandas as pd
import plotly.express as px
import streamlit as st
import gspread
from google import genai

importlib.reload(ferias)

# --- FUSO HORÁRIO OFICIAL BRASÍLIA ---
FUSO_SP = zoneinfo.ZoneInfo("America/Sao_Paulo")

def obter_hoje_brasilia():
    return datetime.now(FUSO_SP).date()

# --- CARREGAMENTO DE DADOS COM CACHE CURTO ---
@st.cache_data(ttl=5)
def carregar_dados():
    try:
        url_sheets = st.secrets.get("GSHEETS_URL", "")
        if url_sheets:
            gc = obter_cliente_gspread()
            sh = gc.open_by_url(url_sheets)
            worksheet = sh.worksheet("equipe")
            dados = worksheet.get_all_records()
            df = pd.DataFrame(dados)
        else:
            df = pd.read_excel("equipe.xlsx") if os.path.exists("equipe.xlsx") else pd.DataFrame()
        
        if not df.empty:
            df.columns = df.columns.str.strip()
            col_adm = next((c for c in df.columns if 'admiss' in str(c).lower() or 'dt_adm' in str(c).lower()), 'Admissão')
            col_nasc = next((c for c in df.columns if 'nasc' in str(c).lower() or 'anivers' in str(c).lower()), 'Nascimento')
            
            df['dt_adm'] = pd.to_datetime(df[col_adm], dayfirst=True, errors='coerce').dt.date if col_adm in df.columns else None
            if col_nasc in df.columns:
                df['dt_nasc_dt'] = pd.to_datetime(df[col_nasc], dayfirst=True, errors='coerce')
                df['dt_nasc'] = df['dt_nasc_dt'].dt.date
            else:
                df['dt_nasc_dt'] = pd.NaT
                df['dt_nasc'] = None

            if 'Vaga' in df.columns:
                df['Vaga'] = df['Vaga'].astype(str).str.replace('.0', '', regex=False)
            if 'Matricula' in df.columns:
                df['Matricula'] = df['Matricula'].astype(str).str.replace('.0', '', regex=False)
            if 'Ultimas_Ferias' not in df.columns:
                df['Ultimas_Ferias'] = None
            else:
                df['Ultimas_Ferias'] = df['Ultimas_Ferias'].astype(str)
                df['dt_ult_ferias'] = pd.to_datetime(df['Ultimas_Ferias'], dayfirst=True, errors='coerce').dt.date
            if 'Decisao_Experiencia' not in df.columns:
                df['Decisao_Experiencia'] = None
            if 'Status' not in df.columns:
                df['Status'] = 'Ativo'
            else:
                df['Status'] = df['Status'].fillna('Ativo').astype(str).str.strip()
        return df
    except Exception:
        return pd.read_excel("equipe.xlsx") if os.path.exists("equipe.xlsx") else pd.DataFrame()

@st.cache_data(ttl=5)
def carregar_faltas():
    cols_padrao = ["Matricula", "Funcionário", "Setor", "Data", "Tipo", "Dias", "CID", "Motivo", "dt_falta"]
    try:
        url_sheets = st.secrets.get("GSHEETS_URL", "")
        if url_sheets:
            gc = obter_cliente_gspread()
            sh = gc.open_by_url(url_sheets)
            worksheet = sh.worksheet("faltas")
            dados = worksheet.get_all_records()
            df_f = pd.DataFrame(dados)
        else:
            df_f = pd.read_excel("faltas.xlsx") if os.path.exists("faltas.xlsx") else pd.DataFrame(columns=cols_padrao)

        if not df_f.empty:
            df_f.columns = df_f.columns.str.strip()
            for col in cols_padrao:
                if col not in df_f.columns and col != 'dt_falta':
                    df_f[col] = ""
            df_f['Data'] = df_f['Data'].astype(str).str.strip()
            df_f['dt_falta'] = pd.to_datetime(df_f['Data'], format='%d/%m/%Y', errors='coerce').dt.date
        else:
            df_f = pd.DataFrame(columns=cols_padrao)
        return df_f
    except Exception:
        return pd.DataFrame(columns=cols_padrao)

# --- FUNÇÕES DE SALVAMENTO COM LIMPEZA EXPLICITA DE CACHE ---
def salvar_dados(df_salvar):
    cols_salvar = [c for c in df_salvar.columns if c not in ['dt_adm', 'dt_nasc', 'dt_nasc_dt', 'dt_ult_ferias', 'exp_45', 'exp_90', 'dias_para_45', 'dias_para_90']]
    url_sheets = st.secrets.get("GSHEETS_URL", "")
    if url_sheets:
        try:
            gc = obter_cliente_gspread()
            sh = gc.open_by_url(url_sheets)
            ws = sh.worksheet("equipe")
            ws.clear()
            ws.update([df_salvar[cols_salvar].columns.values.tolist()] + df_salvar[cols_salvar].fillna("").values.tolist())
        except Exception as e:
            st.error(f"Erro ao salvar equipe no Sheets: {e}")
    df_salvar[cols_salvar].to_excel("equipe.xlsx", index=False)
    st.cache_data.clear()

def salvar_faltas(df_f):
    cols_salvar = [c for c in df_f.columns if c != 'dt_falta']
    url_sheets = st.secrets.get("GSHEETS_URL", "")
    if url_sheets:
        try:
            gc = obter_cliente_gspread()
            sh = gc.open_by_url(url_sheets)
            ws = sh.worksheet("faltas")
            ws.clear()
            ws.update([df_f[cols_salvar].columns.values.tolist()] + df_f[cols_salvar].fillna("").values.tolist())
        except Exception as e:
            st.error(f"Erro ao salvar faltas no Sheets: {e}")
    df_f[cols_salvar].to_excel("faltas.xlsx", index=False)
    st.cache_data.clear()
