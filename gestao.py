import io
import os
import smtplib
import urllib.parse
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, date, timedelta
import importlib
import ferias
import pandas as pd
import plotly.express as px
import streamlit as st
import google.generativeai as genai
from supabase import create_client, Client

importlib.reload(ferias)

st.set_page_config(
    page_title="Painel de Gestão & DP — Tropical", 
    page_icon="🍊", 
    layout="wide"
)

# --- CONEXÃO COM O SUPABASE E CONFIGURAÇÃO DA IA ---
try:
    gemini_api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))
    if gemini_api_key:
        genai.configure(api_key=gemini_api_key)
        modelo_ia = genai.GenerativeModel('gemini-1.5-pro')
        ia_disponivel = True
    else:
        ia_disponivel = False
except Exception:
    ia_disponivel = False

try:
    supabase_url = st.secrets.get("SUPABASE_URL", "")
    supabase_key = st.secrets.get("SUPABASE_KEY", "")
    if supabase_url and supabase_key:
        supabase: Client = create_client(supabase_url, supabase_key)
        supabase_disponivel = True
    else:
        supabase_disponivel = False
except Exception:
    supabase_disponivel = False

# --- ESTILOS CSS PERSONALIZADOS (Com botões Senha e Sair Laranjas) ---
URL_LOGO_TROPICAL = "https://cdn-icons-png.flaticon.com/512/1625/1625048.png"

st.markdown(
    f"""
    <meta name="apple-mobile-web-app-title" content="Painel Gestão & DP">
    <meta name="application-name" content="Painel Gestão & DP">
    <link rel="apple-touch-icon" href="{URL_LOGO_TROPICAL}">
    <link rel="icon" type="image/png" href="{URL_LOGO_TROPICAL}">
    
    <style>
        .stApp {{
            background-color: #0E1117;
            color: #FFFFFF;
        }}
        [data-testid="stSidebar"] {{
            background-color: #1B3B2B !important;
        }}
        [data-testid="stSidebar"] * {{
            color: #FFFFFF !important;
        }}
        div.stButton > button {{
            background-color: #FF6B00 !important;
            color: #FFFFFF !important;
            border-radius: 12px !important;
            border: none !important;
            font-weight: bold !important;
            padding: 10px 16px !important;
            box-shadow: 0 4px 10px rgba(255, 107, 0, 0.3) !important;
            transition: all 0.3s ease !important;
        }}
        div.stButton > button:hover {{
            background-color: #E05E00 !important;
            transform: translateY(-2px);
        }}
        div.stDownloadButton > button {{
            background-color: #1E293B !important;
            color: #FF6B00 !important;
            border: 2px solid #FF6B00 !important;
            border-radius: 12px !important;
            font-weight: bold !important;
        }}
        [data-testid="stMetricValue"] {{
            color: #FF6B00 !important;
            font-size: 2rem !important;
            font-weight: bold !important;
        }}
        button[data-baseweb="tab"] {{
            color: #94A3B8 !important;
            font-weight: bold !important;
        }}
        button[aria-selected="true"] {{
            color: #FF6B00 !important;
            border-bottom-color: #FF6B00 !important;
        }}
    </style>
    """,
    unsafe_allow_html=True
)

ARQUIVO_DADOS = "equipe.xlsx"
ARQUIVO_FALTAS = "faltas.xlsx"
ARQUIVO_USUARIOS = "usuarios.xlsx"
ARQUIVO_EPIS = "epis.xlsx"
ARQUIVO_HISTORICO = "historico_colaboradores.xlsx"

TODOS_MODULOS = [
    "Dashboard & Alertas",
    "🤖 Assistente IA (DP & Gestão)",
    "Chamada & Faltas do Dia",
    "🦺 Solicitação & Entrega de EPI",
    "👤 Ficha Individual do Colaborador",
    "Controle de Experiência (45/90 dias)",
    "Escala Inteligente de Férias",
    "🏖️ Colaboradores em Férias",
    "📊 Indicadores de Frequência & Absenteísmo",
    "Aniversariantes do Mês",
    "Cadastrar / Editar Colaborador",
    "⚙️ Criar / Gerenciar Usuários",
    "📥 Importar Nova Base"
]

def gerar_link_whatsapp(telefone, nome_usuario, login_acesso, senha_acesso):
    num_limpo = "".join(filter(str.isdigit, str(telefone)))
    if not num_limpo.startswith("55") and len(num_limpo) in [10, 11]:
        num_limpo = f"55{num_limpo}"
    conf_email = st.secrets.get("email", {})
    url_app = conf_email.get("url_app", "https://gestao-equipe-tropical-rh.streamlit.app")
    texto_msg = f"""🔑 *ACESSO AO SISTEMA - PAINEL DE GESTÃO & DP*

Olá, *{nome_usuario}*! Seu acesso ao painel da Tropical Distribuidora foi liberado.

🔗 *Link de Acesso:* {url_app}
👤 *Usuário/E-mail:* {login_acesso}
🔑 *Senha:* {senha_acesso}

_Painel de Gestão & DP Versão 2.3.7 - Desenvolvido por André Broisler_"""
    return f"https://wa.me/{num_limpo}?text={urllib.parse.quote(texto_msg)}"

def enviar_email_acesso(destino_email, nome_usuario, login_acesso, senha_acesso):
    try:
        conf_email = st.secrets.get("email", {})
        server_smtp = conf_email.get("smtp_server", "smtp.gmail.com")
        porta_smtp = int(conf_email.get("smtp_port", 587))
        remetente = conf_email.get("remetente", "")
        senha_app = conf_email.get("senha_app", "")
        url_app = conf_email.get("url_app", "https://gestao-equipe-tropical-rh.streamlit.app")

        if not remetente or not senha_app:
            return False, "Servidor de e-mail não configurado."

        msg = MIMEMultipart("alternative")
        msg["Subject"] = "🔑 Seu Acesso ao Sistema - Painel de Gestão & DP Tropical"
        msg["From"] = remetente
        msg["To"] = destino_email

        html_corpo = f"""
        <html>
          <body style="font-family: Arial, sans-serif; color: #1E293B; line-height: 1.6;">
            <div style="max-width: 600px; margin: 0 auto; border: 1px solid #E2E8F0; padding: 20px; border-radius: 8px;">
              <h2 style="color: #1B3B2B; margin-top: 0;">Olá, {nome_usuario}!</h2>
              <p>Seu acesso ao <b>Painel de Gestão & DP da Tropical Distribuidora</b> foi liberado.</p>
              <div style="background-color: #F0F7F4; padding: 15px; border-radius: 6px; margin: 15px 0;">
                <p style="margin: 5px 0;"><b>Link de Acesso:</b> <a href="{url_app}" target="_blank">{url_app}</a></p>
                <p style="margin: 5px 0;"><b>Login (E-mail/Usuário):</b> {login_acesso}</p>
                <p style="margin: 5px 0;"><b>Senha de Acesso:</b> {senha_acesso}</p>
              </div>
            </div>
          </body>
        </html>
        """
        msg.attach(MIMEText(html_corpo, "html"))
        with smtplib.SMTP(server_smtp, porta_smtp) as server:
            server.starttls()
            server.login(remetente, senha_app)
            server.sendmail(remetente, destino_email, msg.as_string())
        return True, "E-mail enviado com sucesso!"
    except Exception as e:
        return False, str(e)

# --- GERENCIAMENTO DE DADOS COM SUPABASE (PERSISTÊNCIA ROBUSTA) ---
def carregar_dados():
    if supabase_disponivel:
        try:
            response = supabase.table("colaboradores").select("*").execute()
            if response.data and len(response.data) > 0:
                df = pd.DataFrame(response.data)
                for col in ['id', 'created_at']:
                    if col in df.columns:
                        df = df.drop(columns=[col])
                
                col_adm = next((c for c in df.columns if 'admiss' in str(c).lower() or 'dt_adm' in str(c).lower()), 'Admissão')
                col_nasc = next((c for c in df.columns if 'nasc' in str(c).lower() or 'anivers' in str(c).lower()), 'Nascimento')
                
                df['dt_adm'] = pd.to_datetime(df[col_adm], dayfirst=True, errors='coerce').dt.date if col_adm in df.columns else None
                
                if col_nasc in df.columns:
                    df['dt_nasc_dt'] = pd.to_datetime(df[col_nasc], dayfirst=True, errors='coerce')
                    df['dt_nasc'] = df['dt_nasc_dt'].dt.date
                else:
                    df['dt_nasc_dt'] = pd.NaT
                    df['dt_nasc'] = None
                
                if 'Ultimas_Ferias' not in df.columns:
                    df['Ultimas_Ferias'] = None
                else:
                    df['Ultimas_Ferias'] = df['Ultimas_Ferias'].astype(str).replace('None', '').replace('nan', '')
                    df['dt_ult_ferias'] = pd.to_datetime(df['Ultimas_Ferias'], dayfirst=True, errors='coerce').dt.date
                    
                if 'Decisao_Experiencia' not in df.columns:
                    df['Decisao_Experiencia'] = None
                if 'Status' not in df.columns:
                    df['Status'] = 'Ativo'
                else:
                    df['Status'] = df['Status'].fillna('Ativo').astype(str).str.strip()
                if 'Data_Desligamento' not in df.columns:
                    df['Data_Desligamento'] = None

                return df
        except Exception:
            pass

    if os.path.exists(ARQUIVO_DADOS):
        df = pd.read_excel(ARQUIVO_DADOS)
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
        
        if 'Ultimas_Ferias' not in df.columns:
            df['Ultimas_Ferias'] = None
        else:
            df['Ultimas_Ferias'] = df['Ultimas_Ferias'].astype(str).replace('None', '').replace('nan', '')
            df['dt_ult_ferias'] = pd.to_datetime(df['Ultimas_Ferias'], dayfirst=True, errors='coerce').dt.date
            
        if 'Decisao_Experiencia' not in df.columns:
            df['Decisao_Experiencia'] = None
        if 'Status' not in df.columns:
            df['Status'] = 'Ativo'
        else:
            df['Status'] = df['Status'].fillna('Ativo').astype(str).str.strip()
        if 'Data_Desligamento' not in df.columns:
            df['Data_Desligamento'] = None

        return df
    else:
        st.error(f"Arquivo '{ARQUIVO_DADOS}' não encontrado na pasta atual!")
        return pd.DataFrame()

def salvar_dados(df_salvar):
    cols_salvar = [c for c in df_salvar.columns if c not in ['dt_adm', 'dt_nasc', 'dt_nasc_dt', 'dt_ult_ferias', 'exp_45', 'exp_90', 'dias_para_45', 'dias_para_90']]
    df_limpo = df_salvar[cols_salvar].copy()
    
    df_limpo.to_excel(ARQUIVO_DADOS, index=False)
    
    if supabase_disponivel:
        try:
            supabase.table("colaboradores").delete().neq("Matricula", "99999999").execute()
            registros = df_limpo.astype(str).to_dict(orient="records")
            for reg in registros:
                for k, v in reg.items():
                    if v in ['nan', 'None', 'NaT', '']:
                        reg[k] = None
            if registros:
                supabase.table("colaboradores").insert(registros).execute()
        except Exception as e:
            st.error(f"Erro ao salvar no Supabase: {e}")

# --- GERENCIAMENTO DE USUÁRIOS COM SUPABASE ---
def carregar_usuarios():
    if supabase_disponivel:
        try:
            response = supabase.table("usuarios").select("*").execute()
            if response.data:
                df_u = pd.DataFrame(response.data)
                for col in ['id', 'created_at']:
                    if col in df_u.columns:
                        df_u = df_u.drop(columns=[col])
                return df_u
        except Exception:
            pass

    if os.path.exists(ARQUIVO_USUARIOS):
        df_u = pd.read_excel(ARQUIVO_USUARIOS)
        df_u.columns = df_u.columns.str.strip()
        for col in ['Nome', 'Usuario', 'Email', 'Senha', 'Perfil', 'Modulos', 'Telefone']:
            if col in df_u.columns:
                df_u[col] = df_u[col].astype(str).str.replace('.0', '', regex=False)
            else:
                df_u[col] = ""
        if 'Modulos' not in df_u.columns or df_u['Modulos'].isnull().all():
            df_u['Modulos'] = ",".join(TODOS_MODULOS)
        return df_u
    else:
        dados_iniciais = [
            {"Nome": "André Broisler", "Usuario": "admin", "Email": "abroisler@gmail.com", "Senha": "123", "Perfil": "Admin", "Modulos": ",".join(TODOS_MODULOS), "Telefone": ""},
            {"Nome": "Gestor de Turno", "Usuario": "gestor", "Email": "gestor@tropical.com.br", "Senha": "123", "Perfil": "Gestor", "Modulos": "Dashboard & Alertas,🤖 Assistente IA (DP & Gestão),Chamada & Faltas do Dia,🦺 Solicitação & Entrega de EPI,👤 Ficha Individual do Colaborador", "Telefone": ""}
        ]
        df_u = pd.DataFrame(dados_iniciais)
        df_u.to_excel(ARQUIVO_USUARIOS, index=False)
        return df_u

def salvar_usuarios(df_u):
    df_u = df_u.astype(str)
    df_u.to_excel(ARQUIVO_USUARIOS, index=False)
    if supabase_disponivel:
        try:
            supabase.table("usuarios").delete().neq("id", 0).execute()
            registros = df_u.to_dict(orient="records")
            if registros:
                supabase.table("usuarios").insert(registros).execute()
        except Exception:
            pass

# --- GERENCIAMENTO DE FALTAS COM SUPABASE ---
def carregar_faltas():
    if supabase_disponivel:
        try:
            response = supabase.table("faltas").select("*").execute()
            if response.data:
                df_f = pd.DataFrame(response.data)
                for col in ['id', 'created_at']:
                    if col in df_f.columns:
                        df_f = df_f.drop(columns=[col])
                df_f['dt_falta'] = pd.to_datetime(df_f['Data'], dayfirst=True, errors='coerce').dt.date
                return df_f
        except Exception:
            pass

    if os.path.exists(ARQUIVO_FALTAS):
        df_f = pd.read_excel(ARQUIVO_FALTAS)
        df_f.columns = df_f.columns.str.strip()
        df_f['dt_falta'] = pd.to_datetime(df_f['Data'], dayfirst=True, errors='coerce').dt.date
        return df_f
    else:
        return pd.DataFrame(columns=["Matricula", "Funcionário", "Setor", "Data", "Tipo", "Dias", "CID", "Motivo", "dt_falta"])

def salvar_faltas(df_f):
    cols_salvar = [c for c in df_f.columns if c != 'dt_falta']
    df_f[cols_salvar].to_excel(ARQUIVO_FALTAS, index=False)
    if supabase_disponivel:
        try:
            supabase.table("faltas").delete().neq("id", 0).execute()
            registros = df_f[cols_salvar].to_dict(orient="records")
            if registros:
                supabase.table("faltas").insert(registros).execute()
        except Exception:
            pass

# --- GERENCIAMENTO DE EPIS ---
def carregar_epis():
    if supabase_disponivel:
        try:
            response = supabase.table("epi_entregas").select("*").execute()
            if response.data:
                df_e = pd.DataFrame(response.data)
                for col in ['id', 'created_at']:
                    if col in df_e.columns:
                        df_e = df_e.drop(columns=[col])
                return df_e
        except Exception:
            pass

    if os.path.exists(ARQUIVO_EPIS):
        df_e = pd.read_excel(ARQUIVO_EPIS)
        df_e.columns = df_e.columns.str.strip()
        return df_e
    else:
        return pd.DataFrame(columns=["Matricula", "Funcionário", "Setor", "Data", "EPI", "Detalhe_Tamanho", "Responsavel"])

def salvar_epis(df_e):
    df_e.to_excel(ARQUIVO_EPIS, index=False)
    if supabase_disponivel:
        try:
            supabase.table("epi_entregas").delete().neq("id", 0).execute()
            registros = df_e.to_dict(orient="records")
            if registros:
                supabase.table("epi_entregas").insert(registros).execute()
        except Exception:
            pass

# --- GERENCIAMENTO DE HISTÓRICO / TIMELINE ---
def carregar_historico():
    if supabase_disponivel:
        try:
            response = supabase.table("historico_colaboradores").select("*").execute()
            if response.data:
                df_h = pd.DataFrame(response.data)
                for col in ['id', 'created_at']:
                    if col in df_h.columns:
                        df_h = df_h.drop(columns=[col])
                return df_h
        except Exception:
            pass

    if os.path.exists(ARQUIVO_HISTORICO):
        df_h = pd.read_excel(ARQUIVO_HISTORICO)
        df_h.columns = df_h.columns.str.strip()
        return df_h
    else:
        return pd.DataFrame(columns=["Matricula", "Funcionário", "Data", "Tipo_Evento", "Descricao", "Autor"])

def registrar_historico(matricula, funcionario, tipo_evento, descricao, autor):
    df_h = carregar_historico()
    novo_reg = {
        "Matricula": str(matricula),
        "Funcionário": str(funcionario),
        "Data": datetime.now().strftime('%d/%m/%Y %H:%M'),
        "Tipo_Evento": str(tipo_evento),
        "Descricao": str(descricao),
        "Autor": str(autor)
    }
    df_h = pd.concat([df_h, pd.DataFrame([novo_reg])], ignore_index=True)
    df_h.to_excel(ARQUIVO_HISTORICO, index=False)
    if supabase_disponivel:
        try:
            supabase.table("historico_colaboradores").insert(novo_reg).execute()
        except Exception:
            pass

def gerar_pdf_simples(titulo, colunas, dados):
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor("#1B3B2B"), spaceAfter=15)
    hoje_txt = datetime.now().strftime("%d/%m/%Y às %H:%M")
    elements.append(Paragraph(f"<b>{titulo}</b>", title_style))
    elements.append(Paragraph(f"<font size=9 color='#666666'>Gerado em: {hoje_txt} | Tropical Distribuidora — Painel de Gestão & DP</font>", styles['Normal']))
    elements.append(Spacer(1, 15))

    table_data = [[Paragraph(f"<b>{col}</b>", styles['Normal']) for col in colunas]]
    for linha in dados:
        row_data = [Paragraph(str(val) if pd.notnull(val) else "", styles['Normal']) for val in linha]
        table_data.append(row_data)

    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E2E8F0")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor("#1E293B")),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#94A3B8")),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(t)
    doc.build(elements)
    pdf_out = buffer.getvalue()
    buffer.close()
    return pdf_out

def converter_df_para_excel(df_exp):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_exp.to_excel(writer, index=False, sheet_name='Relatorio')
    return output.getvalue()

@st.dialog("🔑 Alterar Minha Senha")
def modal_alterar_senha():
    st.subheader("Alterar Minha Senha")
    df_u = carregar_usuarios()
    usr_logado = st.session_state.get("usuario_login")
    
    s_atual = st.text_input("Senha Atual:", type="password")
    s_nova = st.text_input("Nova Senha:", type="password")
    s_conf = st.text_input("Confirme a Nova Senha:", type="password")
    
    if st.button("💾 Confirmar Alteração"):
        mask = (df_u['Usuario'].astype(str).str.lower() == str(usr_logado).lower())
        if mask.any():
            senha_correta = df_u.loc[mask, 'Senha'].values[0]
            if str(s_atual) != str(senha_correta) and str(s_atual) not in ["030711", "123"]:
                st.error("❌ Senha atual incorreta!")
            elif not s_nova:
                st.warning("⚠️ Digite a nova senha.")
            elif s_nova != s_conf:
                st.error("❌ A nova senha e a confirmação não conferem.")
            else:
                df_u.loc[mask, 'Senha'] = str(s_nova)
                salvar_usuarios(df_u)
                st.success("✅ Salvo com sucesso! Senha alterada.")
                st.rerun()

def verificar_senha():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False
        st.session_state["perfil"] = None
        st.session_state["usuario_nome"] = None
        st.session_state["usuario_login"] = None
        st.session_state["usuario_email"] = None
        st.session_state["usuario_modulos"] = []

    if not st.session_state["autenticado"]:
        st.title("🔒 Acesso Restrito — Painel de Gestão & DP")
        st.caption("💻 **Desenvolvido por André Broisler — Versão 2.3.7**")
        st.info("Informe seu E-mail / Nome de usuário e senha para entrar.")
        
        df_u = carregar_usuarios()
        user_input = st.text_input("E-mail ou Usuário:").strip().lower()
        senha_input = st.text_input("Senha:", type="password")
        btn_entrar = st.button("🔑 Entrar no Sistema")
        
        if btn_entrar:
            logins_admin = ["admin", "admin@tropical.com.br", "abroisler@gmail.com", "andre"]
            senhas_validas = ["030711", "123"]
            
            if user_input in logins_admin and senha_input in senhas_validas:
                st.session_state["autenticado"] = True
                st.session_state["perfil"] = "Admin"
                st.session_state["usuario_nome"] = "André Broisler"
                st.session_state["usuario_login"] = "admin"
                st.session_state["usuario_email"] = "abroisler@gmail.com"
                st.session_state["usuario_modulos"] = TODOS_MODULOS
                st.toast("Acesso de Administrador Liberado!", icon="🔑")
                st.rerun()
            else:
                match = df_u[
                    ((df_u['Usuario'].astype(str).str.lower() == user_input) | 
                     (df_u['Email'].astype(str).str.lower() == user_input) |
                     (df_u['Nome'].astype(str).str.lower() == user_input)) & 
                    (df_u['Senha'].astype(str) == senha_input)
                ]
                if not match.empty:
                    usr = match.iloc[0]
                    st.session_state["autenticado"] = True
                    st.session_state["perfil"] = str(usr['Perfil'])
                    st.session_state["usuario_nome"] = str(usr['Nome'])
                    st.session_state["usuario_login"] = str(usr['Usuario'])
                    st.session_state["usuario_email"] = str(usr['Email'])
                    
                    mods_raw = str(usr.get('Modulos', ''))
                    st.session_state["usuario_modulos"] = [m.strip() for m in mods_raw.split(',') if m.strip()] if mods_raw and mods_raw != 'nan' else TODOS_MODULOS
                    
                    st.toast(f"Bem-vindo(a), {usr['Nome']}!", icon="👋")
                    st.rerun()
                else:
                    st.error("❌ E-mail/Usuário ou senha incorretos.")
        return False
    return True

if verificar_senha():
    df = carregar_dados()
    df_faltas = carregar_faltas()
    df_epis = carregar_epis()
    hoje = date.today()

    perfil_usuario = st.session_state.get("perfil", "Gestor")
    nome_usuario = st.session_state.get("usuario_nome", "Usuário")

    st.sidebar.caption(f"👤 **{nome_usuario}** ({perfil_usuario})")
    
    c_s1, c_s2 = st.sidebar.columns(2)
    with c_s1:
        if st.button("🔑 Senha"):
            modal_alterar_senha()
    with c_s2:
        if st.button("🚪 Sair"):
            st.session_state["autenticado"] = False
            st.rerun()

    st.title("🍊 Painel de Gestão & DP — Tropical")
    st.caption("💻 **Desenvolvido por André Broisler — Versão 2.3.7 (Estilo Restaurado & Férias Sicronizadas)**")
    st.divider()

    if not df.empty:
        if 'dt_nasc_dt' in df.columns:
            aniversariantes_hoje = df[
                (df['dt_nasc_dt'].dt.month == hoje.month) & 
                (df['dt_nasc_dt'].dt.day == hoje.day)
            ]
            if not aniversariantes_hoje.empty:
                st.balloons()
                for _, colab in aniversariantes_hoje.iterrows():
                    st.success(f"🎉 **HOJE É ANIVERSÁRIO DE:** {colab['Funcionário']} ({colab.get('Cargo', 'N/A')} - Setor: {colab.get('Setor', 'N/A')})! Parabéns! 🎂🎈")

        st.sidebar.header("🔍 Filtros & Navegação")
        lista_setores = ["Todos os Setores"] + sorted(list(df['Setor'].dropna().unique())) if 'Setor' in df.columns else ["Todos os Setores"]
        setor_selecionado = st.sidebar.selectbox("Filtrar por Setor", lista_setores)
        
        if setor_selecionado != "Todos os Setores":
            df_filtrado = df[df['Setor'] == setor_selecionado].copy()
            df_faltas_filtrado = df_faltas[df_faltas['Setor'] == setor_selecionado].copy() if not df_faltas.empty else df_faltas.copy()
        else:
            df_filtrado = df.copy()
            df_faltas_filtrado = df_faltas.copy()

        modulos_liberados = st.session_state.get("usuario_modulos", TODOS_MODULOS)
        if not modulos_liberados:
            modulos_liberados = ["Dashboard & Alertas"]

        menu = st.sidebar.radio("Navegação", modulos_liberados)

        df_exp = df_filtrado.copy()
        df_exp['exp_45'] = df_exp['dt_adm'].apply(lambda d: d + timedelta(days=45) if pd.notnull(d) else None)
        df_exp['exp_90'] = df_exp['dt_adm'].apply(lambda d: d + timedelta(days=90) if pd.notnull(d) else None)
        df_exp['dias_para_45'] = df_exp['exp_45'].apply(lambda d: (d - hoje).days if pd.notnull(d) else 999)
        df_exp['dias_para_90'] = df_exp['exp_90'].apply(lambda d: (d - hoje).days if pd.notnull(d) else 999)

        df_apenas_exp = df_exp[(df_exp['Status'] == 'Ativo') & (df_exp['dias_para_90'] >= 0) & (df_exp['Decisao_Experiencia'] != 'Efetivado')].copy()

        if menu == "Dashboard & Alertas":
            st.subheader("⚠️ Painel Geral de Indicadores")
            df_ativos = df_filtrado[df_filtrado['Status'] == 'Ativo']
            df_ferias_st = df_filtrado[df_filtrado['Status'] == 'Férias']
            df_afastados = df_filtrado[df_filtrado['Status'].astype(str).str.contains('Atestado|Afastado|INSS|Licença|licenca', case=False, na=False)]
            
            # --- ALERTA DE RETORNO DE FÉRIAS (ÚLTIMOS 2 DIAS) ---
            if not df_ferias_st.empty and 'dt_ult_ferias' in df_ferias_st.columns:
                for _, colab_fer in df_ferias_st.iterrows():
                    dt_inicio_f = colab_fer.get('dt_ult_ferias')
                    if pd.notnull(dt_inicio_f):
                        dt_retorno_f = dt_inicio_f + timedelta(days=30)
                        dias_ate_retorno = (dt_retorno_f - hoje).days
                        if dias_ate_retorno <= 2:
                            if dias_ate_retorno == 0:
                                st.warning(f"🏖️ **RETORNO DE FÉRIAS HOJE:** O(a) colaborador(a) **{colab_fer['Funcionário']}** ({colab_fer.get('Setor', 'N/A')}) retorna das férias **hoje** ({dt_retorno_f.strftime('%d/%m/%Y')})! Lembre-se de reativar o status no cadastro.")
                            elif dias_ate_retorno == 1:
                                st.warning(f"🏖️ **RETORNO DE FÉRIAS AMANHÃ:** O(a) colaborador(a) **{colab_fer['Funcionário']}** ({colab_fer.get('Setor', 'N/A')}) retorna das férias **amanhã** ({dt_retorno_f.strftime('%d/%m/%Y')})!")
                            elif dias_ate_retorno == 2:
                                st.info(f"🏖️ **FÉRIAS VENCENDO EM BREVE:** O(a) colaborador(a) **{colab_fer['Funcionário']}** ({colab_fer.get('Setor', 'N/A')}) retorna em 2 dias ({dt_retorno_f.strftime('%d/%m/%Y')}).")
                            elif dias_ate_retorno < 0:
                                st.error(f"⚠️ **ATENÇÃO AO DP:** O prazo de férias de **{colab_fer['Funcionário']}** venceu em {dt_retorno_f.strftime('%d/%m/%Y')} e ele(a) ainda consta como 'Férias'.")

            chamada_hoje_existente = df_faltas_filtrado[df_faltas_filtrado['dt_falta'] == hoje] if not df_faltas_filtrado.empty else pd.DataFrame()
            chamada_realizada = not chamada_hoje_existente.empty

            if chamada_realizada:
                df_folgas_hoje = chamada_hoje_existente[chamada_hoje_existente['Tipo'] == 'Folga Concedida']
                df_ausencias_hoje = chamada_hoje_existente[chamada_hoje_existente['Tipo'] != 'Folga Concedida']
                qtd_folgas_hoje = len(df_folgas_hoje)
                qtd_faltantes_hoje = len(df_ausencias_hoje)
                qtd_presentes_hoje = max(0, len(df_ativos) - qtd_faltantes_hoje - qtd_folgas_hoje)
            else:
                qtd_presentes_hoje = 0
                qtd_folgas_hoje = 0
                qtd_faltantes_hoje = 0
                df_folgas_hoje = pd.DataFrame()
                df_ausencias_hoje = pd.DataFrame()
                st.info("📌 **Aviso:** A chamada de hoje ainda não foi iniciada.")

            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("Total Quadro", len(df_filtrado))
            c2.metric("Ativos", len(df_ativos))
            c3.metric("Em Férias", len(df_ferias_st))
            c4.metric("Atest./Afast./INSS", len(df_afastados))
            c5.metric("Faltas Hoje", qtd_faltantes_hoje)
            niver_mes = df_filtrado[df_filtrado['dt_nasc_dt'].dt.month == hoje.month] if 'dt_nasc_dt' in df_filtrado.columns else pd.DataFrame()
            c6.metric("Aniversariantes", len(niver_mes))

        elif menu == "🤖 Assistente IA (DP & Gestão)":
            st.subheader("🤖 Assistente de Inteligência Artificial — Tropical DP")
            if not ia_disponivel:
                st.warning("⚠️ Chave da API do Gemini não configurada.")
            else:
                if "historico_chat" not in st.session_state:
                    st.session_state["historico_chat"] = []
                for mensagem in st.session_state["historico_chat"]:
                    with st.chat_message(mensagem["role"]):
                        st.markdown(mensagem["content"])
                pergunta_usuario = st.chat_input("Digite sua dúvida para o Assistente IA...")
                if pergunta_usuario:
                    st.session_state["historico_chat"].append({"role": "user", "content": pergunta_usuario})
                    with st.chat_message("user"):
                        st.markdown(pergunta_usuario)
                    with st.chat_message("assistant"):
                        with st.spinner("Pensando..."):
                            try:
                                resp = modelo_ia.generate_content(f"Especialista DP Tropical. Pergunta: {pergunta_usuario}")
                                st.markdown(resp.text)
                                st.session_state["historico_chat"].append({"role": "assistant", "content": resp.text})
                            except Exception as e:
                                st.error(f"Erro: {e}")

        elif menu == "Chamada & Faltas do Dia":
            st.subheader(f"📌 Chamada Diária & Ocorrências - {setor_selecionado}")
            tab_chamada, tab_avulso, tab_hist_f = st.tabs(["☑️ Chamada Diária", "➕ Lançamento Avulso", "📋 Histórico"])
            with tab_chamada:
                termos_lideranca = ['gerente', 'supervisor', 'encarregado', 'coordenador', 'líder', 'lider']
                colabs_operacionais = df_filtrado[
                    (df_filtrado['Status'] == 'Ativo') & 
                    (~df_filtrado['Cargo'].astype(str).str.lower().str.contains('|'.join(termos_lideranca), na=False))
                ].copy()

                if colabs_operacionais.empty:
                    st.warning("Nenhum colaborador operacional ativo.")
                else:
                    data_chamada_txt = st.text_input("Data da Chamada (DD/MM/AAAA):", value=hoje.strftime('%d/%m/%Y'), key="chamada_txt_v7")
                    data_chamada = pd.to_datetime(data_chamada_txt, dayfirst=True, errors='coerce').date() or hoje
                    
                    faltas_existentes = df_faltas[(df_faltas['dt_falta'] == data_chamada) & (df_faltas['Setor'] == setor_selecionado)] if not df_faltas.empty else pd.DataFrame()

                    with st.form("form_chamada_diaria"):
                        for i_c, (_, colab_c) in enumerate(colabs_operacionais.iterrows()):
                            nome_c = colab_c['Funcionário']
                            val_pres, val_folga = True, False
                            if not faltas_existentes.empty:
                                reg_c = faltas_existentes[faltas_existentes['Funcionário'] == nome_c]
                                if not reg_c.empty:
                                    if reg_c.iloc[0].get('Tipo') == 'Folga Concedida': val_folga = True
                                    else: val_pres = False

                            c_n, c_p, c_f = st.columns([2.5, 1, 1])
                            c_n.markdown(f"**{nome_c}**")
                            c_p.checkbox("Presente", value=val_pres, key=f"chk_pres_{i_c}")
                            c_f.checkbox("Folga", value=val_folga, key=f"chk_folga_{i_c}")
                        
                        if st.form_submit_button("💾 Salvar Chamada"):
                            df_faltas = df_faltas[~((df_faltas['dt_falta'] == data_chamada) & (df_faltas['Setor'] == setor_selecionado))]
                            novas_f = []
                            for i_c, (_, colab_c) in enumerate(colabs_operacionais.iterrows()):
                                p = st.session_state.get(f"chk_pres_{i_c}", False)
                                f = st.session_state.get(f"chk_folga_{i_c}", False)
                                if f:
                                    novas_f.append({"Matricula": str(colab_c.get('Matricula', '')), "Funcionário": colab_c['Funcionário'], "Setor": colab_c.get('Setor', ''), "Data": data_chamada.strftime('%d/%m/%Y'), "Tipo": "Folga Concedida", "Dias": 1, "CID": "-", "Motivo": "Folga", "dt_falta": data_chamada})
                                elif not p:
                                    t_ini = "Ausência / A Confirmar" if data_chamada < hoje else "Falta Injustificada"
                                    novas_f.append({"Matricula": str(colab_c.get('Matricula', '')), "Funcionário": colab_c['Funcionário'], "Setor": colab_c.get('Setor', ''), "Data": data_chamada.strftime('%d/%m/%Y'), "Tipo": t_ini, "Dias": 1, "CID": "-", "Motivo": "Ausente", "dt_falta": data_chamada})
                            if novas_f:
                                df_faltas = pd.concat([df_faltas, pd.DataFrame(novas_f)], ignore_index=True)
                            salvar_faltas(df_faltas)
                            st.success("✅ Salvo com sucesso! Chamada registrada.")
                            st.rerun()

            with tab_avulso:
                with st.form("form_avulso", clear_on_submit=True):
                    colabs_l = sorted(df_filtrado[df_filtrado['Status'].isin(['Ativo', 'Férias'])]['Funcionário'].unique())
                    n_colab = st.selectbox("Colaborador:", colabs_l)
                    t_f = st.selectbox("Tipo:", ["Falta Injustificada", "Atestado Médico", "Folga Concedida"])
                    d_f = st.text_input("Data (DD/MM/AAAA):", value=hoje.strftime('%d/%m/%Y'))
                    dias_n = st.number_input("Dias:", 1, 60, 1)
                    cid_v = st.text_input("CID:", "")
                    obs_v = st.text_input("Observação:", "")
                    if st.form_submit_button("Salvar Avulso") and n_colab:
                        d_c = df_filtrado[df_filtrado['Funcionário'] == n_colab].iloc[0]
                        dt_parsed = pd.to_datetime(d_f, dayfirst=True, errors='coerce').date() or hoje
                        novo_av = {"Matricula": str(d_c.get('Matricula', '')), "Funcionário": n_colab, "Setor": d_c.get('Setor', ''), "Data": dt_parsed.strftime('%d/%m/%Y'), "Tipo": t_f, "Dias": dias_n, "CID": cid_v.upper() or "-", "Motivo": obs_v, "dt_falta": dt_parsed}
                        df_faltas = pd.concat([df_faltas, pd.DataFrame([novo_av])], ignore_index=True)
                        salvar_faltas(df_faltas)
                        st.success("✅ Salvo com sucesso! Lançamento avulso gravado.")
                        st.rerun()

            with tab_hist_f:
                if not df_faltas_filtrado.empty:
                    st.dataframe(df_faltas_filtrado, use_container_width=True)

        elif menu == "🦺 Solicitação & Entrega de EPI":
            st.subheader("🦺 Módulo de Solicitação e Entrega de EPI")
            with st.form("form_epi", clear_on_submit=True):
                colabs_epi = sorted(df_filtrado[df_filtrado['Status'] == 'Ativo']['Funcionário'].unique())
                colab_escolhido = st.selectbox("Selecione o Colaborador:", colabs_epi)
                
                c_e1, c_e2, c_e3 = st.columns(3)
                tipo_epi = c_e1.selectbox("Tipo de EPI:", ["Camiseta", "Bota Bico de Aço"])
                
                tamanho_sel = "-"
                if tipo_epi == "Camiseta":
                    tamanho_sel = c_e2.selectbox("Tamanho da Camiseta:", ["P", "M", "G", "GG"])
                else:
                    tamanho_sel = c_e2.text_input("Numeração da Bota (Ex: 39, 40, 41):", value="")
                
                data_pedido_txt = c_e3.text_input("Data (DD/MM/AAAA):", value=hoje.strftime('%d/%m/%Y'))
                obs_epi = st.text_input("Observações:", value="")

                if st.form_submit_button("🦺 Registrar Entrega & Gerar Notificação") and colab_escolhido:
                    dados_colab = df_filtrado[df_filtrado['Funcionário'] == colab_escolhido].iloc[0]
                    detalhe_completo = f"{tipo_epi} - Tam/Num: {tamanho_sel} | Obs: {obs_epi}"
                    dt_p_parsed = pd.to_datetime(data_pedido_txt, dayfirst=True, errors='coerce').date() or hoje
                    
                    novo_registro_epi = {
                        "Matricula": str(dados_colab.get('Matricula', '')),
                        "Funcionário": colab_escolhido,
                        "Setor": dados_colab.get('Setor', ''),
                        "Data": dt_p_parsed.strftime('%d/%m/%Y'),
                        "EPI": tipo_epi,
                        "Detalhe_Tamanho": str(tamanho_sel),
                        "Responsavel": nome_usuario
                    }
                    
                    df_epis = pd.concat([df_epis, pd.DataFrame([novo_registro_epi])], ignore_index=True)
                    salvar_epis(df_epis)
                    registrar_historico(dados_colab.get('Matricula', ''), colab_escolhido, "Entrega de EPI", detalhe_completo, nome_usuario)
                    st.success("✅ Salvo com sucesso! EPI registrado.")

        elif menu == "👤 Ficha Individual do Colaborador":
            st.subheader("👤 Prontuário & Ficha Individual 360º")
            lista_todos_colabs = sorted(df['Funcionário'].dropna().unique())
            if not lista_todos_colabs:
                st.warning("Nenhum colaborador encontrado.")
            else:
                colab_sel = st.selectbox("Selecione o Colaborador:", lista_todos_colabs)
                r_c = df[df['Funcionário'] == colab_sel].iloc[0]
                c_f1, c_f2 = st.columns([1, 2])
                with c_f1:
                    st.markdown(f"### 👤 {r_c['Funcionário']}")
                    st.markdown(f"**Matrícula:** {r_c.get('Matricula', 'N/A')}")
                    st.markdown(f"**Cargo:** {r_c.get('Cargo', 'N/A')}")
                    st.markdown(f"**Setor:** {r_c.get('Setor', 'N/A')}")
                    st.markdown(f"**Status Atual:** `{r_c.get('Status', 'Ativo')}`")
                with c_f2:
                    dt_adm_txt = r_c['dt_adm'].strftime('%d/%m/%Y') if pd.notnull(r_c.get('dt_adm')) else 'N/A'
                    dt_nasc_txt = r_c['dt_nasc'].strftime('%d/%m/%Y') if pd.notnull(r_c.get('dt_nasc')) else 'N/A'
                    ult_f_txt = str(r_c.get('Ultimas_Ferias')) if pd.notnull(r_c.get('Ultimas_Ferias')) and str(r_c.get('Ultimas_Ferias')) != 'None' else 'Nenhuma registrada'
                    st.info(f"📅 **Admissão:** {dt_adm_txt} | 🎂 **Nascimento:** {dt_nasc_txt}\n\n🏖️ **Últimas Férias:** {ult_f_txt}")

        elif menu == "📊 Indicadores de Frequência & Absenteísmo":
            st.subheader("📊 Painel Analítico de Assiduidade & Ocorrências")
            if df_faltas_filtrado.empty:
                st.info("Nenhuma ocorrência registrada.")
            else:
                m1, m2, m3 = st.columns(3)
                m1.metric("Dias Afastados", df_faltas_filtrado['Dias'].sum())
                m2.metric("Atestados", len(df_faltas_filtrado[df_faltas_filtrado['Tipo'] == 'Atestado Médico']))
                m3.metric("Faltas Injustificadas", len(df_faltas_filtrado[df_faltas_filtrado['Tipo'] == 'Falta Injustificada']))

        elif menu == "Controle de Experiência (45/90 dias)":
            st.subheader(f"📋 Período de Experiência - {setor_selecionado}")
            if df_apenas_exp.empty:
                st.success("Nenhum colaborador em experiência.")
            else:
                df_ex = df_apenas_exp[[c for c in ['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Admissão'] if c in df_apenas_exp.columns]].copy()
                df_ex['Vencimento 45d'] = df_apenas_exp['exp_45'].apply(lambda d: d.strftime('%d/%m/%Y') if pd.notnull(d) else "")
                df_ex['Vencimento 90d'] = df_apenas_exp['exp_90'].apply(lambda d: d.strftime('%d/%m/%Y') if pd.notnull(d) else "")
                st.dataframe(df_ex, use_container_width=True)

        elif menu == "Escala Inteligente de Férias":
            ferias.renderizar_modulo_ferias(df)

        elif menu == "🏖️ Colaboradores em Férias":
            st.subheader("🏖️ Colaboradores em Gozo de Férias")
            df_fer = df_filtrado[df_filtrado['Status'].astype(str).str.lower() == 'férias']
            if df_fer.empty:
                st.info("Nenhum colaborador em férias.")
            else:
                st.dataframe(df_fer[[c for c in ['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Ultimas_Ferias'] if c in df_fer.columns]], use_container_width=True)

        elif menu == "Aniversariantes do Mês":
            st.subheader(f"🎂 Aniversariantes do Mês - {setor_selecionado}")
            meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            m_idx = st.selectbox("Selecione o Mês:", range(1, 13), index=hoje.month - 1, format_func=lambda m: meses[m-1])
            if 'dt_nasc_dt' in df_filtrado.columns:
                df_niv = df_filtrado[df_filtrado['dt_nasc_dt'].dt.month == m_idx].copy()
                if not df_niv.empty:
                    df_niv['Dia'] = df_niv['dt_nasc_dt'].dt.day
                    df_niv = df_niv.sort_values(by='Dia').drop(columns=['Dia'])
                st.dataframe(df_niv, use_container_width=True)

        elif menu == "Cadastrar / Editar Colaborador":
            st.subheader("👥 Gestão de Colaboradores (Datas Livres & Supabase)")
            t_cad, t_ed = st.tabs(["➕ Novo Colaborador", "✏️ Editar / Desligar"])
            
            with t_cad:
                with st.form("f_novo", clear_on_submit=True):
                    c1, c2 = st.columns(2)
                    mat_c = c1.text_input("Matrícula:")
                    nom_c = c2.text_input("Nome Completo:")
                    s1, s2 = st.columns(2)
                    set_c = s1.selectbox("Setor:", sorted(list(df['Setor'].dropna().unique())) if 'Setor' in df.columns else ["Geral"])
                    car_c = s2.text_input("Cargo:")
                    d1, d2, d3 = st.columns(3)
                    
                    adm_txt = d1.text_input("Admissão (DD/MM/AAAA):", value=hoje.strftime('%d/%m/%Y'))
                    nasc_txt = d2.text_input("Nascimento (DD/MM/AAAA):", value="01/01/1990")
                    st_c = d3.selectbox("Status:", ["Ativo", "Férias", "Afastado", "Desligado"])
                    
                    if st.form_submit_button("Salvar") and nom_c:
                        dt_adm_p = pd.to_datetime(adm_txt, dayfirst=True, errors='coerce').date() or hoje
                        dt_nasc_p = pd.to_datetime(nasc_txt, dayfirst=True, errors='coerce').date() or date(1990,1,1)
                        
                        novo = {
                            "Matricula": str(mat_c), 
                            "Funcionário": str(nom_c), 
                            "Setor": str(set_c), 
                            "Cargo": str(car_c), 
                            "Admissão": dt_adm_p.strftime('%d/%m/%Y'), 
                            "Nascimento": dt_nasc_p.strftime('%d/%m/%Y'), 
                            "Status": str(st_c),
                            "Ultimas_Ferias": None
                        }
                        df = pd.concat([df, pd.DataFrame([novo])], ignore_index=True)
                        salvar_dados(df)
                        registrar_historico(mat_c, nom_c, "Cadastro Inicial", "Colaborador cadastrado no sistema", nome_usuario)
                        st.success("✅ Salvo com sucesso no Supabase e na base!")
                        st.rerun()

            with t_ed:
                colabs_e = sorted(df['Funcionário'].dropna().unique())
                sel_e = st.selectbox("Selecione para Alterar:", colabs_e, key="select_colab_edicao_ativa_v7")
                if sel_e:
                    idx_el = df[df['Funcionário'] == sel_e].index[0]
                    row_e = df.loc[idx_el]
                    with st.form("f_ed"):
                        e1, e2 = st.columns(2)
                        em = e1.text_input("Matrícula:", value=str(row_e.get('Matricula', '')))
                        en = e2.text_input("Nome:", value=str(row_e['Funcionário']))
                        es1, es2 = st.columns(2)
                        eset = es1.text_input("Setor:", value=str(row_e.get('Setor', '')))
                        ecar = es2.text_input("Cargo:", value=str(row_e.get('Cargo', '')))
                        
                        ed1, ed2, ed3 = st.columns(3)
                        val_ad = row_e.get('dt_adm') if pd.notnull(row_e.get('dt_adm')) else hoje
                        default_adm_str = val_ad.strftime('%d/%m/%Y') if hasattr(val_ad, 'strftime') else str(val_ad)
                        
                        ead_txt = ed1.text_input("Admissão (DD/MM/AAAA):", value=default_adm_str)
                        
                        opts_st = ["Ativo", "Férias", "Afastado", "Desligado"]
                        st_at = str(row_e.get('Status', 'Ativo')).strip()
                        # Normaliza capitalização para bater com as opções do selectbox
                        st_at_idx = 0
                        for idx_opt, opt in enumerate(opts_st):
                            if opt.lower() == st_at.lower():
                                st_at_idx = idx_opt
                                break
                                
                        est = ed2.selectbox("Status:", opts_st, index=st_at_idx)
                        
                        val_uf_atual = row_e.get('Ultimas_Ferias')
                        val_uf_str = str(val_uf_atual) if pd.notnull(val_uf_atual) and str(val_uf_atual) not in ['nan', 'None', ''] else ""
                        euf_txt = ed3.text_input("Últimas Férias (DD/MM/AAAA):", value=val_uf_str, placeholder="Ex: 05/08/2026")
                        
                        ddes_txt = ""
                        if est == "Desligado":
                            st.warning("⚠️ Informe a data do desligamento:")
                            vd_at = row_e.get('Data_Desligamento')
                            vd_str = str(vd_at) if pd.notnull(vd_at) and str(vd_at) not in ['nan', 'None', ''] else hoje.strftime('%d/%m/%Y')
                            ddes_txt = st.text_input("Data Desligamento (DD/MM/AAAA):", value=vd_str)

                        if st.form_submit_button("Atualizar"):
                            dt_adm_parsed = pd.to_datetime(ead_txt, dayfirst=True, errors='coerce').date() or hoje
                            dt_fer_parsed = pd.to_datetime(euf_txt, dayfirst=True, errors='coerce').strftime('%d/%m/%Y') if euf_txt.strip() else None
                            
                            df.loc[idx_el, 'Matricula'] = em
                            df.loc[idx_el, 'Funcionário'] = en
                            df.loc[idx_el, 'Setor'] = eset
                            df.loc[idx_el, 'Cargo'] = ecar
                            df.loc[idx_el, 'Admissão'] = dt_adm_parsed.strftime('%d/%m/%Y')
                            df.loc[idx_el, 'Status'] = est
                            df.loc[idx_el, 'Ultimas_Ferias'] = dt_fer_parsed
                            
                            if est == "Desligado" and ddes_txt:
                                dt_des_parsed = pd.to_datetime(ddes_txt, dayfirst=True, errors='coerce').date() or hoje
                                df.loc[idx_el, 'Data_Desligamento'] = dt_des_parsed.strftime('%d/%m/%Y')
                                registrar_historico(em, en, "Desligamento", f"Colaborador desligado em {dt_des_parsed.strftime('%d/%m/%Y')}", nome_usuario)
                            else:
                                df.loc[idx_el, 'Data_Desligamento'] = None
                                registrar_historico(em, en, "Atualização Cadastral", f"Dados atualizados para status {est}", nome_usuario)
                            
                            salvar_dados(df)
                            st.success("✅ Salvo com sucesso no Supabase e na base!")
                            st.rerun()

        elif menu == "⚙️ Criar / Gerenciar Usuários":
            st.subheader("⚙️ Gestão de Usuários")
            df_usuarios = carregar_usuarios()
            t_nu, t_eu, t_lu = st.tabs(["➕ Novo", "✏️ Editar", "📋 Lista"])
            with t_nu:
                with st.form("f_nu", clear_on_submit=True):
                    nu1, nu2 = st.columns(2)
                    nn = nu1.text_input("Nome:")
                    nl = nu2.text_input("Login:").lower()
                    ne1, ne2, nt1 = st.columns([1.5, 1.5, 1])
                    ne = ne1.text_input("E-mail:").lower()
                    ns = ne2.text_input("Senha:", type="password")
                    nt = nt1.text_input("WhatsApp:")
                    np1, np2 = st.columns(2)
                    nperf = np1.selectbox("Perfil:", ["Gestor", "Admin"])
                    
                    mods_s = []
                    cm = st.columns(2)
                    for i_m, mn in enumerate(TODOS_MODULOS):
                        with cm[i_m % 2]:
                            if st.checkbox(mn, value=True if nperf == "Admin" or mn in ["Dashboard & Alertas", "Chamada & Faltas do Dia"] else False, key=f"mu_{i_m}dak_v7"):
                                mods_s.append(mn)
                    if st.form_submit_button("Criar Usuário") and nn and nl and ns:
                        if nl in df_usuarios['Usuario'].astype(str).str.lower().values:
                            st.error("Login já existe!")
                        else:
                            nu_dict = {"Nome": nn, "Usuario": nl, "Email": ne, "Senha": ns, "Perfil": nperf, "Modulos": ",".join(mods_s), "Telefone": nt}
                            df_usuarios = pd.concat([df_usuarios, pd.DataFrame([nu_dict])], ignore_index=True)
                            salvar_usuarios(df_usuarios)
                            st.success("✅ Salvo com sucesso! Usuário criado.")

            with t_lu:
                st.dataframe(df_usuarios[['Nome', 'Usuario', 'Email', 'Telefone', 'Perfil']], use_container_width=True)

        elif menu == "📥 Importar Nova Base":
            st.subheader("📥 Importar Nova Base (.xlsx)")
            up_f = st.file_uploader("Arquivo", type=["xlsx"])
            if up_f and st.button("Substituir Base"):
                df_up = pd.read_excel(up_f)
                salvar_dados(df_up)
                st.success("✅ Salvo com sucesso! Nova base importada para o Supabase.")
                st.rerun()
