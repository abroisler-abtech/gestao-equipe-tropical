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

importlib.reload(ferias)

st.set_page_config(page_title="Gestão de Equipe Tropical", page_icon="👥", layout="wide")

ARQUIVO_DADOS = "equipe.xlsx"
ARQUIVO_FALTAS = "faltas.xlsx"
ARQUIVO_USUARIOS = "usuarios.xlsx"

TODOS_MODULOS = [
    "Dashboard & Alertas",
    "Chamada & Faltas do Dia",
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

# --- GERADOR DE LINK PARA O WHATSAPP ---
def gerar_link_whatsapp(telefone, nome_usuario, login_acesso, senha_acesso):
    num_limpo = "".join(filter(str.isdigit, str(telefone)))
    if not num_limpo.startswith("55") and len(num_limpo) in [10, 11]:
        num_limpo = f"55{num_limpo}"
        
    conf_email = st.secrets.get("email", {})
    url_app = conf_email.get("url_app", "https://gestao-equipe-tropical-rh.streamlit.app")
    
    texto_msg = f"""🔑 *ACESSO AO SISTEMA - GESTÃO DE PESSOAS*

Olá, *{nome_usuario}*! Seu acesso ao painel da Tropical Distribuidora foi liberado.

🔗 *Link de Acesso:* {url_app}
👤 *Usuário/E-mail:* {login_acesso}
🔑 *Senha:* {senha_acesso}

_Gestão de Pessoas Versão 2.0 - Desenvolvido por André Broisler_"""

    texto_encoded = urllib.parse.quote(texto_msg)
    return f"https://wa.me/{num_limpo}?text={texto_encoded}"

# --- DISPARO DE E-MAIL SMTP ---
def enviar_email_acesso(destino_email, nome_usuario, login_acesso, senha_acesso):
    try:
        conf_email = st.secrets.get("email", {})
        server_smtp = conf_email.get("smtp_server", "smtp.gmail.com")
        porta_smtp = int(conf_email.get("smtp_port", 587))
        remetente = conf_email.get("remetente", "")
        senha_app = conf_email.get("senha_app", "")
        url_app = conf_email.get("url_app", "https://gestao-equipe-tropical-rh.streamlit.app")

        if not remetente or not senha_app:
            return False, "Servidor de e-mail não configurado. Utilize o link direto de envio por WhatsApp."

        msg = MIMEMultipart("alternative")
        msg["Subject"] = "🔑 Seu Acesso ao Sistema - Gestão de Equipe Tropical"
        msg["From"] = remetente
        msg["To"] = destino_email

        html_corpo = f"""
        <html>
          <body style="font-family: Arial, sans-serif; color: #1E293B; line-height: 1.6;">
            <div style="max-width: 600px; margin: 0 auto; border: 1px solid #E2E8F0; padding: 20px; border-radius: 8px;">
              <h2 style="color: #1E3A8A; margin-top: 0;">Olá, {nome_usuario}!</h2>
              <p>Seu acesso ao painel de <b>Gestão de Equipe da Tropical Distribuidora</b> foi liberado.</p>
              <div style="background-color: #F8FAFC; padding: 15px; border-radius: 6px; margin: 15px 0;">
                <p style="margin: 5px 0;"><b>Link de Acesso:</b> <a href="{url_app}" target="_blank">{url_app}</a></p>
                <p style="margin: 5px 0;"><b>Login (E-mail/Usuário):</b> {login_acesso}</p>
                <p style="margin: 5px 0;"><b>Senha de Acesso:</b> {senha_acesso}</p>
              </div>
              <p>Recomendamos alterar sua senha no primeiro acesso utilizando o botão <b>🔑 Senha</b> no menu lateral.</p>
              <hr style="border: 0; border-top: 1px solid #CBD5E1; margin: 20px 0;" />
              <p style="font-size: 11px; color: #64748B;">Gestão de Pessoas Versão 2.0 | Desenvolvido por <b>André Broisler</b></p>
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

# --- GERENCIAMENTO DE USUÁRIOS E PERMISSÕES ---
def carregar_usuarios():
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
            {"Nome": "Gestor de Turno", "Usuario": "gestor", "Email": "gestor@tropical.com.br", "Senha": "123", "Perfil": "Gestor", "Modulos": "Dashboard & Alertas,Chamada & Faltas do Dia,👤 Ficha Individual do Colaborador", "Telefone": ""}
        ]
        df_u = pd.DataFrame(dados_iniciais)
        df_u.to_excel(ARQUIVO_USUARIOS, index=False)
        return df_u

def salvar_usuarios(df_u):
    df_u = df_u.astype(str)
    df_u.to_excel(ARQUIVO_USUARIOS, index=False)

# --- GERADOR DE RELATÓRIOS EM PDF ---
def gerar_pdf_simples(titulo, colunas, dados):
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor("#1E3A8A"), spaceAfter=15)
    hoje_txt = datetime.now().strftime("%d/%m/%Y às %H:%M")
    elements.append(Paragraph(f"<b>{titulo}</b>", title_style))
    elements.append(Paragraph(f"<font size=9 color='#666666'>Gerado em: {hoje_txt} | Tropical Distribuidora — Gestão de Pessoas Versão 2.0 - Desenvolvido por André Broisler</font>", styles['Normal']))
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

def gerar_pdf_dashboard_completo(setor_nome, df_filtrado, total_q, ativos, ferias_cnt, afastados_cnt, ocorrencias_cnt):
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    elements = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=15, leading=18, textColor=colors.HexColor("#1E3A8A"), spaceAfter=5)
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor("#475569"), spaceAfter=10)

    hoje_txt = datetime.now().strftime("%d/%m/%Y às %H:%M")
    elements.append(Paragraph("<b>RELATÓRIO GERAL DE DASHBOARD & INDICADORES DA EQUIPE</b>", title_style))
    elements.append(Paragraph(f"<b>Setor Filtrado:</b> {setor_nome} | <b>Gerado em:</b> {hoje_txt} | Tropical Distribuidora — Gestão de Pessoas Versão 2.0 - Desenvolvido por André Broisler", sub_style))
    elements.append(Spacer(1, 5))

    indicadores_data = [
        ["Total Quadro", "Ativos", "Em Férias", "Afastados/INSS", "Ocorrências (Mês)"],
        [str(total_q), str(ativos), str(ferias_cnt), str(afastados_cnt), str(ocorrencias_cnt)]
    ]
    t_ind = Table(indicadores_data, colWidths=[130]*5)
    t_ind.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1E3A8A")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,1), (-1,1), colors.HexColor("#F1F5F9")),
        ('TEXTCOLOR', (0,1), (-1,1), colors.HexColor("#0F172A")),
        ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#94A3B8")),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(t_ind)
    elements.append(Spacer(1, 15))

    elements.append(Paragraph("<b>Quadro Atual de Colaboradores</b>", styles['Heading2']))
    elements.append(Spacer(1, 5))
    
    cols_pres = [c for c in ['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Status', 'Admissão'] if c in df_filtrado.columns]
    table_data = [[Paragraph(f"<b>{col}</b>", styles['Normal']) for col in cols_pres]]
    
    for _, row in df_filtrado[cols_pres].iterrows():
        r_data = [Paragraph(str(val) if pd.notnull(val) else "", styles['Normal']) for val in row]
        table_data.append(r_data)

    t_quadro = Table(table_data, repeatRows=1)
    t_quadro.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E2E8F0")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor("#1E293B")),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#94A3B8")),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_quadro)

    doc.build(elements)
    pdf_out = buffer.getvalue()
    buffer.close()
    return pdf_out

def converter_df_para_excel(df_exp):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_exp.to_excel(writer, index=False, sheet_name='Relatorio')
    return output.getvalue()

# --- AUTENTICAÇÃO COM NOME/E-MAIL E ACESSO MESTRE ---
def verificar_senha():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False
        st.session_state["perfil"] = None
        st.session_state["usuario_nome"] = None
        st.session_state["usuario_login"] = None
        st.session_state["usuario_email"] = None
        st.session_state["usuario_modulos"] = []

    if not st.session_state["autenticado"]:
        st.title("🔒 Acesso Restrito - Gestão de Equipe Tropical")
        st.caption("💻 **Gestão de Pessoas Versão 2.0 - Desenvolvido por André Broisler**")
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
    def carregar_dados():
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
                df['Status'] = df['Status'].fillna('Ativo')

            return df
        else:
            st.error(f"Arquivo '{ARQUIVO_DADOS}' não encontrado na pasta atual!")
            return pd.DataFrame()

    def carregar_faltas():
        if os.path.exists(ARQUIVO_FALTAS):
            df_f = pd.read_excel(ARQUIVO_FALTAS)
            df_f.columns = df_f.columns.str.strip()
            df_f['dt_falta'] = pd.to_datetime(df_f['Data'], dayfirst=True, errors='coerce').dt.date
            return df_f
        else:
            return pd.DataFrame(columns=["Matricula", "Funcionário", "Setor", "Data", "Tipo", "Dias", "CID", "Motivo", "dt_falta"])

    def salvar_dados(df_salvar):
        cols_salvar = [c for c in df_salvar.columns if c not in ['dt_adm', 'dt_nasc', 'dt_nasc_dt', 'dt_ult_ferias', 'exp_45', 'exp_90', 'dias_para_45', 'dias_para_90']]
        df_salvar[cols_salvar].to_excel(ARQUIVO_DADOS, index=False)

    def salvar_faltas(df_f):
        cols_salvar = [c for c in df_f.columns if c != 'dt_falta']
        df_f[cols_salvar].to_excel(ARQUIVO_FALTAS, index=False)

    @st.dialog("📋 Lista Detalhada e Exportação")
    def exibir_modal_detalhes(titulo, df_detalhes):
        st.subheader(titulo)
        if df_detalhes.empty:
            st.info("Nenhum colaborador nesta situação.")
        else:
            st.dataframe(df_detalhes, use_container_width=True)
            st.markdown("---")
            st.markdown("##### 📥 Exportar Esta Lista")
            c_d1, c_d2 = st.columns(2)
            with c_d1:
                st.download_button(
                    label="📥 Baixar em Excel (.xlsx)",
                    data=converter_df_para_excel(df_detalhes),
                    file_name=f"{titulo.lower().replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"excel_modal_{titulo}"
                )
            with c_d2:
                pdf_b = gerar_pdf_simples(titulo, list(df_detalhes.columns), df_detalhes.values.tolist())
                st.download_button(
                    label="🖨️ Baixar PDF para Impressão",
                    data=pdf_b,
                    file_name=f"{titulo.lower().replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    key=f"pdf_modal_{titulo}"
                )

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
                    st.toast("✅ Senha alterada com sucesso!", icon="🎉")
                    st.rerun()

    df = carregar_dados()
    df_faltas = carregar_faltas()
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

    st.title("👥 Gestão de Equipe Tropical")
    st.caption("💻 **Gestão de Pessoas Versão 2.0 - Desenvolvido por André Broisler**")
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

        # CÁLCULOS DE EXPERIÊNCIA
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
            
            faltas_hoje = df_faltas_filtrado[df_faltas_filtrado['dt_falta'] == hoje] if not df_faltas_filtrado.empty else pd.DataFrame()
            
            df_folgas_hoje = faltas_hoje[faltas_hoje['Tipo'] == 'Folga Concedida'] if not faltas_hoje.empty else pd.DataFrame()
            df_ausencias_hoje = faltas_hoje[faltas_hoje['Tipo'] != 'Folga Concedida'] if not faltas_hoje.empty else pd.DataFrame()
            
            qtd_folgas_hoje = len(df_folgas_hoje)
            qtd_faltantes_hoje = len(df_ausencias_hoje)
            qtd_presentes_hoje = max(0, len(df_ativos) - qtd_faltantes_hoje - qtd_folgas_hoje)

            exp_criticos = df_apenas_exp[
                ((df_apenas_exp['dias_para_45'] >= 0) & (df_apenas_exp['dias_para_45'] <= 10)) | 
                ((df_apenas_exp['dias_para_90'] >= 0) & (df_apenas_exp['dias_para_90'] <= 10))
            ]
            if not exp_criticos.empty:
                st.warning(f"⏰ **ALERTA DE EXPERIÊNCIA:** Existem {len(exp_criticos)} contrato(s) de experiência vencendo nos próximos 10 dias!")

            cd1, cd2 = st.columns(2)
            with cd1:
                st.download_button(
                    label="📥 Exportar Dados do Dashboard (.xlsx)",
                    data=converter_df_para_excel(df_filtrado),
                    file_name=f"dashboard_dados_{setor_selecionado.lower().replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dash_excel"
                )
            with cd2:
                pdf_dash = gerar_pdf_dashboard_completo(
                    setor_selecionado, df_filtrado, len(df_filtrado), len(df_ativos), len(df_ferias_st), len(df_afastados), qtd_faltantes_hoje
                )
                st.download_button(
                    label="🖨️ Imprimir Relatório Geral do Dashboard (PDF)",
                    data=pdf_dash,
                    file_name=f"dashboard_relatorio_{setor_selecionado.lower().replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    key="btn_dash_pdf"
                )

            st.markdown("---")

            p1, p2, p3 = st.columns(3)
            
            with p1:
                st.metric("🟢 Presentes Hoje", qtd_presentes_hoje)
                if st.button("🔍 Ver Presentes de Hoje", key="btn_ver_pres_hoje"):
                    nomes_ausentes_ou_folga = faltas_hoje['Funcionário'].tolist() if not faltas_hoje.empty else []
                    df_pres_detalhe = df_ativos[~df_ativos['Funcionário'].isin(nomes_ausentes_ou_folga)]
                    cols_m = [c for c in ['Matricula', 'Funcionário', 'Setor', 'Cargo'] if c in df_pres_detalhe.columns]
                    exibir_modal_detalhes(f"Colaboradores Presentes em {hoje.strftime('%d/%m/%Y')}", df_pres_detalhe[cols_m])
                    
            with p2:
                st.metric("🏖️ Folgas Hoje", qtd_folgas_hoje)
                if st.button("🔍 Ver Folgas de Hoje", key="btn_ver_folgas_hoje"):
                    cols_m = [c for c in ['Data', 'Funcionário', 'Setor', 'Motivo'] if c in df_folgas_hoje.columns]
                    exibir_modal_detalhes(f"Colaboradores de Folga em {hoje.strftime('%d/%m/%Y')}", df_folgas_hoje[cols_m] if not df_folgas_hoje.empty else pd.DataFrame())

            with p3:
                st.metric("🔴 Faltantes / Ausentes Hoje", qtd_faltantes_hoje)
                if st.button("🔍 Ver Faltantes de Hoje", key="btn_ver_faltas_hoje"):
                    cols_m = [c for c in ['Data', 'Funcionário', 'Setor', 'Tipo', 'Motivo'] if c in df_ausencias_hoje.columns]
                    exibir_modal_detalhes(f"Colaboradores Ausentes em {hoje.strftime('%d/%m/%Y')}", df_ausencias_hoje[cols_m] if not df_ausencias_hoje.empty else pd.DataFrame())

            st.markdown("---")

            vagas_abertas = df_filtrado[df_filtrado['Status'].astype(str).str.contains('Desligado', case=False, na=False)]
            if not vagas_abertas.empty:
                st.error(f"🚨 **ALERTA DE REPOSIÇÃO DE QUADRO:** Existem {len(vagas_abertas)} vaga(s) aberta(s) por desligamento/término de contrato!")

            c1, c2, c3, c4, c5, c6 = st.columns(6)
            
            c1.metric("Total Quadro", len(df_filtrado))
            if c1.button("🔍 Ver Quadro", key="btn_quadro"):
                cols_m = [c for c in ['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Status', 'Admissão'] if c in df_filtrado.columns]
                exibir_modal_detalhes("Quadro Geral de Colaboradores", df_filtrado[cols_m])
                
            c2.metric("Ativos", len(df_ativos))
            if c2.button("🔍 Ver Ativos", key="btn_ativos"):
                cols_m = [c for c in ['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Admissão'] if c in df_ativos.columns]
                exibir_modal_detalhes("Colaboradores Ativos no Quadro", df_ativos[cols_m])

            c3.metric("Em Férias", len(df_ferias_st))
            if c3.button("🔍 Ver Férias", key="btn_ferias"):
                cols_m = [c for c in ['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Ultimas_Ferias'] if c in df_ferias_st.columns]
                exibir_modal_detalhes("Colaboradores em Gozo de Férias", df_ferias_st[cols_m])

            c4.metric("Atest./Afast./INSS", len(df_afastados))
            if c4.button("🔍 Ver Afastados", key="btn_afastados"):
                cols_m = [c for c in ['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Status', 'Contato'] if c in df_afastados.columns]
                exibir_modal_detalhes("Colaboradores Afastados / Atestado / INSS", df_afastados[cols_m])

            c5.metric("Faltas Hoje", qtd_faltantes_hoje)
            if c5.button("🔍 Ver Faltas", key="btn_faltas_quadro"):
                cols_m = [c for c in ['Data', 'Funcionário', 'Setor', 'Tipo', 'Dias', 'CID', 'Motivo'] if c in df_ausencias_hoje.columns]
                exibir_modal_detalhes(f"Colaboradores Ausentes em {hoje.strftime('%d/%m/%Y')}", df_ausencias_hoje[cols_m] if not df_ausencias_hoje.empty else pd.DataFrame())

            niver_mes = df_filtrado[df_filtrado['dt_nasc_dt'].dt.month == hoje.month] if 'dt_nasc_dt' in df_filtrado.columns else pd.DataFrame()
            c6.metric("Aniversariantes", len(niver_mes))
            if c6.button("🔍 Ver Aniversár.", key="btn_niver_m"):
                cols_m = [c for c in ['Nascimento', 'Funcionário', 'Setor', 'Cargo'] if c in niver_mes.columns]
                exibir_modal_detalhes(f"Aniversariantes do Mês ({hoje.strftime('%m/%Y')})", niver_mes[cols_m])

            st.markdown("---")
            
            g1, g2 = st.columns(2)
            with g1:
                df_status_cnt = df_filtrado['Status'].value_counts().reset_index()
                df_status_cnt.columns = ['Status', 'Quantidade']
                fig_status = px.pie(df_status_cnt, values='Quantidade', names='Status', title="Distribuição de Status do Quadro", hole=0.4)
                st.plotly_chart(fig_status, use_container_width=True)
                
            with g2:
                if not df_faltas_filtrado.empty:
                    df_tipo_falta = df_faltas_filtrado.groupby('Tipo')['Dias'].sum().reset_index()
                    fig_faltas = px.bar(df_tipo_falta, x='Tipo', y='Dias', title="Total de Dias Perdidos por Tipo (Geral)", text_auto=True, color='Tipo')
                    st.plotly_chart(fig_faltas, use_container_width=True)
                else:
                    st.info("Sem dados de ocorrências para gerar o gráfico de ausências.")

        elif menu == "Chamada & Faltas do Dia":
            st.subheader(f"📌 Chamada Diária de Presença & Ocorrências - {setor_selecionado}")
            
            tab_chamada, tab_avulso, tab_hist_f = st.tabs(["☑️ Chamada Diária (Presença)", "➕ Lançamento Avulso", "📋 Histórico Completo"])
            
            with tab_chamada:
                termos_lideranca = ['gerente', 'supervisor', 'encarregado', 'coordenador', 'líder', 'lider']
                colabs_operacionais = df_filtrado[
                    (df_filtrado['Status'] == 'Ativo') & 
                    (~df_filtrado['Cargo'].astype(str).str.lower().str.contains('|'.join(termos_lideranca), na=False))
                ].copy()

                if colabs_operacionais.empty:
                    st.warning("Nenhum colaborador operacional ativo no setor para chamada.")
                else:
                    data_chamada = st.date_input("Data da Chamada:", value=hoje, format="DD/MM/YYYY")
                    st.info("💡 **Instruções:** Marque **☑️ Presente** para quem veio e **🏖️ Folga** se for folga programada. Quem ficar desmarcado é contabilizado como **Ausência/Falta**.")
                    
                    faltas_existentes_data = df_faltas[
                        (df_faltas['dt_falta'] == data_chamada) & 
                        (df_faltas['Setor'] == setor_selecionado)
                    ] if not df_faltas.empty else pd.DataFrame()

                    with st.form("form_chamada_diaria"):
                        st.markdown("---")
                        
                        for i_c, (_, colab_c) in enumerate(colabs_operacionais.iterrows()):
                            nome_c = colab_c['Funcionário']
                            
                            val_pres_def = False
                            val_folga_def = False
                            
                            if not faltas_existentes_data.empty:
                                reg_colab = faltas_existentes_data[faltas_existentes_data['Funcionário'] == nome_c]
                                if not reg_colab.empty:
                                    tipo_reg = reg_colab.iloc[0].get('Tipo', '')
                                    if tipo_reg == 'Folga Concedida':
                                        val_folga_def = True
                                    else:
                                        val_pres_def = False
                                else:
                                    val_pres_def = True
                            
                            c_nome, c_pres, c_folga = st.columns([2.5, 1, 1])
                            
                            with c_nome:
                                st.markdown(f"**{nome_c}**  \n<font size=2 color='#64748B'>{colab_c.get('Cargo', 'N/A')}</font>", unsafe_allow_html=True)
                                
                            with c_pres:
                                st.checkbox("☑️ Presente", value=val_pres_def, key=f"chk_pres_{i_c}")
                                
                            with c_folga:
                                st.checkbox("🏖️ Folga", value=val_folga_def, key=f"chk_folga_{i_c}")
                            
                        btn_salvar_chamada = st.form_submit_button("💾 Salvar Chamada do Dia")
                        
                        if btn_salvar_chamada:
                            df_faltas = df_faltas[~((df_faltas['dt_falta'] == data_chamada) & (df_faltas['Setor'] == setor_selecionado))]
                            
                            novas_f = []
                            for i_c, (_, colab_c) in enumerate(colabs_operacionais.iterrows()):
                                nome_c = colab_c['Funcionário']
                                esteve_presente = st.session_state.get(f"chk_pres_{i_c}", False)
                                esta_de_folga = st.session_state.get(f"chk_folga_{i_c}", False)
                                
                                if esta_de_folga:
                                    novas_f.append({
                                        "Matricula": str(colab_c.get('Matricula', '')),
                                        "Funcionário": nome_c,
                                        "Setor": colab_c.get('Setor', ''),
                                        "Data": data_chamada.strftime('%d/%m/%Y'),
                                        "Tipo": "Folga Concedida",
                                        "Dias": 1,
                                        "CID": "-",
                                        "Motivo": "Folga Programada / Escala",
                                        "dt_falta": data_chamada
                                    })
                                elif not esteve_presente:
                                    novas_f.append({
                                        "Matricula": str(colab_c.get('Matricula', '')),
                                        "Funcionário": nome_c,
                                        "Setor": colab_c.get('Setor', ''),
                                        "Data": data_chamada.strftime('%d/%m/%Y'),
                                        "Tipo": "Ausência / A Confirmar",
                                        "Dias": 1,
                                        "CID": "-",
                                        "Motivo": "Não compareceu no turno",
                                        "dt_falta": data_chamada
                                    })

                            if novas_f:
                                df_faltas = pd.concat([df_faltas, pd.DataFrame(novas_f)], ignore_index=True)
                                salvar_faltas(df_faltas)
                                st.toast("✅ Chamada registrada com sucesso!", icon="📝")
                            else:
                                salvar_faltas(df_faltas)
                                st.toast("✅ Chamada gravada! 100% de presença no turno.", icon="🎉")
                            st.rerun()

                    faltas_da_chamada = df_faltas_filtrado[df_faltas_filtrado['dt_falta'] == data_chamada] if not df_faltas_filtrado.empty else pd.DataFrame()
                    qtd_f_ch = len(faltas_da_chamada[faltas_da_chamada['Tipo'] != 'Folga Concedida'])
                    qtd_folgas_ch = len(faltas_da_chamada[faltas_da_chamada['Tipo'] == 'Folga Concedida'])
                    qtd_p_ch = max(0, len(colabs_operacionais) - qtd_f_ch - qtd_folgas_ch)
                    
                    st.markdown("---")
                    st.markdown("##### 📲 Resumo Formatado para Envio via WhatsApp / Grupo de Trabalho:")
                    
                    txt_wa = f"📊 *RESUMO DE PRESENÇA - TROPICAL DISTRIBUIDORA*\n"
                    txt_wa += f"📅 *Data:* {data_chamada.strftime('%d/%m/%Y')} | *Setor:* {setor_selecionado}\n"
                    txt_wa += f"🟢 *Presentes:* {qtd_p_ch} colaboradores\n"
                    if qtd_folgas_ch > 0:
                        txt_wa += f"🏖️ *Folgas:* {qtd_folgas_ch} colaboradores\n"
                    txt_wa += f"🔴 *Ausentes/Faltas:* {qtd_f_ch} colaboradores\n\n"
                    
                    if not faltas_da_chamada.empty:
                        txt_wa += "*Detalhe do Turno:*\n"
                        for _, f_row in faltas_da_chamada.iterrows():
                            txt_wa += f"• {f_row['Funcionário']} ({f_row.get('Tipo', 'Ausência')}) - {f_row.get('Motivo', '-')}\n"
                    else:
                        txt_wa += "✨ *Turno com 100% de assiduidade!*\n"
                        
                    st.code(txt_wa, language="markdown")

            with tab_avulso:
                with st.form("form_falta_avulsa", clear_on_submit=True):
                    colabs_lista = sorted(df_filtrado[df_filtrado['Status'].isin(['Ativo', 'Férias'])]['Funcionário'].unique()) if not df_filtrado.empty else []
                    nome_colab = st.selectbox("Selecione o Colaborador:", colabs_lista)
                    c_t1, c_t2, c_t3 = st.columns([1.5, 1, 1])
                    tipo_falta = c_t1.selectbox("Tipo de Ocorrência", ["Falta Injustificada", "Atestado Médico", "Folga Concedida"])
                    data_falta = c_t2.date_input("Data do Ocorrido", value=hoje, format="DD/MM/YYYY")
                    dias_falta = c_t3.number_input("Qtd. de Dias", min_value=1, max_value=60, value=1, step=1)
                    c_c1, c_c2 = st.columns([1, 2])
                    cid_input = c_c1.text_input("Código CID (Se atestado)", value="")
                    motivo_obs = c_c2.text_input("Observação / Motivo Geral", value="")
                    btn_salvar_avulso = st.form_submit_button("💾 Salvar Lançamento Avulso")
                    
                    if btn_salvar_avulso and nome_colab:
                        d_colab = df_filtrado[df_filtrado['Funcionário'] == nome_colab].iloc[0]
                        nova_ocorrencia = {
                            "Matricula": str(d_colab.get('Matricula', '')),
                            "Funcionário": nome_colab,
                            "Setor": d_colab.get('Setor', ''),
                            "Data": data_falta.strftime('%d/%m/%Y'),
                            "Tipo": tipo_falta,
                            "Dias": dias_falta,
                            "CID": cid_input.strip().upper() if cid_input.strip() else "-",
                            "Motivo": motivo_obs,
                            "dt_falta": data_falta
                        }
                        df_faltas = pd.concat([df_faltas, pd.DataFrame([nova_ocorrencia])], ignore_index=True)
                        salvar_faltas(df_faltas)
                        st.toast("✅ Ocorrência avulsa salva!", icon="📌")
                        st.rerun()

            with tab_hist_f:
                if df_faltas_filtrado.empty:
                    st.info("Nenhum registro cadastrado até o momento.")
                else:
                    st.dataframe(df_faltas_filtrado, use_container_width=True)

        elif menu == "👤 Ficha Individual do Colaborador":
            st.subheader("👤 Prontuário & Ficha Individual 360º")
            
            lista_todos_colabs = sorted(df['Funcionário'].dropna().unique())
            if not lista_todos_colabs:
                st.warning("Nenhum colaborador encontrado.")
            else:
                colab_sel = st.selectbox("Selecione o Colaborador para visualizar a ficha:", lista_todos_colabs)
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
                    ult_f_txt = str(r_c.get('Ultimas_Ferias')) if pd.notnull(r_c.get('Ultimas_Ferias')) else 'Nenhuma registrada'
                    
                    st.info(f"📅 **Admissão:** {dt_adm_txt} | 🎂 **Nascimento:** {dt_nasc_txt}\n\n🏖️ **Últimas Férias:** {ult_f_txt}")
                
                st.divider()
                
                f_colab = df_faltas[df_faltas['Funcionário'] == colab_sel] if not df_faltas.empty else pd.DataFrame()
                st.markdown("##### 📋 Histórico de Ausências & Ocorrências Lançadas:")
                if f_colab.empty:
                    st.success("Nenhuma ocorrência ou falta registrada para este colaborador.")
                else:
                    st.dataframe(f_colab[['Data', 'Tipo', 'Dias', 'CID', 'Motivo']], use_container_width=True)

        elif menu == "📊 Indicadores de Frequência & Absenteísmo":
            st.subheader("📊 Painel Analítico de Assiduidade & Ocorrências")
            
            if df_faltas_filtrado.empty:
                st.info("Nenhuma ocorrência registrada no período para gerar análise gráfica.")
            else:
                m1, m2, m3 = st.columns(3)
                tot_dias_perdidos = df_faltas_filtrado['Dias'].sum()
                tot_atestados = len(df_faltas_filtrado[df_faltas_filtrado['Tipo'] == 'Atestado Médico'])
                tot_faltas_injust = len(df_faltas_filtrado[df_faltas_filtrado['Tipo'] == 'Falta Injustificada'])
                
                m1.metric("Total Dias Afastados", tot_dias_perdidos)
                m2.metric("Ocorrências Atestado", tot_atestados)
                m3.metric("Faltas Injustificadas", tot_faltas_injust)
                
                st.divider()
                
                fig_setor = px.histogram(df_faltas_filtrado, x='Setor', y='Dias', color='Tipo', barmode='group', title="Total de Dias Perdidos por Setor")
                st.plotly_chart(fig_setor, use_container_width=True)

        elif menu == "Controle de Experiência (45/90 dias)":
            st.subheader(f"📋 Colaboradores em Período de Experiência e Ações - {setor_selecionado}")
            if df_apenas_exp.empty:
                st.success("Nenhum colaborador em período de experiência neste setor.")
            else:
                cols_presentes = [c for c in ['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Admissão'] if c in df_apenas_exp.columns]
                df_exibir = df_apenas_exp[cols_presentes].copy()
                df_exibir['Vencimento 45 Dias'] = df_apenas_exp['exp_45'].apply(lambda d: d.strftime('%d/%m/%Y') if pd.notnull(d) else "")
                df_exibir['Faltam (45d)'] = df_apenas_exp['dias_para_45'].apply(lambda d: f"{d} dias" if d >= 0 else "Já passou")
                df_exibir['Vencimento 90 Dias'] = df_apenas_exp['exp_90'].apply(lambda d: d.strftime('%d/%m/%Y') if pd.notnull(d) else "")
                df_exibir['Faltam (90d)'] = df_apenas_exp['dias_para_90'].apply(lambda d: f"{d} dias")
                st.dataframe(df_exibir, use_container_width=True)

        elif menu == "Escala Inteligente de Férias":
            ferias.renderizar_modulo_ferias(df)

        elif menu == "🏖️ Colaboradores em Férias" or menu == "Gestão de Férias (Histórico)":
            st.subheader(f"🏖️ Colaboradores Atual e Programados em Férias - {setor_selecionado}")
            
            tab_f_atuais, tab_f_prog = st.tabs(["🌴 Em Gozo de Férias Agora", "📅 Próximas Férias Registradas"])
            
            with tab_f_atuais:
                df_em_ferias = df_filtrado[df_filtrado['Status'] == 'Férias'].copy()
                if df_em_ferias.empty:
                    st.info("Nenhum colaborador em gozo de férias no momento neste setor.")
                else:
                    cols_f_exibir = [c for c in ['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Admissão', 'Ultimas_Ferias'] if c in df_em_ferias.columns]
                    st.dataframe(df_em_ferias[cols_f_exibir], use_container_width=True)
                    
            with tab_f_prog:
                df_com_ferias_reg = df_filtrado[df_filtrado['Ultimas_Ferias'].notnull()].copy()
                if df_com_ferias_reg.empty:
                    st.info("Nenhum registro de agendamento de férias encontrado.")
                else:
                    cols_f_prog = [c for c in ['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Ultimas_Ferias', 'Status'] if c in df_com_ferias_reg.columns]
                    st.dataframe(df_com_ferias_reg[cols_f_prog], use_container_width=True)

        elif menu == "Aniversariantes do Mês":
            st.subheader(f"🎂 Aniversariantes do Mês - {setor_selecionado}")
            meses_nomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            mes_sel_idx = st.selectbox("Selecione o Mês", range(1, 13), index=hoje.month - 1, format_func=lambda m: meses_nomes[m-1])
            if 'dt_nasc_dt' in df_filtrado.columns:
                df_niver = df_filtrado[df_filtrado['dt_nasc_dt'].dt.month == mes_sel_idx].copy()
                st.dataframe(df_niver, use_container_width=True)

        elif menu == "Cadastrar / Editar Colaborador":
            st.subheader("👥 Gestão do Cadastro de Colaboradores")
            
            tab_cad, tab_edit_colab = st.tabs(["➕ Novo Colaborador", "✏️ Editar / Inativar Colaborador"])
            
            with tab_cad:
                with st.form("form_novo_colaborador", clear_on_submit=True):
                    st.markdown("##### 📝 Dados do Novo Colaborador")
                    c_c1, c_c2 = st.columns(2)
                    c_mat = c_c1.text_input("Matrícula:")
                    c_nome = c_c2.text_input("Nome Completo:")
                    
                    c_s1, c_s2 = st.columns(2)
                    setores_opts = sorted(list(df['Setor'].dropna().unique())) if 'Setor' in df.columns else ["Geral"]
                    c_setor = c_s1.selectbox("Setor:", setores_opts)
                    c_cargo = c_s2.text_input("Cargo:")
                    
                    c_d1, c_d2, c_d3 = st.columns(3)
                    c_adm = c_d1.date_input("Data de Admissão:", value=hoje, format="DD/MM/YYYY")
                    c_nasc = c_d2.date_input("Data de Nascimento:", value=date(1995, 1, 1), format="DD/MM/YYYY")
                    c_status = c_d3.selectbox("Status Inicial:", ["Ativo", "Férias", "Afastado", "Desligado"])
                    
                    btn_salvar_colab = st.form_submit_button("💾 Salvar Colaborador")
                    
                    if btn_salvar_colab and c_nome:
                        novo_c = {
                            "Matricula": str(c_mat).strip(),
                            "Funcionário": str(c_nome).strip(),
                            "Setor": str(c_setor),
                            "Cargo": str(c_cargo).strip(),
                            "Admissão": c_adm.strftime('%d/%m/%Y'),
                            "Nascimento": c_nasc.strftime('%d/%m/%Y'),
                            "Status": str(c_status)
                        }
                        df = pd.concat([df, pd.DataFrame([novo_c])], ignore_index=True)
                        salvar_dados(df)
                        st.toast(f"✅ Colaborador '{c_nome}' cadastrado com sucesso!", icon="🎉")
                        st.rerun()

            with tab_edit_colab:
                lista_colabs_cad = sorted(df['Funcionário'].dropna().unique())
                colab_sel_edit = st.selectbox("Selecione o Colaborador para Alterar:", lista_colabs_cad)
                
                if colab_sel_edit:
                    idx_c = df[df['Funcionário'] == colab_sel_edit].index[0]
                    colab_row = df.loc[idx_c]
                    
                    with st.form("form_editar_colaborador"):
                        st.info(f"Alterando cadastro de **{colab_row['Funcionário']}**")
                        
                        e_c1, e_c2 = st.columns(2)
                        e_mat = e_c1.text_input("Matrícula:", value=str(colab_row.get('Matricula', '')))
                        e_nome = e_c2.text_input("Nome Completo:", value=str(colab_row['Funcionário']))
                        
                        e_s1, e_s2 = st.columns(2)
                        e_setor = e_s1.text_input("Setor:", value=str(colab_row.get('Setor', '')))
                        e_cargo = e_s2.text_input("Cargo:", value=str(colab_row.get('Cargo', '')))
                        
                        e_st1, e_st2 = st.columns(2)
                        opts_status = ["Ativo", "Férias", "Afastado", "Desligado"]
                        st_atual = colab_row.get('Status', 'Ativo')
                        idx_st = opts_status.index(st_atual) if st_atual in opts_status else 0
                        e_status = e_st1.selectbox("Status Atual:", opts_status, index=idx_st)
                        
                        e_ult_ferias = e_st2.text_input("Últimas Férias (DD/MM/AAAA):", value=str(colab_row.get('Ultimas_Ferias', '')) if pd.notnull(colab_row.get('Ultimas_Ferias')) else "")
                        
                        btn_salvar_edit_c = st.form_submit_button("✏️ Atualizar Cadastro")
                        
                        if btn_salvar_edit_c:
                            df.loc[idx_c, 'Matricula'] = str(e_mat).strip()
                            df.loc[idx_c, 'Funcionário'] = str(e_nome).strip()
                            df.loc[idx_c, 'Setor'] = str(e_setor).strip()
                            df.loc[idx_c, 'Cargo'] = str(e_cargo).strip()
                            df.loc[idx_c, 'Status'] = str(e_status)
                            df.loc[idx_c, 'Ultimas_Ferias'] = str(e_ult_ferias).strip() if e_ult_ferias.strip() else None
                            
                            salvar_dados(df)
                            st.toast("✅ Cadastro de colaborador atualizado!", icon="💾")
                            st.rerun()

        elif menu == "⚙️ Criar / Gerenciar Usuários":
            st.subheader("⚙️ Painel do Administrador - Gestão de Usuários & Permissões")
            df_usuarios = carregar_usuarios()
            
            tab_novo_u, tab_edit_u, tab_lista_u = st.tabs(["➕ Criar Novo Usuário", "✏️ Editar / Módulos", "📋 Lista de Acessos"])
            
            with tab_novo_u:
                with st.form("form_novo_usuario", clear_on_submit=True):
                    c_u1, c_u2 = st.columns(2)
                    nome_u = c_u1.text_input("Nome Completo do Gestor:")
                    login_u = c_u2.text_input("Login de Usuário (Ex: joao.silva):").strip().lower()
                    
                    c_e1, c_e2, c_t1 = st.columns([1.5, 1.5, 1])
                    email_u = c_e1.text_input("E-mail de Acesso:").strip().lower()
                    senha_u = c_e2.text_input("Senha de Acesso:", type="password")
                    tel_u = c_t1.text_input("WhatsApp (DDD+Num):", value="")
                    
                    c_p1, c_p2 = st.columns(2)
                    perfil_u = c_p1.selectbox("Perfil Geral:", ["Gestor", "Admin"])
                    enviar_mail_chk = c_p2.checkbox("📧 Enviar dados de acesso também por e-mail?", value=False)
                    
                    st.markdown("##### 📌 Selecione os Módulos Liberados para este Usuário:")
                    modulos_selecionados = []
                    cols_mod = st.columns(2)
                    for idx_m, mod_nome in enumerate(TODOS_MODULOS):
                        with cols_mod[idx_m % 2]:
                            default_val = True if perfil_u == "Admin" or mod_nome in ["Dashboard & Alertas", "Chamada & Faltas do Dia"] else False
                            if st.checkbox(mod_nome, value=default_val, key=f"mod_cad_{idx_m}"):
                                modulos_selecionados.append(mod_nome)
                    
                    btn_cad_u = st.form_submit_button("💾 Criar Usuário")
                    
                    if btn_cad_u and nome_u and login_u and senha_u:
                        if login_u in df_usuarios['Usuario'].astype(str).str.lower().values:
                            st.error(f"❌ O login '{login_u}' já existe!")
                        elif not modulos_selecionados:
                            st.warning("⚠️ Selecione pelo menos um módulo.")
                        else:
                            str_mods = ",".join(modulos_selecionados)
                            novo_usr = {"Nome": str(nome_u).strip(), "Usuario": str(login_u).strip(), "Email": str(email_u).strip(), "Senha": str(senha_u).strip(), "Perfil": str(perfil_u), "Modulos": str_mods, "Telefone": str(tel_u).strip()}
                            df_usuarios = pd.concat([df_usuarios, pd.DataFrame([novo_usr])], ignore_index=True)
                            salvar_usuarios(df_usuarios)
                            
                            st.toast(f"✅ Usuário '{nome_u}' criado com sucesso!", icon="🎉")
                            
                            if tel_u:
                                link_wa_novo = gerar_link_whatsapp(tel_u, nome_u, email_u if email_u else login_u, senha_u)
                                st.markdown(f"👉 **[📲 Clique aqui para enviar o acesso por WhatsApp]({link_wa_novo})**")

                            if enviar_mail_chk and email_u:
                                ok_m, msg_m = enviar_email_acesso(email_u, nome_u, email_u, senha_u)
                                if ok_m:
                                    st.toast("📧 E-mail de acesso enviado com sucesso!", icon="✉️")
                                else:
                                    st.info(f"ℹ️ {msg_m}")

            with tab_edit_u:
                lista_logins = sorted(df_usuarios['Usuario'].astype(str).unique())
                usr_sel_edit = st.selectbox("Selecione o Usuário para Editar:", lista_logins)
                
                if usr_sel_edit:
                    idx_u = df_usuarios[df_usuarios['Usuario'].astype(str) == usr_sel_edit].index[0]
                    usr_dados = df_usuarios.loc[idx_u]
                    
                    with st.form("form_edit_usr"):
                        st.info(f"Editando dados e permissões do usuário **{usr_dados['Nome']}**")
                        
                        e_u1, e_u2 = st.columns(2)
                        e_nome = e_u1.text_input("Nome Completo:", value=str(usr_dados['Nome']))
                        e_email = e_u2.text_input("E-mail:", value=str(usr_dados.get('Email', '')))
                        
                        e_s1, e_s2, e_t1 = st.columns([1.2, 1.2, 1])
                        e_senha = e_s1.text_input("Senha:", value=str(usr_dados['Senha']), type="password")
                        opts_p = ["Gestor", "Admin"]
                        idx_p = opts_p.index(usr_dados['Perfil']) if usr_dados['Perfil'] in opts_p else 0
                        e_perfil = e_s2.selectbox("Perfil Geral:", opts_p, index=idx_p)
                        e_tel = e_t1.text_input("WhatsApp (DDD+Num):", value=str(usr_dados.get('Telefone', '')))
                        
                        reenviar_mail_chk = st.checkbox("📧 Reenviar e-mail com os novos dados de acesso?", value=False)
                        
                        st.markdown("##### 📌 Módulos Liberados:")
                        mods_atuais = [m.strip() for m in str(usr_dados.get('Modulos', '')).split(',') if m.strip()]
                        e_modulos = []
                        cols_e_mod = st.columns(2)
                        for idx_m, mod_nome in enumerate(TODOS_MODULOS):
                            with cols_e_mod[idx_m % 2]:
                                is_chk = mod_nome in mods_atuais
                                if st.checkbox(mod_nome, value=is_chk, key=f"mod_edit_{idx_m}"):
                                    e_modulos.append(mod_nome)
                                    
                        btn_salvar_edit = st.form_submit_button("✏️ Atualizar Usuário e Permissões")
                        
                        if btn_salvar_edit:
                            df_usuarios.loc[idx_u, 'Nome'] = str(e_nome).strip()
                            df_usuarios.loc[idx_u, 'Email'] = str(e_email).strip()
                            df_usuarios.loc[idx_u, 'Senha'] = str(e_senha).strip()
                            df_usuarios.loc[idx_u, 'Perfil'] = str(e_perfil)
                            df_usuarios.loc[idx_u, 'Modulos'] = ",".join(e_modulos)
                            df_usuarios.loc[idx_u, 'Telefone'] = str(e_tel).strip()
                            salvar_usuarios(df_usuarios)
                            
                            if str(usr_dados['Usuario']).lower() == str(st.session_state.get("usuario_login")).lower():
                                st.session_state["usuario_nome"] = str(e_nome).strip()
                                st.session_state["perfil"] = str(e_perfil)
                                st.session_state["usuario_modulos"] = e_modulos

                            st.toast(f"✅ Permissões de '{e_nome}' atualizadas!", icon="💾")
                            
                            if e_tel:
                                link_wa_edit = gerar_link_whatsapp(e_tel, e_nome, e_email if e_email else e_nome, e_senha)
                                st.markdown(f"👉 **[📲 Clique aqui para enviar os novos dados via WhatsApp]({link_wa_edit})**")

                            if reenviar_mail_chk and e_email:
                                ok_m, msg_m = enviar_email_acesso(e_email, e_nome, e_email, e_senha)
                                if ok_m:
                                    st.toast("📧 E-mail atualizado enviado!", icon="✉️")
                                else:
                                    st.info(f"ℹ️ {msg_m}")

            with tab_lista_u:
                st.markdown("##### 👥 Usuários e Módulos Cadastrados:")
                st.dataframe(df_usuarios[['Nome', 'Usuario', 'Email', 'Telefone', 'Perfil', 'Modulos']], use_container_width=True)
                
                st.markdown("---")
                st.markdown("##### 📲 Envio Rápido de Convite por WhatsApp")
                usr_wa_sel = st.selectbox("Selecione o usuário para enviar acesso por WhatsApp:", df_usuarios['Usuario'].dropna().unique(), key="sel_wa_u")
                if usr_wa_sel:
                    u_row = df_usuarios[df_usuarios['Usuario'] == usr_wa_sel].iloc[0]
                    u_tel = u_row.get('Telefone', '')
                    if u_tel and u_tel != 'nan':
                        l_wa_direto = gerar_link_whatsapp(u_tel, u_row['Nome'], u_row['Email'] if u_row['Email'] else u_row['Usuario'], u_row['Senha'])
                        st.markdown(f"👉 **[📲 Abrir WhatsApp e Enviar Acesso para {u_row['Nome']}]({l_wa_direto})**")
                    else:
                        st.info("⚠️ Este usuário não possui um número de WhatsApp cadastrado.")

                st.markdown("---")
                st.markdown("##### 🗑️ Excluir Acesso de Usuário")
                usr_del = st.selectbox("Selecione o usuário para remover:", df_usuarios['Usuario'].dropna().unique(), key="sel_del_u")
                if st.button("❌ Excluir Usuário Selecionado", type="primary"):
                    if usr_del == "admin":
                        st.error("⚠️ O usuário padrão 'admin' não pode ser excluído.")
                    else:
                        df_usuarios = df_usuarios[df_usuarios['Usuario'] != usr_del].reset_index(drop=True)
                        salvar_usuarios(df_usuarios)
                        st.toast(f"Acesso de '{usr_del}' excluído!", icon="🗑️")
                        st.rerun()

        elif menu == "📥 Importar Nova Base":
            st.subheader("📥 Atualizar Base Geral de Colaboradores (.xlsx)")
            arquivo_upload = st.file_uploader("Arraste ou selecione o arquivo .xlsx", type=["xlsx"])
            if arquivo_upload is not None and st.button("Confirmar e Substituir Base"):
                df_novo_up = pd.read_excel(arquivo_upload)
                df_novo_up.to_excel(ARQUIVO_DADOS, index=False)
                st.toast("Nova base importada com sucesso!", icon="📥")
                st.rerun()
