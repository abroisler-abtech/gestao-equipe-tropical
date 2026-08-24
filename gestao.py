import io
import os
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
    "Controle de Experiência (45/90 dias)",
    "Escala Inteligente de Férias",
    "Gestão de Férias (Histórico)",
    "Aniversariantes do Mês",
    "Cadastrar / Editar Colaborador",
    "⚙️ Criar / Gerenciar Usuários",
    "📥 Importar Nova Base"
]

# --- GERENCIAMENTO DE USUÁRIOS E PERMISSÕES ---
def carregar_usuarios():
    if os.path.exists(ARQUIVO_USUARIOS):
        df_u = pd.read_excel(ARQUIVO_USUARIOS)
        df_u.columns = df_u.columns.str.strip()
        if 'Modulos' not in df_u.columns:
            df_u['Modulos'] = ",".join(TODOS_MODULOS)
        return df_u
    else:
        dados_iniciais = [
            {"Nome": "Administrador", "Usuario": "admin", "Senha": "123", "Perfil": "Admin", "Modulos": ",".join(TODOS_MODULOS)},
            {"Nome": "Gestor de Turno", "Usuario": "gestor", "Senha": "123", "Perfil": "Gestor", "Modulos": "Dashboard & Alertas,Chamada & Faltas do Dia"}
        ]
        df_u = pd.DataFrame(dados_iniciais)
        df_u.to_excel(ARQUIVO_USUARIOS, index=False)
        return df_u

def salvar_usuarios(df_u):
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
    elements.append(Paragraph(f"<font size=9 color='#666666'>Gerado em: {hoje_txt} | Tropical Distribuidora</font>", styles['Normal']))
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
    elements.append(Paragraph(f"<b>Setor Filtrado:</b> {setor_nome} | <b>Gerado em:</b> {hoje_txt} | Tropical Distribuidora", sub_style))
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

# --- AUTENTICAÇÃO DINÂMICA ---
def verificar_senha():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False
        st.session_state["perfil"] = None
        st.session_state["usuario_nome"] = None
        st.session_state["usuario_login"] = None
        st.session_state["usuario_modulos"] = []

    if not st.session_state["autenticado"]:
        st.title("🔒 Acesso Restrito - Gestão de Equipe Tropical")
        st.info("Informe seu usuário e senha para entrar no sistema.")
        
        df_u = carregar_usuarios()
        
        user_input = st.text_input("Usuário:").strip().lower()
        senha_input = st.text_input("Senha:", type="password")
        btn_entrar = st.button("🔑 Entrar no Sistema")
        
        if btn_entrar:
            match = df_u[(df_u['Usuario'].astype(str).str.lower() == user_input) & (df_u['Senha'].astype(str) == senha_input)]
            if not match.empty:
                usr = match.iloc[0]
                st.session_state["autenticado"] = True
                st.session_state["perfil"] = usr['Perfil']
                st.session_state["usuario_nome"] = usr['Nome']
                st.session_state["usuario_login"] = usr['Usuario']
                
                mods_raw = str(usr.get('Modulos', ''))
                st.session_state["usuario_modulos"] = [m.strip() for m in mods_raw.split(',') if m.strip()] if mods_raw else TODOS_MODULOS
                
                st.success(f"Acesso liberado! Bem-vindo(a), {usr['Nome']}")
                st.rerun()
            else:
                st.error("❌ Usuário ou senha incorretos.")
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
                if str(s_atual) != str(senha_correta):
                    st.error("❌ Senha atual incorreta!")
                elif not s_nova:
                    st.warning("⚠️ Digite a nova senha.")
                elif s_nova != s_conf:
                    st.error("❌ A nova senha e a confirmação não conferem.")
                else:
                    df_u.loc[mask, 'Senha'] = str(s_nova)
                    salvar_usuarios(df_u)
                    st.success("✅ Senha alterada com sucesso!")
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
            st.subheader("⚠️ Painel Geral")
            
            df_ativos = df_filtrado[df_filtrado['Status'] == 'Ativo']
            df_ferias_st = df_filtrado[df_filtrado['Status'] == 'Férias']
            df_afastados = df_filtrado[df_filtrado['Status'].astype(str).str.contains('Atestado|Afastado|INSS|Licença|licenca', case=False, na=False)]
            
            faltas_hoje = df_faltas_filtrado[df_faltas_filtrado['dt_falta'] == hoje] if not df_faltas_filtrado.empty else pd.DataFrame()
            qtd_faltantes_hoje = len(faltas_hoje)
            qtd_presentes_hoje = max(0, len(df_ativos) - qtd_faltantes_hoje)

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

            p1, p2 = st.columns(2)
            p1.metric("🟢 Presentes Hoje", qtd_presentes_hoje)
            p2.metric("🔴 Faltantes / Ausentes Hoje", qtd_faltantes_hoje)

            if not faltas_hoje.empty:
                st.error("🚨 **JANELA DE AUSÊNCIAS/FALTANTES DO DIA HOJE:**")
                cols_f_hoje = st.columns(min(len(faltas_hoje), 3))
                for i_fh, (_, fh) in enumerate(faltas_hoje.iterrows()):
                    with cols_f_hoje[i_fh % 3]:
                        st.warning(f"👤 **{fh['Funcionário']}**\n\n**Setor:** {fh.get('Setor', 'N/A')}\n\n**Tipo:** {fh.get('Tipo', 'Falta')}\n\n**Motivo:** {fh.get('Motivo', '-')}")

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
            if c5.button("🔍 Ver Faltas Hoje", key="btn_faltas_hoje"):
                cols_m = [c for c in ['Data', 'Funcionário', 'Setor', 'Tipo', 'Dias', 'CID', 'Motivo'] if c in faltas_hoje.columns]
                exibir_modal_detalhes(f"Colaboradores Ausentes em {hoje.strftime('%d/%m/%Y')}", faltas_hoje[cols_m] if not faltas_hoje.empty else pd.DataFrame())

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
                    fig_faltas = px.bar(df_tipo_falta, x='Tipo', y='Dias', title="Total de Dias Afastados por Tipo (Geral)", text_auto=True, color='Tipo')
                    st.plotly_chart(fig_faltas, use_container_width=True)
                else:
                    st.info("Sem dados de ocorrências para gerar o gráfico de ausências.")

        elif menu == "Chamada & Faltas do Dia":
            st.subheader(f"📌 Chamada Diária de Presença & Ocorrências - {setor_selecionado}")
            
            tab_chamada, tab_avulso, tab_hist_f = st.tabs(["☑️ Chamada Diária (Presença)", "➕ Lançamento Avulso", "📋 Histórico Completo"])
            
            with tab_chamada:
                colabs_ativos = df_filtrado[df_filtrado['Status'] == 'Ativo'].copy()
                if colabs_ativos.empty:
                    st.warning("Nenhum colaborador ativo no setor para chamada.")
                else:
                    data_chamada = st.date_input("Data da Chamada:", value=hoje, format="DD/MM/YYYY")
                    
                    with st.form("form_chamada_diaria"):
                        st.markdown("---")
                        presencas = {}
                        motivos_falta = {}
                        
                        for i_c, (_, colab_c) in enumerate(colabs_ativos.iterrows()):
                            nome_c = colab_c['Funcionário']
                            c_pres, c_tipo_f, c_obs_f = st.columns([1.5, 1.2, 1.5])
                            
                            with c_pres:
                                is_pres = st.checkbox(f"**{nome_c}** ({colab_c.get('Cargo', 'N/A')})", value=True, key=f"chk_{i_c}")
                                presencas[nome_c] = is_pres
                                
                            with c_tipo_f:
                                if not is_pres:
                                    tp_falta = st.selectbox("Tipo", ["Falta Injustificada", "Atestado Médico", "Folga Concedida"], key=f"tp_{i_c}")
                                else:
                                    tp_falta = None
                                    
                            with c_obs_f:
                                if not is_pres:
                                    obs_f = st.text_input("Observação/Motivo", value="", key=f"obs_{i_c}")
                                else:
                                    obs_f = ""
                                    
                            motivos_falta[nome_c] = (tp_falta, obs_f)
                            
                        btn_salvar_chamada = st.form_submit_button("💾 Salvar Chamada do Dia")
                        
                        if btn_salvar_chamada:
                            novas_f = []
                            for nome_c, esteve_presente in presencas.items():
                                if not esteve_presente:
                                    tp_f, obs_f = motivos_falta[nome_c]
                                    d_colab = colabs_ativos[colabs_ativos['Funcionário'] == nome_c].iloc[0]
                                    novas_f.append({
                                        "Matricula": str(d_colab.get('Matricula', '')),
                                        "Funcionário": nome_c,
                                        "Setor": d_colab.get('Setor', ''),
                                        "Data": data_chamada.strftime('%d/%m/%Y'),
                                        "Tipo": tp_f,
                                        "Dias": 1,
                                        "CID": "-",
                                        "Motivo": obs_f if obs_f else "Registrado pela Chamada Diária",
                                        "dt_falta": data_chamada
                                    })
                            if novas_f:
                                df_faltas = pd.concat([df_faltas, pd.DataFrame(novas_f)], ignore_index=True)
                                salvar_faltas(df_faltas)
                                st.success(f"✅ Chamada gravada com {len(novas_f)} falta(s) registrada(s)!")
                            else:
                                st.success("✅ Chamada gravada! 100% de presença no turno.")
                            st.rerun()

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
                        st.success("Ocorrência avulsa salva!")
                        st.rerun()

            with tab_hist_f:
                if df_faltas_filtrado.empty:
                    st.info("Nenhum registro cadastrado até o momento.")
                else:
                    st.dataframe(df_faltas_filtrado, use_container_width=True)

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

        elif menu == "Gestão de Férias (Histórico)":
            st.subheader(f"🏖️ Controle de Períodos Aquisitivos e Concessivos - {setor_selecionado}")
            lista_ferias = []
            for idx_h, r in df_filtrado[df_filtrado['Status'].isin(['Ativo', 'Férias'])].iterrows():
                adm = r['dt_adm']
                ult_ferias = pd.to_datetime(r.get('Ultimas_Ferias'), dayfirst=True, errors='coerce').date() if pd.notnull(r.get('Ultimas_Ferias')) else None
                data_base = ult_ferias if ult_ferias else adm
                if pd.notnull(data_base):
                    anos = (hoje - data_base).days // 365
                    inicio_aq = data_base + timedelta(days=365 * max(0, anos))
                    fim_aq = inicio_aq + timedelta(days=365)
                    limite_conc = fim_aq + timedelta(days=365)
                    lista_ferias.append({
                        "Matrícula": r.get('Matricula', 'N/A'),
                        "Funcionário": r['Funcionário'],
                        "Setor": r.get('Setor', 'N/A'),
                        "Admissão": r.get('Admissão', 'N/A'),
                        "Status Atual": r['Status'],
                        "Últimas Férias": str(r.get('Ultimas_Ferias')) if pd.notnull(r.get('Ultimas_Ferias')) else 'Nunca registradas',
                        "Limite Concessivo": limite_conc.strftime('%d/%m/%Y')
                    })
            st.dataframe(pd.DataFrame(lista_ferias), use_container_width=True)

        elif menu == "Aniversariantes do Mês":
            st.subheader(f"🎂 Aniversariantes do Mês - {setor_selecionado}")
            meses_nomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            mes_sel_idx = st.selectbox("Selecione o Mês", range(1, 13), index=hoje.month - 1, format_func=lambda m: meses_nomes[m-1])
            if 'dt_nasc_dt' in df_filtrado.columns:
                df_niver = df_filtrado[df_filtrado['dt_nasc_dt'].dt.month == mes_sel_idx].copy()
                st.dataframe(df_niver, use_container_width=True)

        elif menu == "Cadastrar / Editar Colaborador":
            st.subheader("👥 Gestão do Cadastro de Colaboradores")

        elif menu == "⚙️ Criar / Gerenciar Usuários":
            st.subheader("⚙️ Painel do Administrador - Gestão de Usuários & Permissões")
            df_usuarios = carregar_usuarios()
            
            tab_novo_u, tab_edit_u, tab_lista_u = st.tabs(["➕ Criar Novo Usuário", "✏️ Editar / Módulos", "📋 Lista de Acessos"])
            
            with tab_novo_u:
                with st.form("form_novo_usuario", clear_on_submit=True):
                    c_u1, c_u2 = st.columns(2)
                    nome_u = c_u1.text_input("Nome do Gestor / Colaborador:")
                    login_u = c_u2.text_input("Login de Acesso (Ex: joao.silva):").strip().lower()
                    
                    c_p1, c_p2 = st.columns(2)
                    senha_u = c_p1.text_input("Senha de Acesso:", type="password")
                    perfil_u = c_p2.selectbox("Perfil Geral:", ["Gestor", "Admin"])
                    
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
                            st.error(f"❌ O login '{login_u}' já existe! Escolha outro login.")
                        elif not modulos_selecionados:
                            st.warning("⚠️ Selecione pelo menos um módulo para liberar o acesso.")
                        else:
                            str_mods = ",".join(modulos_selecionados)
                            novo_usr = {"Nome": nome_u.strip(), "Usuario": login_u, "Senha": senha_u.strip(), "Perfil": perfil_u, "Modulos": str_mods}
                            df_usuarios = pd.concat([df_usuarios, pd.DataFrame([novo_usr])], ignore_index=True)
                            salvar_usuarios(df_usuarios)
                            st.success(f"✅ Usuário '{login_u}' criado com {len(modulos_selecionados)} módulo(s) liberado(s)!")
                            st.rerun()

            with tab_edit_u:
                lista_logins = sorted(df_usuarios['Usuario'].astype(str).unique())
                usr_sel_edit = st.selectbox("Selecione o Usuário para Editar:", lista_logins)
                
                if usr_sel_edit:
                    idx_u = df_usuarios[df_usuarios['Usuario'].astype(str) == usr_sel_edit].index[0]
                    usr_dados = df_usuarios.loc[idx_u]
                    
                    with st.form("form_edit_usr"):
                        st.info(f"Editando dados e permissões do usuário **{usr_dados['Usuario']}**")
                        
                        e_u1, e_u2 = st.columns(2)
                        e_nome = e_u1.text_input("Nome:", value=str(usr_dados['Nome']))
                        e_senha = e_u2.text_input("Senha:", value=str(usr_dados['Senha']), type="password")
                        
                        e_p1, _ = st.columns(2)
                        opts_p = ["Gestor", "Admin"]
                        idx_p = opts_p.index(usr_dados['Perfil']) if usr_dados['Perfil'] in opts_p else 0
                        e_perfil = e_p1.selectbox("Perfil Geral:", opts_p, index=idx_p)
                        
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
                            df_usuarios.loc[idx_u, 'Nome'] = e_nome.strip()
                            df_usuarios.loc[idx_u, 'Senha'] = e_senha.strip()
                            df_usuarios.loc[idx_u, 'Perfil'] = e_perfil
                            df_usuarios.loc[idx_u, 'Modulos'] = ",".join(e_modulos)
                            salvar_usuarios(df_usuarios)
                            st.success(f"✅ Usuário '{usr_sel_edit}' atualizado com sucesso!")
                            st.rerun()

            with tab_lista_u:
                st.markdown("##### 👥 Usuários e Módulos Cadastrados:")
                st.dataframe(df_usuarios[['Nome', 'Usuario', 'Perfil', 'Modulos']], use_container_width=True)
                
                st.markdown("---")
                st.markdown("##### 🗑️ Excluir Acesso de Usuário")
                usr_del = st.selectbox("Selecione o usuário para remover:", df_usuarios['Usuario'].dropna().unique(), key="sel_del_u")
                if st.button("❌ Excluir Usuário Selecionado", type="primary"):
                    if usr_del == "admin":
                        st.error("⚠️ O usuário padrão 'admin' não pode ser excluído.")
                    else:
                        df_usuarios = df_usuarios[df_usuarios['Usuario'] != usr_del].reset_index(drop=True)
                        salvar_usuarios(df_usuarios)
                        st.success(f"Acesso do usuário '{usr_del}' excluído com sucesso!")
                        st.rerun()

        elif menu == "📥 Importar Nova Base":
            st.subheader("📥 Atualizar Base Geral de Colaboradores (.xlsx)")
            arquivo_upload = st.file_uploader("Arraste ou selecione o arquivo .xlsx", type=["xlsx"])
            if arquivo_upload is not None and st.button("Confirmar e Substituir Base"):
                df_novo_up = pd.read_excel(arquivo_upload)
                df_novo_up.to_excel(ARQUIVO_DADOS, index=False)
                st.success("Nova base importada com sucesso!")
                st.rerun()
