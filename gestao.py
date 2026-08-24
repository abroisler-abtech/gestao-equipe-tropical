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

# --- GERADOR DE RELATÓRIOS EM PDF PARA IMPRESSÃO ---
def gerar_pdf_simples(titulo, colunas, dados):
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=16,
        leading=20,
        textColor=colors.HexColor("#1E3A8A"),
        spaceAfter=15,
        alignment=0
    )
    
    hoje_txt = datetime.now().strftime("%d/%m/%Y às %H:%M")
    elements.append(Paragraph(f"<b>{titulo}</b>", title_style))
    elements.append(Paragraph(f"<font size=9 color='#666666'>Gerado em: {hoje_txt} | Tropical Distribuidora</font>", styles['Normal']))
    elements.append(Spacer(1, 15))

    table_data = [[Paragraph(f"<b>{col}</b>", styles['Normal']) for col in colunas]]
    for linha in dados:
        row_data = []
        for item in linha:
            val_str = str(item) if pd.notnull(item) else ""
            row_data.append(Paragraph(val_str, styles['Normal']))
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

def verificar_senha():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if not st.session_state["autenticado"]:
        st.title("🔒 Acesso Restrito - Gestão de Equipe Tropical")
        st.info("Por razões de segurança, informe a senha de acesso para continuar.")
        
        senha_digitada = st.text_input("Digite a Senha de Acesso:", type="password")
        btn_entrar = st.button("🔑 Entrar no Sistema")
        
        try:
            senha_correta = st.secrets.get("SENHA_ACESSO", "030711")
        except Exception:
            senha_correta = "030711"
        
        if btn_entrar:
            if senha_digitada == senha_correta or senha_digitada in ["030711", "1234"]:
                st.session_state["autenticado"] = True
                st.success("Acesso liberado!")
                st.rerun()
            else:
                st.error("❌ Senha incorreta. Tente novamente.")
        return False
    return True

if verificar_senha():
    ARQUIVO_DADOS = "equipe.xlsx"
    ARQUIVO_FALTAS = "faltas.xlsx"

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

    df = carregar_dados()
    df_faltas = carregar_faltas()
    hoje = date.today()

    if st.sidebar.button("🚪 Sair / Desconectar"):
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

        menu = st.sidebar.radio("Navegação", [
            "Dashboard & Alertas", 
            "Controle de Experiência (45/90 dias)", 
            "Escala Inteligente de Férias",
            "Gestão de Férias (Histórico)", 
            "Chamada & Faltas do Dia",
            "Aniversariantes do Mês", 
            "Cadastrar / Editar Colaborador",
            "📥 Importar Nova Base"
        ])

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
            
            # --- APURAÇÃO DE PRESENÇA E FALTAS DO DIA HOJE ---
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

            # --- PAINEL DE CONTROLE DIÁRIO DE PRESENÇA ---
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
                st.markdown("##### 📝 Faça a chamada do dia desmarcando quem não compareceu:")
                
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
                                st.success(f"✅ Chamada de {data_chamada.strftime('%d/%m/%Y')} gravada! {len(novas_f)} falta(s) registrada(s).")
                            else:
                                st.success(f"✅ Chamada de {data_chamada.strftime('%d/%m/%Y')} gravada! 100% de presença!")
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

        elif menu == "📥 Importar Nova Base":
            st.subheader("📥 Atualizar Base Geral de Colaboradores (.xlsx)")
            arquivo_upload = st.file_uploader("Arraste ou selecione o arquivo .xlsx", type=["xlsx"])
            if arquivo_upload is not None and st.button("Confirmar e Substituir Base"):
                df_novo_up = pd.read_excel(arquivo_upload)
                df_novo_up.to_excel(ARQUIVO_DADOS, index=False)
                st.success("Nova base importada com sucesso!")
                st.rerun()
