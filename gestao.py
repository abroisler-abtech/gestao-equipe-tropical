import io
import os
from datetime import datetime, date, timedelta
import importlib
import ferias
import pandas as pd
import plotly.express as px
import streamlit as st

importlib.reload(ferias)

# --- FUNÇÃO DE EXPORTAÇÃO PARA EXCEL ---
def converter_para_excel(df_exp, nome_aba='Base_Tropical'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_exp.to_excel(writer, index=False, sheet_name=nome_aba)
    return output.getvalue()

# --- SISTEMA DE AUTENTICAÇÃO POR SENHA ---
def verificar_senha():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if not st.session_state["autenticado"]:
        st.title("🔒 Acesso Restrito - Gestão de Equipe Tropical")
        st.info("Informe a senha de acesso para continuar.")

        senha_digitada = st.text_input("Digite a Senha de Acesso:", type="password")
        btn_entrar = st.button("🔑 Entrar no Sistema")

        senha_correta = "030711"

        if btn_entrar:
            if senha_digitada == senha_correta:
                st.session_state["autenticado"] = True
                st.success("Acesso liberado!")
                st.rerun()
            else:
                st.error("❌ Senha incorreta. Tente novamente.")
        return False
    return True


if verificar_senha():
    ARQUIVO_DADOS = "equipe.xlsx"
    ARQUIVO_OCORRENCIAS = "ocorrencias.xlsx"

    @st.cache_data(ttl=5)
    def carregar_dados():
        if os.path.exists(ARQUIVO_DADOS):
            df = pd.read_excel(ARQUIVO_DADOS)
            df.columns = [str(c).strip() for c in df.columns]

            # Mapeamento do Nome do Colaborador
            col_nome = next((c for c in df.columns if 'func' in c.lower() or 'nome' in c.lower()), 'Funcionário')
            if col_nome in df.columns and col_nome != 'Funcionário':
                df['Funcionário'] = df[col_nome]

            # Mapeamento da Admissão
            col_adm = next((c for c in df.columns if 'admiss' in c.lower() or 'dt_adm' in c.lower()), None)
            if col_adm:
                df['dt_adm'] = pd.to_datetime(df[col_adm], dayfirst=True, errors='coerce')
            else:
                df['dt_adm'] = pd.NaT

            # Mapeamento do Nascimento
            col_nasc = next((c for c in df.columns if 'nasc' in c.lower() or 'anivers' in c.lower()), None)
            if col_nasc:
                df['dt_nascimento'] = pd.to_datetime(df[col_nasc], dayfirst=True, errors='coerce')
            else:
                df['dt_nascimento'] = pd.NaT

            # Mapeamento do Status
            col_status = next((c for c in df.columns if 'status' in c.lower()), None)
            if col_status and col_status != 'Status':
                df['Status'] = df[col_status]
            elif 'Status' not in df.columns:
                df['Status'] = 'Ativo'
            
            df['Status'] = df['Status'].fillna('Ativo').astype(str)

            # Mapeamento do Setor
            if 'Setor' not in df.columns:
                col_set = next((c for c in df.columns if 'setor' in c.lower()), None)
                df['Setor'] = df[col_set] if col_set else 'Geral'

            return df
        return pd.DataFrame()

    @st.cache_data(ttl=5)
    def carregar_ocorrencias():
        if os.path.exists(ARQUIVO_OCORRENCIAS):
            df_oc = pd.read_excel(ARQUIVO_OCORRENCIAS)
            df_oc.columns = [str(c).strip() for c in df_oc.columns]
            if 'Data_Inicio' in df_oc.columns:
                df_oc['Data_Inicio'] = pd.to_datetime(df_oc['Data_Inicio'], dayfirst=True, errors='coerce').dt.date
            if 'Data_Fim' in df_oc.columns:
                df_oc['Data_Fim'] = pd.to_datetime(df_oc['Data_Fim'], dayfirst=True, errors='coerce').dt.date
            return df_oc
        return pd.DataFrame(columns=["Data_Registro", "Funcionário", "Setor", "Tipo_Ocorrencia", "Data_Inicio", "Data_Fim", "Dias", "Motivo_Observacao"])

    df = carregar_dados()
    df_oc = carregar_ocorrencias()

    if df.empty:
        st.error("⚠️ Nenhuma base de dados encontrada ou arquivo `equipe.xlsx` vazio.")
    else:
        hoje = date.today()

        # Classificação do Status
        status_str = df['Status'].str.lower()
        is_ferias = status_str.str.contains('férias|ferias')
        is_inss = status_str.str.contains('inss|afastado|licença|licenca')

        df_ativos = df[~is_ferias & ~is_inss].copy()
        df_em_ferias = df[is_ferias].copy()
        df_inss = df[is_inss].copy()

        # Ocorrências Ativas
        if not df_oc.empty and 'Data_Inicio' in df_oc.columns and 'Data_Fim' in df_oc.columns:
            oc_hoje = df_oc[(df_oc['Data_Inicio'] <= hoje) & (df_oc['Data_Fim'] >= hoje)].copy()
        else:
            oc_hoje = pd.DataFrame()

        # Experiência (45/90d)
        if 'dt_adm' in df_ativos.columns and not df_ativos['dt_adm'].isna().all():
            df_ativos['Venc_45_dias'] = df_ativos['dt_adm'].dt.date + timedelta(days=45)
            df_ativos['Venc_90_dias'] = df_ativos['dt_adm'].dt.date + timedelta(days=90)
            df_exp = df_ativos[df_ativos['dt_adm'].dt.date >= (hoje - timedelta(days=90))].copy()
        else:
            df_exp = pd.DataFrame()

        # Aniversariantes do Dia e do Mês
        if 'dt_nascimento' in df.columns and not df['dt_nascimento'].isna().all():
            df_valid_nasc = df.dropna(subset=['dt_nascimento']).copy()
            aniv_hoje = df_valid_nasc[
                (df_valid_nasc['dt_nascimento'].dt.day == hoje.day) & 
                (df_valid_nasc['dt_nascimento'].dt.month == hoje.month)
            ].copy()
            df_aniversariantes = df_valid_nasc[df_valid_nasc['dt_nascimento'].dt.month == hoje.month].copy()
        else:
            aniv_hoje = pd.DataFrame()
            df_aniversariantes = pd.DataFrame()

        # --- MENU LATERAL DE NAVEGAÇÃO ---
        st.sidebar.title("🌴 Gestão Tropical")
        st.sidebar.metric("👷 Operação Ativa", f"{len(df_ativos)} colab.")
        st.sidebar.metric("🏖️ Em Férias Hoje", f"{len(df_em_ferias)} colab.")
        st.sidebar.metric("🏥 Afastados (INSS)", f"{len(df_inss)} colab.")

        setores = ["Todos"] + list(df['Setor'].dropna().unique())
        setor_selecionado = st.sidebar.selectbox("Filtrar Setor:", setores)

        opcoes_menu = [
            "🏠 Dashboard Principal",
            "🏥 Janela - Afastados (INSS)",
            "🏖️ Janela - Equipe em Férias",
            "📋 Ocorrências (Faltas/Atestados/Folgas)",
            "🎂 Aniversariantes do Mês",
            "⏳ Contratos de Experiência (45/90d)",
            "📅 Escala Inteligente de Férias",
            "👥 Cadastrar / Editar Colaborador"
        ]

        menu = st.sidebar.radio("Selecione o Módulo:", opcoes_menu, key="modulo_radio")

        # Filtro de Setor
        if setor_selecionado != "Todos":
            df_f = df[df['Setor'] == setor_selecionado]
            df_ativos_f = df_ativos[df_ativos['Setor'] == setor_selecionado]
            df_ferias_f = df_em_ferias[df_em_ferias['Setor'] == setor_selecionado]
            df_inss_f = df_inss[df_inss['Setor'] == setor_selecionado]
            df_exp_f = df_exp[df_exp['Setor'] == setor_selecionado] if not df_exp.empty else pd.DataFrame()
            df_aniv_f = df_aniversariantes[df_aniversariantes['Setor'] == setor_selecionado] if not df_aniversariantes.empty else pd.DataFrame()
            aniv_hoje_f = aniv_hoje[aniv_hoje['Setor'] == setor_selecionado] if not aniv_hoje.empty else pd.DataFrame()
            oc_hoje_f = oc_hoje[oc_hoje['Setor'] == setor_selecionado] if not oc_hoje.empty else pd.DataFrame()
            df_oc_f = df_oc[df_oc['Setor'] == setor_selecionado] if not df_oc.empty else pd.DataFrame()
        else:
            df_f = df
            df_ativos_f = df_ativos
            df_ferias_f = df_em_ferias
            df_inss_f = df_inss
            df_exp_f = df_exp
            df_aniv_f = df_aniversariantes
            aniv_hoje_f = aniv_hoje
            oc_hoje_f = oc_hoje
            df_oc_f = df_oc

        # --- MÓDULO 1: DASHBOARD PRINCIPAL ---
        if menu == "🏠 Dashboard Principal":
            st.title("🌴 Dashboard Gestão de Equipe - Tropical")

            if not aniv_hoje_f.empty:
                st.balloons()
                for _, row in aniv_hoje_f.iterrows():
                    st.success(f"🎉 **HOJE É DIA DE FESTA!** Parabéns ao colaborador **{row['Funcionário']}** ({row['Setor']}) pelo seu aniversário hoje! 🎂🎈")

            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Quadro Total", len(df_f))
            col2.metric("👷 Ativos na Operação", len(df_ativos_f))
            col3.metric("🏖️ Em Férias", len(df_ferias_f))
            col4.metric("🏥 INSS / Afastados", len(df_inss_f))
            col5.metric("📋 Ocorrências Hoje", len(oc_hoje_f))

            st.markdown("---")

            st.info(f"🎂 **Aniversariantes do Mês ({hoje.strftime('%m/%Y')}):** {len(df_aniv_f)} colaborador(es) comemorando aniversário este mês.")

            if not df_inss_f.empty:
                st.warning(f"🏥 **Atenção:** Existem {len(df_inss_f)} colaborador(es) afastado(s) pelo INSS/Licença no setor **{setor_selecionado}**. Acesse o menu lateral '🏥 Janela - Afastados (INSS)' para ver a lista.")

            if not df_exp_f.empty:
                vencendo_7d = df_exp_f[
                    ((df_exp_f['Venc_45_dias'] >= hoje) & (df_exp_f['Venc_45_dias'] <= hoje + timedelta(days=7))) |
                    ((df_exp_f['Venc_90_dias'] >= hoje) & (df_exp_f['Venc_90_dias'] <= hoje + timedelta(days=7)))
                ]
                if not vencendo_7d.empty:
                    st.warning(f"⚠️ **Alerta RH:** Existem {len(vencendo_7d)} contrato(s) de experiência atingindo prazo nos próximos 7 dias!")

            g_col1, g_col2 = st.columns(2)
            with g_col1:
                st.subheader("📊 Distribuição por Setor")
                fig_setor = px.pie(df_ativos_f, names='Setor', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                st.plotly_chart(fig_setor, use_container_width=True)

            with g_col2:
                st.subheader("📈 Status do Quadro")
                fig_status = px.bar(df_f, x='Setor', color='Status', barmode='group', color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_status, use_container_width=True)

        # --- MÓDULO 2: JANELA DO INSS (RESTAURADA E GARANTIDA) ---
        elif menu == "🏥 Janela - Afastados (INSS)":
            st.title("🏥 Colaboradores Afastados (INSS / Licença)")
            st.caption(f"Visualizando colaboradores com status 'INSS' ou 'Afastado' no setor {setor_selecionado}.")

            if df_inss_f.empty:
                st.success("✅ Nenhum colaborador deste setor está em situação de afastamento pelo INSS no momento.")
            else:
                st.warning(f"📋 Encontrado(s) {len(df_inss_f)} colaborador(es) em situação de INSS / Afastamento:")
                
                # Exibe a tabela completa de afastados com todas as colunas
                st.dataframe(df_inss_f, use_container_width=True, hide_index=True)

                st.download_button(
                    label="📥 Exportar Lista do INSS em Excel (.xlsx)",
                    data=converter_para_excel(df_inss_f, "Afastados_INSS"),
                    file_name=f"afastados_inss_tropical_{hoje.strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # --- MÓDULO 3: JANELA DE FÉRIAS ---
        elif menu == "🏖️ Janela - Equipe em Férias":
            st.title("🏖️ Equipe em Gozo de Férias")
            st.caption(f"Visualizando colaboradores com status 'Férias' no setor {setor_selecionado}.")

            if df_ferias_f.empty:
                st.success("✅ Nenhum colaborador deste setor está em férias no momento.")
            else:
                st.warning(f"🏖️ Encontrado(s) {len(df_ferias_f)} colaborador(es) em férias:")
                st.dataframe(df_ferias_f, use_container_width=True, hide_index=True)
                st.download_button(
                    label="📥 Exportar Férias em Excel (.xlsx)",
                    data=converter_para_excel(df_ferias_f, "Ferias"),
                    file_name=f"equipe_em_ferias_{hoje.strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # --- MÓDULO 4: OCORRÊNCIAS ---
        elif menu == "📋 Ocorrências (Faltas/Atestados/Folgas)":
            st.title("📋 Controle de Ocorrências e Frequência")
            tab_reg, tab_hist = st.tabs(["➕ Registrar Nova Ocorrência", "📜 Histórico de Ocorrências"])

            with tab_reg:
                with st.form("form_nova_ocorrencia"):
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        func_selecionado = st.selectbox("Selecione o Colaborador:", df['Funcionário'].dropna().unique())
                        tipo_oc = st.selectbox("Tipo de Ocorrência:", [
                            "Atestado Médico", "Falta Justificada", "Falta Injustificada", 
                            "Folga Compensatória", "Licença Maternidade/Paternidade", "Advertência / Suspensão"
                        ])
                    with col_f2:
                        dt_inicio = st.date_input("Data de Início:", value=hoje)
                        dt_fim = st.date_input("Data de Término:", value=hoje)

                    obs_oc = st.text_area("Observação / Motivo:")
                    btn_salvar_oc = st.form_submit_button("💾 Salvar Ocorrência")

                    if btn_salvar_oc:
                        setor_func = df[df['Funcionário'] == func_selecionado]['Setor'].values[0] if 'Setor' in df.columns else "N/A"
                        qtd_dias = (dt_fim - dt_inicio).days + 1
                        nova_oc = {
                            "Data_Registro": hoje.strftime("%d/%m/%Y"),
                            "Funcionário": func_selecionado,
                            "Setor": setor_func,
                            "Tipo_Ocorrencia": tipo_oc,
                            "Data_Inicio": dt_inicio.strftime("%Y-%m-%d"),
                            "Data_Fim": dt_fim.strftime("%Y-%m-%d"),
                            "Dias": qtd_dias,
                            "Motivo_Observacao": obs_oc
                        }
                        df_oc_novo = pd.concat([df_oc, pd.DataFrame([nova_oc])], ignore_index=True)
                        df_oc_novo.to_excel(ARQUIVO_OCORRENCIAS, index=False)
                        st.success("Ocorrência salva com sucesso!")
                        st.cache_data.clear()
                        st.rerun()

            with tab_hist:
                st.dataframe(df_oc_f, use_container_width=True, hide_index=True)
                st.download_button(
                    label="📥 Exportar Relatório de Ocorrências em Excel",
                    data=converter_para_excel(df_oc_f, "Ocorrencias"),
                    file_name=f"ocorrencias_{hoje.strftime('%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # --- MÓDULO 5: ANIVERSARIANTES ---
        elif menu == "🎂 Aniversariantes do Mês":
            st.title("🎂 Aniversariantes do Mês Vigente")
            if df_aniv_f.empty:
                st.info("Nenhum aniversariante encontrado no mês atual para o setor selecionado.")
            else:
                cols_aniv = [c for c in ['Funcionário', 'Setor', 'Cargo', 'dt_nascimento'] if c in df_aniv_f.columns]
                df_aniv_exibir = df_aniv_f[cols_aniv].copy()
                if 'dt_nascimento' in df_aniv_exibir.columns:
                    df_aniv_exibir['Dia'] = df_aniv_exibir['dt_nascimento'].dt.day
                    df_aniv_exibir = df_aniv_exibir.sort_values(by='Dia')
                    df_aniv_exibir['Data de Nascimento'] = df_aniv_exibir['dt_nascimento'].dt.strftime('%d/%m')
                    df_aniv_exibir = df_aniv_exibir.drop(columns=['dt_nascimento', 'Dia'])

                st.dataframe(df_aniv_exibir, use_container_width=True, hide_index=True)
                st.download_button("📥 Exportar Aniversariantes (.xlsx)", converter_para_excel(df_aniv_exibir, "Aniversariantes"), file_name=f"aniversariantes_{hoje.strftime('%m_%Y')}.xlsx")

        # --- MÓDULO 6: EXPERIÊNCIA ---
        elif menu == "⏳ Contratos de Experiência (45/90d)":
            st.title("⏳ Controle de Contratos de Experiência")
            if df_exp_f.empty:
                st.success("✅ Nenhum colaborador em período de experiência no momento.")
            else:
                cols_exp = [c for c in ['Funcionário', 'Setor', 'Cargo', 'dt_adm', 'Venc_45_dias', 'Venc_90_dias'] if c in df_exp_f.columns]
                df_exp_exibir = df_exp_f[cols_exp].copy()
                if 'dt_adm' in df_exp_exibir.columns:
                    df_exp_exibir['Admissão'] = df_exp_exibir['dt_adm'].dt.strftime('%d/%m/%Y')
                    df_exp_exibir = df_exp_exibir.drop(columns=['dt_adm'])
                st.dataframe(df_exp_exibir, use_container_width=True, hide_index=True)
                st.download_button("📥 Exportar Experiência (.xlsx)", converter_para_excel(df_exp_exibir, "Experiencia"), file_name=f"experiencia_{hoje.strftime('%d_%m_%Y')}.xlsx")

        # --- MÓDULO 7: ESCALA INTELIGENTE DE FÉRIAS ---
        elif menu == "📅 Escala Inteligente de Férias":
            ferias.renderizar_modulo_ferias(df)

        # --- MÓDULO 8: CADASTRAR / EDITAR COLABORADOR ---
        elif menu == "👥 Cadastrar / Editar Colaborador":
            st.title("👥 Cadastrar ou Alterar Colaboradores")
            tab_cad, tab_edit = st.tabs(["➕ Novo Colaborador", "✏️ Editar Base"])

            with tab_cad:
                with st.form("form_novo_colab"):
                    f_nome = st.text_input("Nome Completo:")
                    c1, c2 = st.columns(2)
                    with c1:
                        f_setor = st.selectbox("Setor:", ["Separação", "Carregamento", "Recebimento", "Motorista", "Administrativo", "Limpeza", "Outros"])
                        f_cargo = st.text_input("Cargo:")
                    with c2:
                        f_adm = st.date_input("Data de Admissão:", value=hoje)
                        f_nasc = st.date_input("Data de Nascimento:", value=date(1990, 1, 1))
                    f_status = st.selectbox("Status Inicial:", ["Ativo", "Férias", "INSS", "Afastado"])
                    
                    if st.form_submit_button("💾 Salvar Colaborador"):
                        if f_nome.strip():
                            novo = {
                                "Funcionário": f_nome.strip(),
                                "Setor": f_setor,
                                "Cargo": f_cargo,
                                "Admissão": f_adm.strftime("%d/%m/%Y"),
                                "Data_Nascimento": f_nasc.strftime("%d/%m/%Y"),
                                "Status": f_status
                            }
                            df_novo = pd.concat([df, pd.DataFrame([novo])], ignore_index=True)
                            df_novo.to_excel(ARQUIVO_DADOS, index=False)
                            st.success("Colaborador cadastrado!")
                            st.cache_data.clear()
                            st.rerun()

            with tab_edit:
                df_editado = st.data_editor(df_f, use_container_width=True, key="editor_geral")
                if st.button("💾 Salvar Alterações"):
                    df_editado.to_excel(ARQUIVO_DADOS, index=False)
                    st.success("Base atualizada!")
                    st.cache_data.clear()
                    st.rerun()

            st.download_button("📥 Baixar Base Completa (.xlsx)", converter_para_excel(df_f, "Base_Equipe"), file_name=f"base_equipe_{hoje.strftime('%d_%m_%Y')}.xlsx")
