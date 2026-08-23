import io
import os
from datetime import datetime, date, timedelta
import importlib
import ferias
import pandas as pd
import plotly.express as px
import streamlit as st

importlib.reload(ferias)

# --- FUNÇÃO AUXILIAR PARA EXPORTAR EXCEL ---
def converter_para_excel(df_exp, nome_aba='Base_Tropical'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_exp.to_excel(writer, index=False, sheet_name=nome_aba)
    return output.getvalue()

# --- SISTEMA DE AUTENTICAÇÃO POR SENHA SEGURA ---
def verificar_senha():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if not st.session_state["autenticado"]:
        st.title("🔒 Acesso Restrito - Gestão de Equipe Tropical")
        st.info("Por razões de segurança, informe a senha de acesso para continuar.")

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

    @st.cache_data(ttl=10)
    def carregar_dados():
        if os.path.exists(ARQUIVO_DADOS):
            df = pd.read_excel(ARQUIVO_DADOS)

            # Padronização de Colunas
            df.columns = df.columns.str.strip()

            # Trata Data de Admissão
            col_adm = next((c for c in df.columns if 'admiss' in str(c).lower() or 'dt_adm' in str(c).lower()), None)
            if col_adm:
                df['dt_adm'] = pd.to_datetime(df[col_adm], dayfirst=True, errors='coerce')

            # Trata Data de Nascimento
            col_nasc = next((c for c in df.columns if 'nasc' in str(c).lower() or 'anivers' in str(c).lower()), None)
            if col_nasc:
                df['dt_nascimento'] = pd.to_datetime(df[col_nasc], dayfirst=True, errors='coerce')

            # Trata Status (Padrão: Ativo)
            if 'Status' not in df.columns:
                df['Status'] = 'Ativo'
            else:
                df['Status'] = df['Status'].fillna('Ativo')

            return df
        return pd.DataFrame()

    @st.cache_data(ttl=10)
    def carregar_ocorrencias():
        if os.path.exists(ARQUIVO_OCORRENCIAS):
            df_oc = pd.read_excel(ARQUIVO_OCORRENCIAS)
            df_oc.columns = df_oc.columns.str.strip()
            if 'Data_Inicio' in df_oc.columns:
                df_oc['Data_Inicio'] = pd.to_datetime(df_oc['Data_Inicio'], dayfirst=True, errors='coerce').dt.date
            if 'Data_Fim' in df_oc.columns:
                df_oc['Data_Fim'] = pd.to_datetime(df_oc['Data_Fim'], dayfirst=True, errors='coerce').dt.date
            return df_oc
        return pd.DataFrame(columns=["Data_Registro", "Funcionário", "Setor", "Tipo_Ocorrencia", "Data_Inicio", "Data_Fim", "Dias", "Motivo_Observacao"])

    df = carregar_dados()
    df_oc = carregar_ocorrencias()

    if df.empty:
        st.error("⚠️ Nenhuma base de dados encontrada no arquivo `equipe.xlsx`.")
    else:
        hoje = date.today()

        # Classificação de Status
        status_str = df['Status'].astype(str).str.lower()
        is_ferias = status_str.str.contains('férias|ferias')
        is_inss = status_str.str.contains('inss|afastado|licença|licenca')

        df_ativos = df[~is_ferias & ~is_inss].copy()
        df_em_ferias = df[is_ferias].copy()
        df_inss = df[is_inss].copy()

        # Identifica ocorrências ativas hoje (Ex: atestado/folga/falta vigente no dia)
        if not df_oc.empty and 'Data_Inicio' in df_oc.columns and 'Data_Fim' in df_oc.columns:
            oc_hoje = df_oc[(df_oc['Data_Inicio'] <= hoje) & (df_oc['Data_Fim'] >= hoje)].copy()
        else:
            oc_hoje = pd.DataFrame()

        # --- CONTRATOS DE EXPERIÊNCIA (45 e 90 Dias) ---
        if 'dt_adm' in df_ativos.columns:
            df_ativos['Venc_45_dias'] = df_ativos['dt_adm'].dt.date + timedelta(days=45)
            df_ativos['Venc_90_dias'] = df_ativos['dt_adm'].dt.date + timedelta(days=90)
            df_exp = df_ativos[df_ativos['dt_adm'].dt.date >= (hoje - timedelta(days=90))].copy()
        else:
            df_exp = pd.DataFrame()

        # --- ANIVERSARIANTES DO DIA E DO MÊS ---
        if 'dt_nascimento' in df.columns:
            df_valid_nasc = df.dropna(subset=['dt_nascimento']).copy()
            
            aniv_hoje = df_valid_nasc[
                (df_valid_nasc['dt_nascimento'].dt.day == hoje.day) & 
                (df_valid_nasc['dt_nascimento'].dt.month == hoje.month)
            ].copy()
            
            df_aniversariantes = df_valid_nasc[df_valid_nasc['dt_nascimento'].dt.month == hoje.month].copy()
        else:
            aniv_hoje = pd.DataFrame()
            df_aniversariantes = pd.DataFrame()

        # --- LISTA DE MÓDULOS ---
        opcoes_menu = [
            "🏠 Dashboard Principal",
            "📋 Ocorrências (Faltas/Atestados/Folgas)",
            "🎂 Aniversariantes do Mês",
            "⏳ Contratos de Experiência (45/90d)",
            "🏖️ Janela - Equipe em Férias",
            "🏥 Janela - Afastados (INSS)",
            "📅 Escala Inteligente de Férias",
            "👥 Cadastrar / Editar Colaborador"
        ]

        if "modulo_ativo" not in st.session_state:
            st.session_state["modulo_ativo"] = "🏠 Dashboard Principal"

        # BARRA LATERAL (NAVEGAÇÃO)
        st.sidebar.title("🌴 Gestão Tropical")
        st.sidebar.metric("👷 Operação Ativa", f"{len(df_ativos)} colab.")
        st.sidebar.metric("🏖️ Em Férias Hoje", f"{len(df_em_ferias)} colab.")
        st.sidebar.metric("🏥 Afastados (INSS)", f"{len(df_inss)} colab.")
        st.sidebar.metric("📋 Ocorrências Hoje", f"{len(oc_hoje)} regist.")

        setores = ["Todos"] + list(df['Setor'].dropna().unique())
        setor_selecionado = st.sidebar.selectbox("Filtrar Setor:", setores)

        idx_menu = opcoes_menu.index(st.session_state["modulo_ativo"]) if st.session_state["modulo_ativo"] in opcoes_menu else 0
        menu_escolhido = st.sidebar.radio("Selecione o Módulo:", opcoes_menu, index=idx_menu, key="nav_radio")
        st.session_state["modulo_ativo"] = menu_escolhido

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

        def botao_voltar():
            if st.button("⬅️ Voltar ao Dashboard Principal"):
                st.session_state["modulo_ativo"] = "🏠 Dashboard Principal"
                st.rerun()

        # --- MÓDULO 1: DASHBOARD PRINCIPAL ---
        if st.session_state["modulo_ativo"] == "🏠 Dashboard Principal":
            st.title("🌴 Dashboard Gestão de Equipe - Tropical")

            if not aniv_hoje_f.empty:
                st.balloons()
                for _, row in aniv_hoje_f.iterrows():
                    st.success(f"🎉 **HOJE É DIA DE FESTA!** Parabéns ao colaborador **{row['Funcionário']}** ({row['Setor']}) pelo seu aniversário hoje! 🎂🎈")

            st.markdown("##### 📌 Cartões de Indicadores (Clique no botão para detalhar):")

            col1, col2, col3, col4, col5 = st.columns(5)

            with col1:
                st.metric("Quadro Total", len(df_f))

            with col2:
                st.metric("👷 Ativos na Operação", len(df_ativos_f))
                if st.button("Ver Ativos 🔍", key="btn_atv"):
                    st.session_state["modulo_ativo"] = "👥 Cadastrar / Editar Colaborador"
                    st.rerun()

            with col3:
                st.metric("🏖️ Em Férias", len(df_ferias_f))
                if st.button("Ver Férias 🔍", key="btn_fer"):
                    st.session_state["modulo_ativo"] = "🏖️ Janela - Equipe em Férias"
                    st.rerun()

            with col4:
                st.metric("🏥 INSS / Afastados", len(df_inss_f))
                if st.button("Ver INSS 🔍", key="btn_ins"):
                    st.session_state["modulo_ativo"] = "🏥 Janela - Afastados (INSS)"
                    st.rerun()

            with col5:
                st.metric("📋 Ocorrências Hoje", len(oc_hoje_f))
                if st.button("Ver Faltas/Folgas 🔍", key="btn_oc"):
                    st.session_state["modulo_ativo"] = "📋 Ocorrências (Faltas/Atestados/Folgas)"
                    st.rerun()

            st.markdown("---")

            # BARRA DE EXPORTAÇÃO E ANIVERSARIANTES
            c_aniv1, c_aniv2, c_aniv3 = st.columns([3, 1, 1])
            with c_aniv1:
                st.info(f"🎂 **Aniversariantes do Mês ({hoje.strftime('%m/%Y')}):** {len(df_aniv_f)} colaborador(es) comemorando aniversário este mês.")
            with c_aniv2:
                if st.button("Ver Aniversariantes 🎂", key="btn_aniv"):
                    st.session_state["modulo_ativo"] = "🎂 Aniversariantes do Mês"
                    st.rerun()
            with c_aniv3:
                st.download_button(
                    label="📥 Baixar Base (.xlsx)",
                    data=converter_para_excel(df_f),
                    file_name=f"base_equipe_tropical_{setor_selecionado.lower()}_{hoje.strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # ALERTAS DE OCORRÊNCIAS HOJE
            if not oc_hoje_f.empty:
                st.warning(f"⚠️ **Atenção:** Existem {len(oc_hoje_f)} ocorrência(s) registrada(s) para o dia de hoje (atestados, faltas ou folgas).")

            # ALERTAS DE EXPERIÊNCIA
            if not df_exp_f.empty:
                vencendo_7d = df_exp_f[
                    ((df_exp_f['Venc_45_dias'] >= hoje) & (df_exp_f['Venc_45_dias'] <= hoje + timedelta(days=7))) |
                    ((df_exp_f['Venc_90_dias'] >= hoje) & (df_exp_f['Venc_90_dias'] <= hoje + timedelta(days=7)))
                ]
                if not vencendo_7d.empty:
                    st.warning(f"⚠️ **Alerta RH:** Existem {len(vencendo_7d)} contrato(s) de experiência atingindo prazo nos próximos 7 dias!")

            # GRÁFICOS
            g_col1, g_col2 = st.columns(2)

            with g_col1:
                st.subheader("📊 Distribuição por Setor")
                fig_setor = px.pie(df_ativos_f, names='Setor', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                st.plotly_chart(fig_setor, use_container_width=True)

            with g_col2:
                st.subheader("📈 Status do Quadro")
                fig_status = px.bar(df_f, x='Setor', color='Status', barmode='group', color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_status, use_container_width=True)

        # --- MÓDULO 2: GESTÃO DE OCORRÊNCIAS (FALTAS, ATESTADOS E FOLGAS) ---
        elif st.session_state["modulo_ativo"] == "📋 Ocorrências (Faltas/Atestados/Folgas)":
            botao_voltar()
            st.title("📋 Controle de Ocorrências e Frequência")
            st.caption("Registro e acompanhamento de Faltas, Atestados Médicos, Folgas Compensatórias e Licenças.")

            tab_reg, tab_hist = st.tabs(["➕ Registrar Nova Ocorrência", "📜 Histórico de Ocorrências"])

            with tab_reg:
                st.subheader("Lançamento de Ocorrência")
                with st.form("form_nova_ocorrencia"):
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        func_selecionado = st.selectbox("Selecione o Colaborador:", df['Funcionário'].dropna().unique())
                        tipo_oc = st.selectbox("Tipo de Ocorrência:", [
                            "Atestado Médico", 
                            "Falta Justificada", 
                            "Falta Injustificada", 
                            "Folga Compensatória", 
                            "Licença Maternidade/Paternidade",
                            "Advertência / Suspensão"
                        ])
                    with col_f2:
                        dt_inicio = st.date_input("Data de Início:", value=hoje)
                        dt_fim = st.date_input("Data de Término:", value=hoje)

                    obs_oc = st.text_area("Observação / Motivo (Ex: CID do atestado, folga referente a domingo trabalhado):")

                    btn_salvar_oc = st.form_submit_button("💾 Salvar Ocorrência")

                    if btn_salvar_oc:
                        if dt_fim < dt_inicio:
                            st.error("A Data de Término não pode ser anterior à Data de Início.")
                        else:
                            # Busca o setor do funcionário
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
                            st.success(f"Ocorrência de **{tipo_oc}** lançada para **{func_selecionado}** ({qtd_dias} dia(s))!")
                            st.cache_data.clear()
                            st.rerun()

            with tab_hist:
                st.subheader(f"Histórico Registrado ({len(df_oc_f)} ocorrência(s))")

                if df_oc_f.empty:
                    st.info("Nenhuma ocorrência registrada para o setor selecionado.")
                else:
                    st.dataframe(df_oc_f, use_container_width=True, hide_index=True)

                    st.download_button(
                        label="📥 Exportar Relatório de Ocorrências em Excel (.xlsx)",
                        data=converter_para_excel(df_oc_f, "Ocorrencias"),
                        file_name=f"relatorio_ocorrencias_tropical_{hoje.strftime('%m_%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        # --- MÓDULO 3: ANIVERSARIANTES DO MÊS ---
        elif st.session_state["modulo_ativo"] == "🎂 Aniversariantes do Mês":
            botao_voltar()
            st.title("🎂 Aniversariantes do Mês Vigente")
            st.caption(f"Colaboradores do setor {setor_selecionado} com aniversário no mês {hoje.strftime('%m/%Y')}.")

            if df_aniv_f.empty:
                st.info("Nenhum aniversariante encontrado para o mês atual no setor selecionado.")
            else:
                cols_aniv = [c for c in ['Funcionário', 'Setor', 'Cargo', 'dt_nascimento'] if c in df_aniv_f.columns]
                df_aniv_exibir = df_aniv_f[cols_aniv].copy()
                if 'dt_nascimento' in df_aniv_exibir.columns:
                    df_aniv_exibir['Dia'] = df_aniv_exibir['dt_nascimento'].dt.day
                    df_aniv_exibir = df_aniv_exibir.sort_values(by='Dia')
                    df_aniv_exibir['Data de Nascimento'] = df_aniv_exibir['dt_nascimento'].dt.strftime('%d/%m')
                    df_aniv_exibir = df_aniv_exibir.drop(columns=['dt_nascimento', 'Dia'])

                st.dataframe(df_aniv_exibir, use_container_width=True, hide_index=True)

                st.download_button(
                    label="📥 Exportar Lista de Aniversariantes em Excel",
                    data=converter_para_excel(df_aniv_exibir, "Aniversariantes"),
                    file_name=f"aniversariantes_{hoje.strftime('%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # --- MÓDULO 4: CONTRATOS DE EXPERIÊNCIA ---
        elif st.session_state["modulo_ativo"] == "⏳ Contratos de Experiência (45/90d)":
            botao_voltar()
            st.title("⏳ Controle de Contratos de Experiência")
            st.caption("Acompanhamento de prazos de 45 e 90 dias de contratação.")

            if df_exp_f.empty:
                st.success("✅ Nenhum colaborador em período de experiência (primeiros 90 dias) no setor selecionado.")
            else:
                cols_exp = ['Funcionário', 'Setor', 'Cargo', 'dt_adm', 'Venc_45_dias', 'Venc_90_dias']
                cols_exp_presentes = [c for c in cols_exp if c in df_exp_f.columns]

                df_exp_exibir = df_exp_f[cols_exp_presentes].copy()
                if 'dt_adm' in df_exp_exibir.columns:
                    df_exp_exibir['Admissão'] = df_exp_exibir['dt_adm'].dt.strftime('%d/%m/%Y')
                    df_exp_exibir = df_exp_exibir.drop(columns=['dt_adm'])

                st.dataframe(df_exp_exibir, use_container_width=True, hide_index=True)

                st.download_button(
                    label="📥 Exportar Contratos de Experiência em Excel",
                    data=converter_para_excel(df_exp_exibir, "Experiencia"),
                    file_name=f"contratos_experiencia_{hoje.strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # --- MÓDULO 5: JANELA - EQUIPE EM FÉRIAS ---
        elif st.session_state["modulo_ativo"] == "🏖️ Janela - Equipe em Férias":
            botao_voltar()
            st.title("🏖️ Equipe em Gozo de Férias")
            st.caption("Acompanhamento de colaboradores temporariamente ausentes da operação.")

            if df_ferias_f.empty:
                st.success("✅ Nenhum colaborador deste setor está em férias no momento.")
            else:
                st.warning(f"⚠️ Existem {len(df_ferias_f)} colaborador(es) em férias.")
                cols_ferias = [c for c in ['Funcionário', 'Setor', 'Cargo', 'Inicio_Ferias', 'Fim_Ferias', 'Status'] if c in df_ferias_f.columns]
                st.dataframe(df_ferias_f[cols_ferias if cols_ferias else df_ferias_f.columns], use_container_width=True, hide_index=True)

                st.download_button(
                    label="📥 Exportar Colaboradores em Férias em Excel",
                    data=converter_para_excel(df_ferias_f, "Ferias"),
                    file_name=f"colaboradores_em_ferias_{hoje.strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # --- MÓDULO 6: JANELA - AFASTADOS (INSS) ---
        elif st.session_state["modulo_ativo"] == "🏥 Janela - Afastados (INSS)":
            botao_voltar()
            st.title("🏥 Colaboradores Afastados (INSS / Licença)")
            st.caption("Controle de afastamentos médicos e licenças temporárias.")

            if df_inss_f.empty:
                st.success("✅ Nenhum colaborador deste setor está afastado pelo INSS no momento.")
            else:
                st.info(f"📋 Registrados {len(df_inss_f)} colaborador(es) em situação de INSS / Afastamento.")
                cols_inss = [c for c in ['Funcionário', 'Setor', 'Cargo', 'Admissão', 'Status'] if c in df_inss_f.columns]
                st.dataframe(df_inss_f[cols_inss if cols_inss else df_inss_f.columns], use_container_width=True, hide_index=True)

                st.download_button(
                    label="📥 Exportar Lista do INSS em Excel",
                    data=converter_para_excel(df_inss_f, "INSS"),
                    file_name=f"afastados_inss_{hoje.strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # --- MÓDULO 7: ESCALA INTELIGENTE DE FÉRIAS ---
        elif st.session_state["modulo_ativo"] == "📅 Escala Inteligente de Férias":
            ferias.renderizar_modulo_ferias(df)

        # --- MÓDULO 8: CADASTRAR / EDITAR COLABORADOR ---
        elif st.session_state["modulo_ativo"] == "👥 Cadastrar / Editar Colaborador":
            botao_voltar()
            st.title("👥 Cadastrar ou Alterar Status de Colaborador")
            st.caption("Gerencie o quadro oficial da Tropical. Todas as alterações serão salvas no arquivo `equipe.xlsx`.")

            tab_cad, tab_edit = st.tabs(["➕ Novo Colaborador", "✏️ Editar Status / Dados Existententes"])

            with tab_cad:
                st.subheader("Formulário de Novo Cadastro")
                with st.form("form_novo_colab"):
                    f_nome = st.text_input("Nome Completo do Funcionário:")
                    c_col1, c_col2 = st.columns(2)
                    with c_col1:
                        f_setor = st.selectbox("Setor:", ["Separação", "Carregamento", "Recebimento", "Motorista", "Administrativo", "Limpeza", "Outros"])
                        f_cargo = st.text_input("Cargo:")
                    with c_col2:
                        f_adm = st.date_input("Data de Admissão:", value=hoje)
                        f_nasc = st.date_input("Data de Nascimento:", value=date(1990, 1, 1))
                    
                    f_status = st.selectbox("Status Inicial:", ["Ativo", "Férias", "INSS", "Afastado"])
                    
                    btn_salvar_novo = st.form_submit_button("💾 Salvar Colaborador")

                    if btn_salvar_novo:
                        if not f_nome.strip():
                            st.error("Por favor, preencha o Nome Completo.")
                        else:
                            novo_registro = {
                                "Funcionário": f_nome.strip(),
                                "Setor": f_setor,
                                "Cargo": f_cargo,
                                "Admissão": f_adm.strftime("%d/%m/%Y"),
                                "Data_Nascimento": f_nasc.strftime("%d/%m/%Y"),
                                "Status": f_status
                            }
                            df_atualizado = pd.concat([df, pd.DataFrame([novo_registro])], ignore_index=True)
                            df_atualizado.to_excel(ARQUIVO_DADOS, index=False)
                            st.success(f"Colaborador **{f_nome}** cadastrado com sucesso!")
                            st.cache_data.clear()
                            st.rerun()

            with tab_edit:
                st.subheader("Base Geral da Tropical (Edição Rápida)")
                st.info("Para alterar o Status (Ex: mudar de 'Ativo' para 'INSS' ou 'Férias'), edite a tabela abaixo e clique em **Salvar Alterações**.")

                df_editado_geral = st.data_editor(
                    df_f,
                    use_container_width=True,
                    num_rows="dynamic",
                    key="editor_base_geral"
                )

                if st.button("💾 Salvar Alterações na Base"):
                    df_editado_geral.to_excel(ARQUIVO_DADOS, index=False)
                    st.success("Base de dados `equipe.xlsx` atualizada com sucesso!")
                    st.cache_data.clear()
                    st.rerun()

            st.markdown("---")
            st.download_button(
                label="📥 Baixar Planilha Completa da Equipe (.xlsx)",
                data=converter_para_excel(df_f, "Equipe_Completa"),
                file_name=f"base_equipe_completa_{hoje.strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
