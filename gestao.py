import io
import os
from datetime import datetime, date, timedelta
import importlib
import ferias
import pandas as pd
import plotly.express as px
import streamlit as st

importlib.reload(ferias)

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

    @st.cache_data(ttl=60)
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

    df = carregar_dados()

    if df.empty:
        st.error("⚠️ Nenhuma base de dados encontrada no arquivo `equipe.xlsx`.")
    else:
        hoje = date.today()

        # Classificação do Status
        status_str = df['Status'].astype(str).str.lower()
        is_ferias = status_str.str.contains('férias|ferias')
        is_inss = status_str.str.contains('inss|afastado|licença|licenca')

        # Operação Ativa: nem em férias nem no INSS
        df_ativos = df[~is_ferias & ~is_inss].copy()
        df_em_ferias = df[is_ferias].copy()
        df_inss = df[is_inss].copy()

        # --- CÁLCULO DOS CONTRATOS DE EXPERIÊNCIA (45 e 90 Dias) ---
        if 'dt_adm' in df_ativos.columns:
            df_ativos['Venc_45_dias'] = df_ativos['dt_adm'].dt.date + timedelta(days=45)
            df_ativos['Venc_90_dias'] = df_ativos['dt_adm'].dt.date + timedelta(days=90)
            df_exp = df_ativos[df_ativos['dt_adm'].dt.date >= (hoje - timedelta(days=90))].copy()
        else:
            df_exp = pd.DataFrame()

        # --- ANIVERSARIANTES DO DIA E DO MÊS ---
        if 'dt_nascimento' in df.columns:
            aniv_hoje = df[(df['dt_nascimento'].dt.day == hoje.day) & (df['dt_nascimento'].dt.month == hoje.month)].copy()
            df_aniversariantes = df[df['dt_nascimento'].dt.month == hoje.month].copy()
        else:
            aniv_hoje = pd.DataFrame()
            df_aniversariantes = pd.DataFrame()

        # --- BARRA LATERAL (NAVEGAÇÃO) ---
        st.sidebar.title("🌴 Gestão Tropical")

        st.sidebar.metric("👷 Operação Ativa", f"{len(df_ativos)} colab.")
        st.sidebar.metric("🏖️ Em Férias Hoje", f"{len(df_em_ferias)} colab.")
        st.sidebar.metric("🏥 Afastados (INSS)", f"{len(df_inss)} colab.")
        if not df_exp.empty:
            st.sidebar.metric("⏳ Em Experiência", f"{len(df_exp)} colab.")

        setores = ["Todos"] + list(df['Setor'].dropna().unique())
        setor_selecionado = st.sidebar.selectbox("Filtrar Setor:", setores)

        menu = st.sidebar.radio(
            "Selecione o Módulo:",
            [
                "🏠 Dashboard Principal",
                "🎂 Aniversariantes do Mês",
                "⏳ Contratos de Experiência (45/90d)",
                "🏖️ Janela - Equipe em Férias",
                "🏥 Janela - Afastados (INSS)",
                "📅 Escala Inteligente de Férias",
                "👥 Cadastrar / Editar Colaborador"
            ]
        )

        # Filtro por Setor
        if setor_selecionado != "Todos":
            df_ativos_f = df_ativos[df_ativos['Setor'] == setor_selecionado]
            df_ferias_f = df_em_ferias[df_em_ferias['Setor'] == setor_selecionado]
            df_inss_f = df_inss[df_inss['Setor'] == setor_selecionado]
            df_exp_f = df_exp[df_exp['Setor'] == setor_selecionado] if not df_exp.empty else pd.DataFrame()
            df_aniv_f = df_aniversariantes[df_aniversariantes['Setor'] == setor_selecionado] if not df_aniversariantes.empty else pd.DataFrame()
            aniv_hoje_f = aniv_hoje[aniv_hoje['Setor'] == setor_selecionado] if not aniv_hoje.empty else pd.DataFrame()
        else:
            df_ativos_f = df_ativos
            df_ferias_f = df_em_ferias
            df_inss_f = df_inss
            df_exp_f = df_exp
            df_aniv_f = df_aniversariantes
            aniv_hoje_f = aniv_hoje

        # --- MÓDULO 1: DASHBOARD PRINCIPAL ---
        if menu == "🏠 Dashboard Principal":
            st.title("🌴 Dashboard Gestão de Equipe - Tropical")

            # BANNER DESTACADO DE ANIVERSARIANTES DO DIA
            if not aniv_hoje_f.empty:
                st.balloons()
                for _, row in aniv_hoje_f.iterrows():
                    st.success(f"🎉 **HOJE É DIA DE FESTA!** Parabéns ao colaborador **{row['Funcionário']}** ({row['Setor']}) pelo seu aniversário hoje! 🎂🎈")

            # CARDS DE VISÃO GERAL
            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Quadro Total", len(df))
            col2.metric("👷 Ativos na Operação", len(df_ativos_f))
            col3.metric("🏖️ Em Férias", len(df_ferias_f))
            col4.metric("🏥 INSS / Afastados", len(df_inss_f))
            col5.metric("⏳ Em Experiência", len(df_exp_f))

            st.markdown("---")

            # ALERTAS OPERACIONAIS
            if not df_exp_f.empty:
                vencendo_7d = df_exp_f[
                    ((df_exp_f['Venc_45_dias'] >= hoje) & (df_exp_f['Venc_45_dias'] <= hoje + timedelta(days=7))) |
                    ((df_exp_f['Venc_90_dias'] >= hoje) & (df_exp_f['Venc_90_dias'] <= hoje + timedelta(days=7)))
                ]
                if not vencendo_7d.empty:
                    st.warning(f"⚠️ **Alerta RH:** Existem {len(vencendo_7d)} contrato(s) de experiência atingindo prazo nos próximos 7 dias!")

            if not df_inss_f.empty:
                st.info(f"ℹ️ **Informação:** {len(df_inss_f)} colaborador(es) estão temporariamente afastados via INSS/Licença.")

            # GRÁFICOS ESTATÍSTICOS
            g_col1, g_col2 = st.columns(2)

            with g_col1:
                st.subheader("📊 Distribuição da Equipe por Setor")
                fig_setor = px.pie(df_ativos_f, names='Setor', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                st.plotly_chart(fig_setor, use_container_width=True)

            with g_col2:
                st.subheader("📈 Status do Quadro de Colaboradores")
                fig_status = px.bar(df, x='Setor', color='Status', barmode='group', color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_status, use_container_width=True)

        # --- MÓDULO 2: ANIVERSARIANTES DO MÊS ---
        elif menu == "🎂 Aniversariantes do Mês":
            st.title("🎂 Aniversariantes do Mês Vigente")
            st.caption(f"Colaboradores que comemoram aniversário no mês {hoje.strftime('%m/%Y')}.")

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

        # --- MÓDULO 3: CONTRATOS DE EXPERIÊNCIA ---
        elif menu == "⏳ Contratos de Experiência (45/90d)":
            st.title("⏳ Controle de Contratos de Experiência")
            st.caption("Acompanhamento das duas etapas de avaliação: 45 dias e 90 dias.")

            if df_exp_f.empty:
                st.success("✅ Nenhum colaborador em período de experiência (primeiros 90 dias) no momento.")
            else:
                cols_exp = ['Funcionário', 'Setor', 'Cargo', 'dt_adm', 'Venc_45_dias', 'Venc_90_dias']
                cols_exp_presentes = [c for c in cols_exp if c in df_exp_f.columns]

                df_exp_exibir = df_exp_f[cols_exp_presentes].copy()
                if 'dt_adm' in df_exp_exibir.columns:
                    df_exp_exibir['Admissão'] = df_exp_exibir['dt_adm'].dt.strftime('%d/%m/%Y')
                    df_exp_exibir = df_exp_exibir.drop(columns=['dt_adm'])

                st.dataframe(df_exp_exibir, use_container_width=True, hide_index=True)

        # --- MÓDULO 4: JANELA - EQUIPE EM FÉRIAS ---
        elif menu == "🏖️ Janela - Equipe em Férias":
            st.title("🏖️ Equipe em Gozo de Férias")
            st.caption("Acompanhamento de colaboradores temporariamente ausentes por férias.")

            if df_ferias_f.empty:
                st.success("✅ Nenhum colaborador deste setor está em férias no momento. Quadro 100% ativo!")
            else:
                st.warning(f"⚠️ Existem {len(df_ferias_f)} colaborador(es) afastado(s) em férias.")
                cols_ferias = [c for c in ['Funcionário', 'Setor', 'Cargo', 'Inicio_Ferias', 'Fim_Ferias', 'Status'] if c in df_ferias_f.columns]
                st.dataframe(df_ferias_f[cols_ferias if cols_ferias else df_ferias_f.columns], use_container_width=True, hide_index=True)

        # --- MÓDULO 5: JANELA - AFASTADOS (INSS) ---
        elif menu == "🏥 Janela - Afastados (INSS)":
            st.title("🏥 Colaboradores Afastados (INSS / Licença)")
            st.caption("Controle de afastamentos médicos e licenças temporárias.")

            if df_inss_f.empty:
                st.success("✅ Nenhum colaborador deste setor está afastado pelo INSS no momento.")
            else:
                st.info(f"📋 Registrados {len(df_inss_f)} colaborador(es) em situação de INSS / Afastamento.")
                cols_inss = [c for c in ['Funcionário', 'Setor', 'Cargo', 'Admissão', 'Status'] if c in df_inss_f.columns]
                st.dataframe(df_inss_f[cols_inss if cols_inss else df_inss_f.columns], use_container_width=True, hide_index=True)

        # --- MÓDULO 6: ESCALA INTELIGENTE DE FÉRIAS ---
        elif menu == "📅 Escala Inteligente de Férias":
            ferias.renderizar_modulo_ferias(df)

        # --- MÓDULO 7: CADASTRO / EDIÇÃO ---
        elif menu == "👥 Cadastrar / Editar Colaborador":
            st.title("👥 Cadastrar ou Alterar Status de Colaborador")
            st.info("Abaixo está a base geral de colaboradores da Tropical para consulta e edições.")
            st.dataframe(df, use_container_width=True)
