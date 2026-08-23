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
            
            # Padronização de Colunas da Tropical
            df.columns = df.columns.str.strip()
            
            # Trata Data de Admissão
            col_adm = next((c for c in df.columns if 'admiss' in str(c).lower() or 'dt_adm' in str(c).lower()), None)
            if col_adm:
                df['dt_adm'] = pd.to_datetime(df[col_adm], dayfirst=True, errors='coerce')
            
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
        # --- FILTRAGEM AUTOMÁTICA DE ATIVOS X FÉRIAS ---
        hoje = date.today()
        
        # Identifica quem está em gozo de férias (por Status)
        is_ferias = df['Status'].astype(str).str.lower().str.contains('férias|ferias')
        
        df_ativos = df[~is_ferias].copy()
        df_em_ferias = df[is_ferias].copy()

        # --- BARRA LATERAL (NAVEGAÇÃO) ---
        st.sidebar.title("🌴 Gestão Tropical")
        
        # Indicadores Rápidos no Menu Lateral
        st.sidebar.metric("👷 Operação Ativa", f"{len(df_ativos)} colab.")
        st.sidebar.metric("🏖️ Em Férias Hoje", f"{len(df_em_ferias)} colab.")
        
        setores = ["Todos"] + list(df['Setor'].dropna().unique())
        setor_selecionado = st.sidebar.selectbox("Filtrar Setor:", setores)

        menu = st.sidebar.radio(
            "Selecione o Módulo:",
            [
                "📊 Painel Geral de Ativos",
                "🏖️ Janela - Equipe em Férias",
                "📅 Escala Inteligente de Férias",
                "👥 Cadastrar / Editar Colaborador"
            ]
        )

        # Filtra os quadros pelo setor selecionado
        if setor_selecionado != "Todos":
            df_ativos_f = df_ativos[df_ativos['Setor'] == setor_selecionado]
            df_ferias_f = df_em_ferias[df_em_ferias['Setor'] == setor_selecionado]
        else:
            df_ativos_f = df_ativos
            df_ferias_f = df_em_ferias

        # --- MÓDULO 1: PAINEL GERAL DE ATIVOS ---
        if menu == "📊 Painel Geral de Ativos":
            st.title("📊 Painel Geral de Quadro Ativo")
            st.caption("Apenas colaboradores operacionais em atividade no momento.")

            col1, col2, col3 = st.columns(3)
            col1.metric("Total em Operação", len(df_ativos_f))
            col2.metric("Setor de Separação", len(df_ativos_f[df_ativos_f['Setor'].str.lower().str.contains('separa', na=False)]))
            col3.metric("Demais Setores", len(df_ativos_f[~df_ativos_f['Setor'].str.lower().str.contains('separa', na=False)]))

            st.subheader("📋 Lista de Colaboradores Ativos")
            cols_exibir = [c for c in ['Funcionário', 'Setor', 'Cargo', 'Admissão', 'Status'] if c in df_ativos_f.columns]
            st.dataframe(df_ativos_f[cols_exibir], use_container_width=True, hide_index=True)

        # --- MÓDULO 2: JANELA - EQUIPE EM FÉRIAS ---
        elif menu == "🏖️ Janela - Equipe em Férias":
            st.title("🏖️ Equipe em Gozo de Férias")
            st.caption("Acompanhamento de colaboradores temporariamente ausentes da operação.")

            if df_ferias_f.empty:
                st.success("✅ Nenhum colaborador deste setor está em férias no momento. Quadro 100% ativo!")
            else:
                st.warning(f"⚠️ Existem {len(df_ferias_f)} colaborador(es) afastado(s) em férias.")
                
                cols_ferias = [c for c in ['Funcionário', 'Setor', 'Cargo', 'Inicio_Ferias', 'Fim_Ferias', 'Status'] if c in df_ferias_f.columns]
                st.dataframe(df_ferias_f[cols_ferias if cols_ferias else df_ferias_f.columns], use_container_width=True, hide_index=True)

        # --- MÓDULO 3: ESCALA INTELIGENTE DE FÉRIAS ---
        elif menu == "📅 Escala Inteligente de Férias":
            ferias.renderizar_modulo_ferias(df)

        # --- MÓDULO 4: CADASTRO / EDIÇÃO ---
        elif menu == "👥 Cadastrar / Editar Colaborador":
            st.title("👥 Cadastrar ou Alterar Status de Colaborador")
            st.info("Aqui você pode alterar as informações do quadro da Tropical.")
            
            st.dataframe(df, use_container_width=True)
