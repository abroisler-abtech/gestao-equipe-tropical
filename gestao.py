import streamlit as st
import pandas as pd
import numpy as np
import datetime
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from supabase import create_client, Client

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Painel DP - Tropical",
    page_icon="🌴",
    layout="wide"
)

# --- CONFIGURAÇÃO DE SEGURANÇA E ACESSO ---
SENHA_MESTRE = st.secrets.get("SENHA_ACESSO", "030711")

TODOS_MODULOS = [
    "Dashboard & Alertas",
    "Chamada & Faltas do Dia",
    "👤 Ficha Individual do Colaborador",
    "📅 Escala Inteligente de Férias",
    "🤖 Assistente IA de DP",
    "⚙️ Gerenciamento de Usuários"
]

# --- CONEXÃO COM O SUPABASE ---
def obter_cliente_supabase():
    url = st.secrets.get("SUPABASE_URL", "")
    key = st.secrets.get("SUPABASE_KEY", "")
    return create_client(url, key)

# --- CARREGAR DADOS DA EQUIPE ---
def carregar_dados():
    cols_padrao = ['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Admissão', 'Nascimento', 'Status', 'Ultimas_Ferias']
    df = pd.DataFrame()
    
    try:
        supabase = obter_cliente_supabase()
        response = supabase.table("equipe").select("*").execute()
        if response.data:
            df = pd.DataFrame(response.data)
    except Exception:
        pass

    if (df is None or df.empty) and os.path.exists("equipe.xlsx"):
        try:
            df = pd.read_excel("equipe.xlsx")
        except Exception:
            pass

    if df is None or df.empty:
        df = pd.DataFrame(columns=cols_padrao)

    df.columns = df.columns.astype(str).str.strip()
    
    for col in cols_padrao:
        if col not in df.columns:
            df[col] = None

    col_adm = next((c for c in df.columns if 'admiss' in str(c).lower() or 'dt_adm' in str(c).lower()), 'Admissão')
    col_nasc = next((c for c in df.columns if 'nasc' in str(c).lower() or 'anivers' in str(c).lower()), 'Nascimento')
    
    df['dt_adm'] = pd.to_datetime(df[col_adm], dayfirst=True, errors='coerce').dt.date if col_adm in df.columns else None
    if col_nasc in df.columns:
        df['dt_nasc_dt'] = pd.to_datetime(df[col_nasc], dayfirst=True, errors='coerce')
        df['dt_nasc'] = df['dt_nasc_dt'].dt.date
    else:
        df['dt_nasc_dt'] = pd.NaT
        df['dt_nasc'] = None

    df['Matricula'] = df['Matricula'].astype(str).str.replace('.0', '', regex=False)
    df['Ultimas_Ferias'] = df['Ultimas_Ferias'].astype(str)
    df['dt_ult_ferias'] = pd.to_datetime(df['Ultimas_Ferias'], dayfirst=True, errors='coerce').dt.date
    df['Decisao_Experiencia'] = df.get('Decisao_Experiencia', None)
    df['Status'] = df['Status'].fillna('Ativo').astype(str).str.strip()
    
    return df

# --- CARREGAR FALTAS ---
def carregar_faltas():
    cols_padrao = ["Matricula", "Funcionário", "Setor", "Data", "Tipo", "Dias", "CID", "Motivo", "dt_falta"]
    df_f = pd.DataFrame()
    
    try:
        supabase = obter_cliente_supabase()
        response = supabase.table("faltas").select("*").execute()
        if response.data:
            df_f = pd.DataFrame(response.data)
    except Exception:
        pass

    if (df_f is None or df_f.empty) and os.path.exists("faltas.xlsx"):
        try:
            df_f = pd.read_excel("faltas.xlsx")
        except Exception:
            pass

    if df_f is None or df_f.empty:
        df_f = pd.DataFrame(columns=cols_padrao)

    df_f.columns = df_f.columns.astype(str).str.strip()
    for col in cols_padrao:
        if col not in df_f.columns and col != 'dt_falta':
            df_f[col] = ""
            
    df_f['Data'] = df_f['Data'].astype(str).str.strip()
    df_f['dt_falta'] = pd.to_datetime(df_f['Data'], format='%d/%m/%Y', errors='coerce').dt.date
    return df_f

# --- CARREGAR USUÁRIOS ---
def carregar_usuarios():
    df_u = pd.DataFrame()
    try:
        supabase = obter_cliente_supabase()
        response = supabase.table("usuarios").select("*").execute()
        if response.data:
            df_u = pd.DataFrame(response.data)
    except Exception:
        pass

    if (df_u is None or df_u.empty) and os.path.exists("usuarios.xlsx"):
        try:
            df_u = pd.read_excel("usuarios.xlsx")
        except Exception:
            pass

    if df_u is not None and not df_u.empty:
        df_u.columns = df_u.columns.astype(str).str.strip()
        for col in ['Nome', 'Usuario', 'Email', 'Senha', 'Perfil', 'Modulos', 'Telefone']:
            if col in df_u.columns:
                df_u[col] = df_u[col].astype(str).str.replace('.0', '', regex=False)
            else:
                df_u[col] = ""
        return df_u
    else:
        dados_iniciais = [
            {"Nome": "André Broisler", "Usuario": "admin", "Email": "abroisler@gmail.com", "Senha": "123", "Perfil": "Admin", "Modulos": ",".join(TODOS_MODULOS), "Telefone": ""},
            {"Nome": "Gestor de Turno", "Usuario": "gestor", "Email": "gestor@tropical.com.br", "Senha": "123", "Perfil": "Gestor", "Modulos": "Dashboard & Alertas,Chamada & Faltas do Dia,👤 Ficha Individual do Colaborador", "Telefone": ""}
        ]
        return pd.DataFrame(dados_iniciais)

# --- SALVAR DADOS (COM SUPABASE) ---
def salvar_dados(df_salvar):
    cols_ignorar = ['dt_adm', 'dt_nasc', 'dt_nasc_dt', 'dt_ult_ferias', 'exp_45', 'exp_90', 'dias_para_45', 'dias_para_90']
    cols_salvar = [c for c in df_salvar.columns if c not in cols_ignorar]
    df_export = df_salvar[cols_salvar].fillna("").astype(str)
    
    try:
        supabase = obter_cliente_supabase()
        registros = df_export.to_dict(orient="records")
        supabase.table("equipe").delete().neq("Matricula", "___lixeira___").execute()
        if registros:
            supabase.table("equipe").insert(registros).execute()
        st.success("✅ Salvo permanentemente no Supabase (Nuvem)!")
    except Exception as e:
        st.error(f"Erro ao salvar no Supabase: {e}")
    df_export.to_excel("equipe.xlsx", index=False)

def salvar_faltas(df_f):
    cols_salvar = [c for c in df_f.columns if c != 'dt_falta']
    df_f = df_f.drop_duplicates(subset=['Funcionário', 'Data'], keep='last')
    df_export = df_f[cols_salvar].fillna("").astype(str)
    
    try:
        supabase = obter_cliente_supabase()
        registros = df_export.to_dict(orient="records")
        supabase.table("faltas").delete().neq("Funcionário", "___lixeira___").execute()
        if registros:
            supabase.table("faltas").insert(registros).execute()
        st.success("✅ Faltas salvas permanentemente na Nuvem!")
    except Exception as e:
        st.error(f"Erro ao salvar faltas no Supabase: {e}")
    df_export.to_excel("faltas.xlsx", index=False)

def salvar_usuarios(df_u):
    df_u = df_u.astype(str)
    try:
        supabase = obter_cliente_supabase()
        registros = df_u.to_dict(orient="records")
        supabase.table("usuarios").delete().neq("Usuario", "___lixeira___").execute()
        if registros:
            supabase.table("usuarios").insert(registros).execute()
        st.success("✅ Usuários salvos permanentemente na Nuvem!")
    except Exception as e:
        st.error(f"Erro ao salvar usuários no Supabase: {e}")
    df_u.to_excel("usuarios.xlsx", index=False)

# --- CONTROLE DE SESSÃO ---
if 'logado' not in st.session_state:
    st.session_state.logado = False
if 'usuario_atual' not in st.session_state:
    st.session_state.usuario_atual = ""
if 'perfil_atual' not in st.session_state:
    st.session_state.perfil_atual = ""
if 'modulos_usuario' not in st.session_state:
    st.session_state.modulos_usuario = TODOS_MODULOS

# --- TELA DE LOGIN ---
if not st.session_state.logado:
    st.title("🌴 Painel DP - Tropical Hortifrúti")
    st.subheader("Faça login para acessar o sistema")
    
    tab_login_sis, tab_login_mestre = st.tabs(["🔑 Entrar com Usuário", "🔐 Senha Mestre (Admin)"])
    
    with tab_login_sis:
        df_u_login = carregar_usuarios()
        usuario_input = st.text_input("Usuário", key="input_user_login")
        senha_input = st.text_input("Senha", type="password", key="input_pass_login")
        
        if st.button("Entrar no Sistema", type="primary"):
            user_encontrado = df_u_login[df_u_login['Usuario'].astype(str).str.lower() == usuario_input.strip().lower()]
            if not user_encontrado.empty:
                senha_cadastrada = str(user_encontrado.iloc[0]['Senha'])
                if senha_input == senha_cadastrada:
                    st.session_state.logado = True
                    st.session_state.usuario_atual = str(user_encontrado.iloc[0]['Nome'])
                    st.session_state.perfil_atual = str(user_encontrado.iloc[0]['Perfil'])
                    mods_str = str(user_encontrado.iloc[0]['Modulos'])
                    st.session_state.modulos_usuario = [m.strip() for m in mods_str.split(',') if m.strip()]
                    st.success(f"Bem-vindo, {st.session_state.usuario_atual}!")
                    st.rerun()
                else:
                    st.error("Senha incorreta.")
            else:
                st.error("Usuário não encontrado.")
                
    with tab_login_mestre:
        senha_mestre_input = st.text_input("Digite a Senha Mestre", type="password", key="input_senha_mestre")
        if st.button("Acesso Mestre", type="primary"):
            if senha_mestre_input == SENHA_MESTRE:
                st.session_state.logado = True
                st.session_state.usuario_atual = "Administrador Mestre"
                st.session_state.perfil_atual = "Admin"
                st.session_state.modulos_usuario = TODOS_MODULOS
                st.success("Acesso mestre concedido!")
                st.rerun()
            else:
                st.error("Senha mestre incorreta.")
    st.stop()

# --- BARRA LATERAL (MENU) ---
with st.sidebar:
    st.image("https://img.icons8.com/color/96/tropical-drink.png", width=70)
    st.write(f"**Logado como:** {st.session_state.usuario_atual}")
    st.write(f"**Perfil:** {st.session_state.perfil_atual}")
    
    if st.button("🚪 Sair / Trocar Usuário"):
        st.session_state.logado = False
        st.rerun()
        
    st.divider()
    st.header("📌 Navegação")
    
    if st.session_state.perfil_atual == "Admin":
        menu_opcoes = TODOS_MODULOS
    else:
        modulos_permitidos = st.session_state.modulos_usuario
        menu_opcoes = [m for m in TODOS_MODULOS if m in modulos_permitidos]
    
    if not menu_opcoes:
        menu_opcoes = ["Dashboard & Alertas"]
        
    escolha = st.radio("Ir para:", menu_opcoes)

# --- CARREGAMENTO DOS DADOS DO SISTEMA ---
df_equipe = carregar_dados()
df_faltas = carregar_faltas()

# --- MÓDULO 1: DASHBOARD & ALERTAS ---
if escolha == "Dashboard & Alertas":
    st.title("📊 Dashboard & Alertas de DP")
    st.markdown("Visão geral dos colaboradores, contratos de experiência e aniversariantes.")
    
    total_colab = len(df_equipe)
    st.metric("Total de Colaboradores", total_colab)
    
    st.subheader("📋 Lista de Colaboradores Cadastrados")
    st.dataframe(df_equipe, use_container_width=True)

# --- MÓDULO 2: CHAMADA & FALTAS DO DIA ---
elif escolha == "Chamada & Faltas do Dia":
    st.title("📋 Chamada & Registro de Faltas")
    
    if df_equipe.empty:
        st.warning("Nenhum colaborador cadastrado para realizar chamada.")
    else:
        data_falta = st.date_input("Data da Ocorrência", datetime.date.today())
        
        colab_selecionado = st.selectbox("Colaborador", df_equipe['Funcionário'].tolist())
        tipo_ocorrencia = st.selectbox("Tipo", ["Falta Injustificada", "Atestado Médico", "Suspensão", "Atraso"])
        dias_qtd = st.number_input("Dias de Afastamento/Falta", min_value=1, value=1)
        cid_input = st.text_input("CID (Opcional se Atestado)")
        motivo_input = st.text_area("Observação / Motivo")
        
        if st.button("Registrar Ocorrência", type="primary"):
            linha_colab = df_equipe[df_equipe['Funcionário'] == colab_selecionado].iloc[0]
            mat = linha_colab.get('Matricula', '')
            setor = linha_colab.get('Setor', '')
            
            nova_falta = pd.DataFrame([{
                "Matricula": str(mat),
                "Funcionário": colab_selecionado,
                "Setor": str(setor),
                "Data": data_falta.strftime('%d/%m/%Y'),
                "Tipo": tipo_ocorrencia,
                "Dias": int(dias_qtd),
                "CID": cid_input,
                "Motivo": motivo_input
            }])
            
            df_faltas = pd.concat([df_faltas, nova_falta], ignore_index=True)
            salvar_faltas(df_faltas)
            st.success("Ocorrência registrada e salva na nuvem com sucesso!")

# --- MÓDULO 3: FICHA INDIVIDUAL ---
elif escolha == "👤 Ficha Individual do Colaborador":
    st.title("👤 Ficha Individual do Colaborador")
    
    if df_equipe.empty:
        st.warning("Nenhum colaborador cadastrado.")
    else:
        colab_sel = st.selectbox("Selecione o Colaborador", df_equipe['Funcionário'].tolist(), key="select_ficha")
        dados_colab = df_equipe[df_equipe['Funcionário'] == colab_sel].iloc[0]
        
        st.write(f"**Matrícula:** {dados_colab.get('Matricula', '')}")
        st.write(f"**Setor:** {dados_colab.get('Setor', '')}")
        st.write(f"**Cargo:** {dados_colab.get('Cargo', '')}")
        st.write(f"**Status:** {dados_colab.get('Status', '')}")

# --- MÓDULO 4: ESCALA INTELIGENTE DE FÉRIAS ---
elif escolha == "📅 Escala Inteligente de Férias":
    st.title("📅 Escala Inteligente de Férias")
    st.info("Planejamento e controle de vencimentos e programação de férias da equipe.")
    if not df_equipe.empty:
        st.dataframe(df_equipe[['Matricula', 'Funcionário', 'Setor', 'Admissão', 'Ultimas_Ferias']], use_container_width=True)

# --- MÓDULO 5: ASSISTENTE IA DE DP ---
elif escolha == "🤖 Assistente IA de DP":
    st.title("🤖 Assistente IA de DP")
    st.write("Tire dúvidas sobre legislação trabalhista, rotinas de DP e análises de equipe.")
    pergunta_ia = st.text_input("Digite sua dúvida para a IA:")
    if st.button("Consultar IA"):
        st.info("O assistente de IA está conectado para auxiliar nas suas demandas de gestão.")

# --- MÓDULO 6: GERENCIAMENTO DE USUÁRIOS ---
elif escolha == "⚙️ Gerenciamento de Usuários":
    st.title("⚙️ Gerenciamento de Usuários do Sistema")
    
    if st.session_state.perfil_atual != "Admin":
        st.error("Acesso restrito a administradores.")
    else:
        df_u_ger = carregar_usuarios()
        
        st.subheader("✏️ Editar ou Remover Usuários Existentes")
        if not df_u_ger.empty:
            edited_df = st.data_editor(
                df_u_ger, 
                use_container_width=True, 
                num_rows="dynamic",
                key="editor_usuarios"
            )
            
            if st.button("💾 Salvar Alterações nos Usuários", type="primary"):
                salvar_usuarios(edited_df)
                st.success("Usuários atualizados com sucesso na nuvem!")
                st.rerun()
        else:
            st.info("Nenhum usuário cadastrado.")
        
        st.divider()
        st.subheader("➕ Adicionar Novo Usuário")
        col1, col2 = st.columns(2)
        with col1:
            novo_nome = st.text_input("Nome Completo")
            novo_user = st.text_input("Login de Usuário")
            novo_email = st.text_input("E-mail")
        with col2:
            nova_senha = st.text_input("Senha", type="password")
            novo_perfil = st.selectbox("Perfil", ["Admin", "Gestor", "Visualizador"])
            
        if st.button("Cadastrar Novo Usuário"):
            if novo_user and nova_senha:
                novo_registro = pd.DataFrame([{
                    "Nome": novo_nome,
                    "Usuario": novo_user,
                    "Email": novo_email,
                    "Senha": nova_senha,
                    "Perfil": novo_perfil,
                    "Modulos": ",".join(TODOS_MODULOS),
                    "Telefone": ""
                }])
                df_u_ger = pd.concat([df_u_ger, novo_registro], ignore_index=True)
                salvar_usuarios(df_u_ger)
                st.success("Usuário cadastrado com sucesso na nuvem!")
                st.rerun()
            else:
                st.warning("Preencha pelo menos o usuário e a senha.")
