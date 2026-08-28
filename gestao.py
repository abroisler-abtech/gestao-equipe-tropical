"""Painel de Gestão & DP — Versão Definitiva Consolidada.

Execute com: streamlit run gestao_corrigido.py
"""

from __future__ import annotations

import io
import locale
import os
import re
import shutil
import unicodedata
import uuid
from datetime import date, datetime, timedelta
from html import escape
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import streamlit as st

try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except Exception:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except Exception:
        pass

try:
    from supabase import Client, create_client
except ImportError:
    Client = Any
    create_client = None


st.set_page_config(page_title="Painel de Gestão & DP", page_icon="🍊", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "dados"
BACKUP_DIR = DATA_DIR / "backups"
DATA_DIR.mkdir(exist_ok=True)
BACKUP_DIR.mkdir(exist_ok=True)

ARQUIVOS = {
    "colaboradores": DATA_DIR / "equipe.xlsx",
    "usuarios": DATA_DIR / "usuarios.xlsx",
    "faltas": DATA_DIR / "faltas.xlsx",
    "epis": DATA_DIR / "epis.xlsx",
    "historico": DATA_DIR / "historico_colaboradores.xlsx",
}

COLUNAS: dict[str, list[str]] = {
    "colaboradores": [
        "matricula", "funcionario", "setor", "cargo", "admissao", "nascimento",
        "status", "ultimas_ferias", "data_retorno_ferias", "decisao_experiencia",
        "data_desligamento", "aprovacao_rh", "fracionamento", "escala_confirmada", "data_pre_agendada",
    ],
    "usuarios": ["usuario", "senha", "perfil"],
    "faltas": [
        "registro_id", "matricula", "funcionario", "setor", "data", "tipo", "dias",
        "cid", "motivo", "origem",
    ],
    "epis": ["entrega_id", "matricula", "funcionario", "setor", "data", "epi", "detalhe_tamanho", "responsavel", "tipo_registro"],
    "historico": ["evento_id", "matricula", "funcionario", "data", "tipo_evento", "descricao", "autor"],
}

CHAVES_PRIMARIAS = {
    "colaboradores": "matricula",
    "usuarios": "usuario",
    "faltas": "registro_id",
    "epis": "entrega_id",
    "historico": "evento_id",
}

TIPOS_OCORRENCIA = ("Falta Injustificada", "Atestado Médico", "Folga Concedida", "Ausência / A Confirmar")
STATUS_COLABORADOR = ("Ativo", "Férias", "Afastado", "Desligado")
TODOS_MODULOS = (
    "Dashboard & Alertas",
    "Assistente IA (DP & Gestão)",
    "Chamada & Faltas do Dia",
    "Solicitação & Entrega de EPI",
    "Ficha Individual do Colaborador",
    "Controle de Experiência (45/90 dias)",
    "Escala Inteligente de Férias & Folga",
    "Colaboradores em Férias",
    "Indicadores de Frequência & Absenteísmo",
    "Aniversariantes do Mês",
    "Cadastrar / Editar Colaborador",
    "Criar / Gerenciar Usuários",
    "Importar Nova Base",
)

MODULOS_ADMIN = {"Criar / Gerenciar Usuários", "Importar Nova Base"}

ALIAS_COLUNAS = {
    "colaboradores": {
        "matricula": "matricula", "matrícula": "matricula", "funcionario": "funcionario",
        "funcionário": "funcionario", "nome": "funcionario", "setor": "setor", "cargo": "cargo",
        "admissao": "admissao", "admissão": "admissao", "dt_adm": "admissao",
        "nascimento": "nascimento", "data_nascimento": "nascimento", "ultimas_ferias": "ultimas_ferias",
        "últimas_férias": "ultimas_ferias", "data_retorno_ferias": "data_retorno_ferias",
        "retorno_ferias": "data_retorno_ferias", "status": "status",
        "decisao_experiencia": "decisao_experiencia", "decisão_experiência": "decisao_experiencia",
        "data_desligamento": "data_desligamento", "aprovacao_rh": "aprovacao_rh",
        "fracionamento": "fracionamento", "escala_confirmada": "escala_confirmada", "data_pre_agendada": "data_pre_agendada",
    },
    "usuarios": {"usuario": "usuario", "senha": "senha", "perfil": "perfil"},
    "faltas": {
        "registro_id": "registro_id", "matricula": "matricula", "matrícula": "matricula",
        "funcionario": "funcionario", "funcionário": "funcionario", "setor": "setor", "data": "data",
        "tipo": "tipo", "dias": "dias", "cid": "cid", "motivo": "motivo", "origem": "origem",
    },
    "epis": {
        "entrega_id": "entrega_id", "matricula": "matricula", "matrícula": "matricula",
        "funcionario": "funcionario", "funcionário": "funcionario", "setor": "setor", "data": "data",
        "epi": "epi", "detalhe_tamanho": "detalhe_tamanho", "responsavel": "responsavel", "responsável": "responsavel",
        "tipo_registro": "tipo_registro",
    },
    "historico": {
        "evento_id": "evento_id", "matricula": "matricula", "matrícula": "matricula",
        "funcionario": "funcionario", "funcionário": "funcionario", "data": "data",
        "tipo_evento": "tipo_evento", "descricao": "descricao", "descrição": "descricao", "autor": "autor",
    },
}


def segredo(nome: str, padrao: str = "") -> str:
    try:
        return str(st.secrets.get(nome, os.getenv(nome, padrao)))
    except Exception:
        return os.getenv(nome, padrao)


def normalizar_cabecalho(nome: object) -> str:
    texto = unicodedata.normalize("NFKD", str(nome)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "_", texto.lower()).strip("_")


def limpar_texto(valor: object) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    return "" if texto.lower() in {"nan", "none", "nat"} else texto


def limpar_matricula(valor: object) -> str:
    texto = limpar_texto(valor)
    return re.sub(r"\.0$", "", texto)


def para_data(valor: object) -> Optional[date]:
    if valor is None or pd.isna(valor) or limpar_texto(valor) == "":
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    convertido = pd.to_datetime(str(valor), dayfirst=True, errors="coerce")
    if pd.isna(convertido):
        convertido = pd.to_datetime(str(valor), dayfirst=False, errors="coerce")
    return None if pd.isna(convertido) else convertido.date()


def data_iso(valor: object) -> Optional[str]:
    convertido = para_data(valor)
    return convertido.isoformat() if convertido else None


def formatar_data(valor: object) -> str:
    convertido = para_data(valor)
    return convertido.strftime("%d/%m/%Y") if convertido else "—"


@st.cache_resource(show_spinner=False)
def obter_supabase() -> Optional[Client]:
    url = segredo("SUPABASE_URL")
    chave = segredo("SUPABASE_SERVICE_ROLE_KEY") or segredo("SUPABASE_KEY")
    if not url or not chave or create_client is None:
        return None
    try:
        return create_client(url, chave)
    except Exception:
        return None


def renomear_colunas(df: pd.DataFrame, entidade: str) -> pd.DataFrame:
    resultado = df.copy()
    aliases = {normalizar_cabecalho(k): v for k, v in ALIAS_COLUNAS[entidade].items()}
    novos_nomes = {coluna: aliases.get(normalizar_cabecalho(coluna), normalizar_cabecalho(coluna)) for coluna in resultado.columns}
    resultado = resultado.rename(columns=novos_nomes)
    return resultado.loc[:, ~resultado.columns.duplicated()].copy()


def normalizar_entidade(df: pd.DataFrame, entidade: str) -> tuple[pd.DataFrame, bool]:
    df = renomear_colunas(df, entidade)
    for coluna in COLUNAS[entidade]:
        if coluna not in df.columns:
            if coluna in {"dias"}:
                df[coluna] = 0
            elif coluna in {"status"}:
                df[coluna] = "Ativo"
            elif coluna in {"perfil"}:
                df[coluna] = "Admin"
            elif coluna in {"tipo_registro"}:
                df[coluna] = "Entrega"
            elif coluna in {"escala_confirmada"}:
                df[coluna] = False
            else:
                df[coluna] = ""

    if entidade == "colaboradores":
        for coluna in ("matricula", "funcionario", "setor", "cargo", "status", "decisao_experiencia", "aprovacao_rh", "fracionamento"):
            if coluna in df.columns:
                df[coluna] = df[coluna].map(limpar_texto)
        df["matricula"] = df["matricula"].map(limpar_matricula)
        df["status"] = df["status"].replace({"Ferias": "Férias", "ferias": "Férias"}).replace("", "Ativo")
        for coluna in ("admissao", "nascimento", "ultimas_ferias", "data_retorno_ferias", "data_desligamento", "data_pre_agendada"):
            if coluna in df.columns:
                df[coluna] = df[coluna].map(data_iso)
    elif entidade == "usuarios":
        for coluna in ("usuario", "senha", "perfil"):
            df[coluna] = df[coluna].map(limpar_texto)
        df["usuario"] = df["usuario"].str.lower()
        df["perfil"] = df["perfil"].replace("", "Admin")
    elif entidade in {"faltas", "epis", "historico"}:
        id_coluna = CHAVES_PRIMARIAS[entidade]
        df[id_coluna] = df[id_coluna].map(limpar_texto)
        em_branco = df[id_coluna] == ""
        df.loc[em_branco, id_coluna] = [str(uuid.uuid4()) for _ in range(int(em_branco.sum()))]
        for coluna in df.columns:
            if coluna not in {"dias"}:
                df[coluna] = df[coluna].map(limpar_texto)
        df["data"] = df["data"].map(data_iso)
        if entidade == "faltas":
            df["dias"] = pd.to_numeric(df["dias"], errors="coerce").fillna(0).astype(int)
            df["origem"] = df["origem"].replace("", "Avulso")
        elif entidade == "epis":
            df["tipo_registro"] = df["tipo_registro"].replace("", "Entrega")

    return df[COLUNAS[entidade]].copy(), False


def ler_excel(entidade: str) -> pd.DataFrame:
    caminho = ARQUIVOS[entidade]
    if not caminho.exists() or caminho.stat().st_size < 100:
        return pd.DataFrame(columns=COLUNAS[entidade])
    try:
        return pd.read_excel(caminho, dtype=object)
    except Exception:
        return pd.DataFrame(columns=COLUNAS[entidade])


def copiar_backup(caminho: Path) -> None:
    if caminho.exists() and caminho.stat().st_size >= 100:
        destino = BACKUP_DIR / f"{caminho.stem}_{datetime.now():%Y%m%d_%H%M%S}{caminho.suffix}"
        shutil.copy2(caminho, destino)


def salvar_excel_atomico(df: pd.DataFrame, entidade: str) -> None:
    caminho = ARQUIVOS[entidade]
    temporario = caminho.with_name(f".{caminho.stem}.tmp.xlsx")
    copiar_backup(caminho)
    df.to_excel(temporario, index=False, engine="openpyxl")
    os.replace(temporario, caminho)


def carregar_entidade(entidade: str) -> tuple[pd.DataFrame, str]:
    if entidade in {"usuarios", "colaboradores"}:
        df, _ = normalizar_entidade(ler_excel(entidade), entidade)
        return df, "Local"

    cliente = obter_supabase()
    if cliente is not None:
        try:
            resposta = cliente.table(entidade).select("*").execute()
            dados = getattr(resposta, "data", None)
            if dados is not None and len(dados) > 0:
                df, _ = normalizar_entidade(pd.DataFrame(dados), entidade)
                return df, "Supabase"
        except Exception:
            pass

    df, _ = normalizar_entidade(ler_excel(entidade), entidade)
    return df, "Excel local"


def salvar_entidade(entidade: str, df: pd.DataFrame, mostrar_feedback: bool = True) -> bool:
    df_normalizado, _ = normalizar_entidade(df, entidade)
    if entidade not in {"usuarios", "colaboradores"}:
        cliente = obter_supabase()
        if cliente is not None:
            try:
                registros = df_normalizado.to_dict(orient="records")
                if registros:
                    cliente.table(entidade).upsert(registros, on_conflict=CHAVES_PRIMARIAS[entidade]).execute()
            except Exception:
                pass

    try:
        salvar_excel_atomico(df_normalizado, entidade)
    except Exception:
        pass

    if mostrar_feedback:
        st.success("Dados salvos com sucesso.")
    return True


def registrar_historico(matricula: str, funcionario: str, tipo: str, descricao: str, autor: str) -> None:
    historico, _ = carregar_entidade("historico")
    novo = {
        "evento_id": str(uuid.uuid4()), "matricula": limpar_matricula(matricula),
        "funcionario": limpar_texto(funcionario), "data": datetime.now().isoformat(timespec="minutes"),
        "tipo_evento": limpar_texto(tipo), "descricao": limpar_texto(descricao), "autor": limpar_texto(autor),
    }
    salvar_entidade("historico", pd.concat([historico, pd.DataFrame([novo])], ignore_index=True), False)


def iniciar_estado_sessao() -> None:
    padroes = {
        "autenticado": False, "usuario": "", "perfil": "",
        "fonte_colaboradores": "", "tentativas_login": 0,
    }
    for chave, valor in padroes.items():
        st.session_state.setdefault(chave, valor)


def provisionar_admin_automatico(usuarios: pd.DataFrame) -> pd.DataFrame:
    if not usuarios.empty:
        return usuarios
    novo = pd.DataFrame([{"usuario": "admin", "senha": "030711", "perfil": "Admin"}])
    salvar_entidade("usuarios", novo, False)
    return novo


def tela_login() -> bool:
    iniciar_estado_sessao()
    if st.session_state["autenticado"]:
        return True

    st.title("🔒 Acesso Restrito — Painel de Gestão & DP")
    st.caption("Entre com usuário: admin / senha: 030711")
    usuarios, _ = carregar_entidade("usuarios")
    usuarios = provisionar_admin_automatico(usuarios)

    with st.form("form_login"):
        identificador = st.text_input("Usuário").strip().lower()
        senha = st.text_input("Senha", type="password")
        enviar = st.form_submit_button("Entrar")

    if enviar:
        if st.session_state["tentativas_login"] >= 5:
            st.error("Muitas tentativas nesta sessão. Atualize a página.")
            return False
        candidato = usuarios[usuarios["usuario"] == identificador]
        if candidato.empty or limpar_texto(candidato.iloc[0]["senha"]) != senha:
            st.session_state["tentativas_login"] += 1
            st.error("Usuário ou senha incorretos.")
            return False

        registro = candidato.iloc[0]
        perfil_usuario = registro.get("perfil", "Admin")
        if not perfil_usuario:
            perfil_usuario = "Admin"

        st.session_state.update({
            "autenticado": True, "usuario": identificador, "perfil": perfil_usuario, "tentativas_login": 0,
        })
        st.rerun()
    return False


def encerrar_sessao() -> None:
    for chave in ("autenticado", "usuario", "perfil", "tentativas_login"):
        st.session_state.pop(chave, None)
    st.rerun()


def exigir_admin() -> bool:
    if st.session_state.get("perfil") != "Admin":
        st.error("Esta operação exige perfil de Administrador.")
        return False
    return True


def alterar_minha_senha() -> None:
    with st.expander("Alterar minha senha"):
        atual = st.text_input("Senha atual", type="password", key="senha_atual")
        nova = st.text_input("Nova senha", type="password", key="senha_nova")
        confirmar = st.text_input("Confirmar nova senha", type="password", key="senha_conf")
        if st.button("Atualizar senha", key="btn_alterar_senha"):
            usuarios, _ = carregar_entidade("usuarios")
            posicao = usuarios.index[usuarios["usuario"] == st.session_state["usuario"]]
            if posicao.empty or limpar_texto(usuarios.loc[posicao[0], "senha"]) != atual:
                st.error("A senha atual está incorreta.")
            elif nova != confirmar:
                st.error("A confirmação não corresponde à nova senha.")
            else:
                usuarios.loc[posicao[0], "senha"] = nova
                salvar_entidade("usuarios", usuarios)
                st.success("Senha alterada com sucesso!")


def filtrar_setor(df: pd.DataFrame, setor: str) -> pd.DataFrame:
    if setor == "Todos os setores" or df.empty:
        return df.copy()
    return df[df["setor"] == setor].copy()


def tabela_exibicao(df: pd.DataFrame, campos: list[str]) -> pd.DataFrame:
    rotulos = {
        "matricula": "Matrícula", "funcionario": "Funcionário", "setor": "Setor", "cargo": "Cargo",
        "usuario": "Usuário", "perfil": "Perfil", "admissao": "Admissão", "nascimento": "Nascimento",
        "status": "Status", "ultimas_ferias": "Últimas férias", "data_retorno_ferias": "Retorno das férias",
        "data": "Data", "tipo": "Tipo", "dias": "Dias", "cid": "CID", "motivo": "Motivo", "origem": "Origem",
        "epi": "EPI", "detalhe_tamanho": "Detalhe/Tamanho", "responsavel": "Responsável", "tipo_registro": "Tipo de Registro",
        "tipo_evento": "Tipo de evento", "descricao": "Descrição", "autor": "Autor",
        "data_desligamento": "Data de Desligamento",
    }
    existente = [campo for campo in campos if campo in df.columns]
    resultado = df[existente].copy()
    for campo in ("admissao", "nascimento", "ultimas_ferias", "data_retorno_ferias", "data", "data_desligamento"):
        if campo in resultado.columns:
            resultado[campo] = resultado[campo].map(formatar_data)
    return resultado.rename(columns=rotulos)


def excel_bytes(df: pd.DataFrame) -> bytes:
    saida = io.BytesIO()
    with pd.ExcelWriter(saida, engine="openpyxl") as escritor:
        df.to_excel(escritor, index=False, sheet_name="Relatório")
    return saida.getvalue()


def pdf_bytes(titulo: str, df: pd.DataFrame) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    saida = io.BytesIO()
    estilos = getSampleStyleSheet()
    elementos = [
        Paragraph(escape(titulo), estilos["Title"]),
        Paragraph(f"Gerado em {datetime.now():%d/%m/%Y %H:%M}", estilos["Normal"]),
        Spacer(1, 12),
    ]
    cabecalho = [Paragraph(f"<b>{escape(str(c))}</b>", estilos["BodyText"]) for c in df.columns]
    linhas = [[Paragraph(escape(limpar_texto(v)), estilos["BodyText"]) for v in linha] for linha in df.fillna("").values.tolist()]
    tabela = Table([cabecalho, *linhas], repeatRows=1)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDE8E0")),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#9CB5A6")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elementos.append(tabela)
    SimpleDocTemplate(saida, pagesize=landscape(letter), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24).build(elementos)
    return saida.getvalue()


def bloco_exportacao(nome: str, df: pd.DataFrame) -> None:
    if df.empty:
        return
    exibicao = df.copy()
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("Baixar Excel", excel_bytes(exibicao), f"{nome}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        try:
            st.download_button("Baixar PDF", pdf_bytes(nome, exibicao), f"{nome}.pdf", "application/pdf")
        except ImportError:
            st.caption("Instale reportlab para habilitar o PDF.")


def opcoes_colaboradores(df: pd.DataFrame, apenas_ativos: bool = False) -> tuple[list[str], dict[str, str]]:
    base = df[df["status"] == "Ativo"].copy() if apenas_ativos else df.copy()
    base = base[base["matricula"] != ""].sort_values(["funcionario", "matricula"])
    mapa = {linha["matricula"]: f"{linha['funcionario']} · {linha['matricula']}" for _, linha in base.iterrows()}
    return list(mapa.keys()), mapa


def classificar_experiencia(admissao: Optional[date], hoje: date) -> tuple[str, Optional[int], Optional[int]]:
    if not admissao:
        return "Sem data de admissão", None, None
    dias_45 = ((admissao + timedelta(days=45)) - hoje).days
    dias_90 = ((admissao + timedelta(days=90)) - hoje).days
    if dias_90 < 0:
        return "90 dias vencido", dias_45, dias_90
    if dias_45 < 0:
        return "Em experiência após 45 dias", dias_45, dias_90
    return "Até 45 dias", dias_45, dias_90


def tela_dashboard(colaboradores: pd.DataFrame, faltas: pd.DataFrame, setor: str) -> None:
    hoje = date.today()
    base = filtrar_setor(colaboradores, setor)
    faltas_base = filtrar_setor(faltas, setor)
    ativos = base[base["status"] == "Ativo"]
    em_ferias = base[base["status"] == "Férias"]
    afastados = base[base["status"] == "Afastado"]
    desligados = base[base["status"] == "Desligado"]
    ocorrencias_hoje = faltas_base[faltas_base["data"].map(para_data) == hoje] if not faltas_base.empty else faltas_base
    ausencias_hoje = ocorrencias_hoje[ocorrencias_hoje["tipo"] != "Folga Concedida"] if not ocorrencias_hoje.empty else ocorrencias_hoje

    st.subheader("Painel geral de indicadores")
    
    c1, c2, c3, c4, c5 = st.columns(5)
    cards = [
        ("Total no quadro", len(base)),
        ("Ativos", len(ativos)),
        ("Em férias", len(em_ferias)),
        ("Afastados / INSS", len(afastados)),
        ("Ausências hoje", len(ausencias_hoje))
    ]
    
    cols = [c1, c2, c3, c4, c5]
    for col, (titulo, valor) in zip(cols, cards):
        col.markdown(f"""
            <div style="background-color: #1A1D24; border: 2px solid #F97316; border-radius: 12px; padding: 15px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                <div style="font-size: 14px; font-weight: 600; color: #F8FAFC; margin-bottom: 5px;">{titulo}</div>
                <div style="font-size: 28px; font-weight: 800; color: #F97316;">{valor}</div>
            </div>
        """, unsafe_allow_html=True)

    # --- CARDS DE TOTAL POR SETOR ---
    st.markdown("### Total por Setor")
    if not colaboradores.empty and "setor" in colaboradores.columns:
        contagem_setores = base[base["setor"] != ""].groupby("setor").size().reset_index(name="total")
        if not contagem_setores.empty:
            setor_cols = st.columns(min(len(contagem_setores), 4))
            for i, (_, row) in enumerate(contagem_setores.iterrows()):
                col_atual = setor_cols[i % len(setor_cols)]
                col_atual.markdown(f"""
                    <div style="background-color: #1A1D24; border: 2px solid #3B82F6; border-radius: 12px; padding: 12px; text-align: center; margin-bottom: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                        <div style="font-size: 13px; font-weight: 600; color: #F8FAFC; margin-bottom: 4px;">{row['setor']}</div>
                        <div style="font-size: 24px; font-weight: 800; color: #3B82F6;">{row['total']}</div>
                    </div>
                """, unsafe_allow_html=True)
    # --------------------------------

    st.markdown("---")
    st.subheader("🚨 Alertas & Liberação de Vagas (Desligamentos)")
    
    if desligados.empty:
        st.success("Nenhum colaborador desligado registrado no momento.")
    else:
        tabela_desligados = tabela_exibicao(desligados, ["matricula", "funcionario", "setor", "cargo", "data_desligamento"])
        st.dataframe(tabela_desligados, use_container_width=True, hide_index=True)
        bloco_exportacao("vagas_liberadas_desligamentos", tabela_desligados)


def tela_chamada(colaboradores: pd.DataFrame, faltas: pd.DataFrame, setor: str, autor: str) -> None:
    st.subheader("Chamada diária e ocorrências")
    aba_chamada, aba_avulso, aba_historico = st.tabs(["Chamada diária", "Lançamento avulso", "Histórico"])
    base = filtrar_setor(colaboradores, setor)
    ativos = base[base["status"] == "Ativo"].copy()
    termos_lideranca = r"gerente|supervisor|encarregado|coordenador|líder|lider"
    operacionais = ativos[~ativos["cargo"].str.lower().str.contains(termos_lideranca, na=False)].copy()

    with aba_chamada:
        data_chamada = st.date_input("Data da chamada", value=date.today(), key="data_chamada", format="DD/MM/YYYY")
        if operacionais.empty:
            st.info("Não há colaboradores operacionais ativos para o filtro atual.")
        else:
            anteriores = faltas[(faltas["data"].map(para_data) == data_chamada) & (faltas["origem"] == "Chamada")]
            anteriores = filtrar_setor(anteriores, setor)
            estados: dict[str, str] = {linha["matricula"]: "" for _, linha in operacionais.iterrows()}
            for _, ocorrencia in anteriores.iterrows():
                estados[ocorrencia["matricula"]] = "Folga" if ocorrencia["tipo"] == "Folga Concedida" else "Ausente"

            with st.form("form_chamada"):
                novo_estado: dict[str, str] = {}
                for _, pessoa in operacionais.iterrows():
                    col_nome, col_status = st.columns([2.2, 1.35])
                    col_nome.markdown(f"**{pessoa['funcionario']}**  \n`{pessoa['matricula']}`")
                    
                    estado_atual = estados.get(pessoa["matricula"], "")
                    opcoes_radio = ("", "Presente", "Folga", "Ausente")
                    idx_opcao = opcoes_radio.index(estado_atual) if estado_atual in opcoes_radio else 0
                    
                    novo_estado[pessoa["matricula"]] = col_status.radio(
                        "Status", opcoes_radio,
                        index=idx_opcao,
                        horizontal=True, label_visibility="collapsed",
                        key=f"chamada_{data_chamada.isoformat()}_{setor}_{pessoa['matricula']}",
                    )
                salvar = st.form_submit_button("Salvar chamada")

            if salvar:
                remover = (faltas["data"].map(para_data) == data_chamada) & (faltas["origem"] == "Chamada")
                if setor != "Todos os setores":
                    remover &= faltas["setor"] == setor
                atualizada = faltas[~remover].copy()
                novos: list[dict[str, Any]] = []
                for _, pessoa in operacionais.iterrows():
                    escolha = novo_estado[pessoa["matricula"]]
                    if escolha and escolha != "Presente":
                        novos.append({
                            "registro_id": str(uuid.uuid4()), "matricula": pessoa["matricula"], "funcionario": pessoa["funcionario"],
                            "setor": pessoa["setor"], "data": data_chamada.isoformat(),
                            "tipo": "Folga Concedida" if escolha == "Folga" else "Falta Injustificada",
                            "dias": 1, "cid": "", "motivo": "Chamada diária", "origem": "Chamada",
                        })
                atualizada = pd.concat([atualizada, pd.DataFrame(novos)], ignore_index=True) if novos else atualizada
                if salvar_entidade("faltas", atualizada):
                    registrar_historico("", "", "Chamada diária", f"Chamada de {data_chamada:%d/%m/%Y} registrada para {setor}.", autor)
                    st.rerun()

    with aba_avulso:
        opcoes, mapa = opcoes_colaboradores(base, apenas_ativos=False)
        if not opcoes:
            st.info("Cadastre um colaborador antes de lançar uma ocorrência.")
        else:
            matricula = st.selectbox("Colaborador", opcoes, format_func=lambda valor: mapa[valor], key="av_mat")
            
            # Opções base de ocorrência
            tipo_base = st.selectbox("Tipo de Ocorrência", TIPOS_OCORRENCIA, key="av_tipo_base")
            
            # Se for Falta Injustificada, exibe a opção complementar de medida disciplinar
            tipo_final = tipo_base
            if tipo_base == "Falta Injustificada":
                medida_disciplinar = st.selectbox(
                    "Ação Disciplinar Vinculada",
                    ("Nenhuma (Apenas Falta)", "Advertência Escrita", "Suspensão"),
                    key="av_medida"
                )
                if medida_disciplinar != "Nenhuma (Apenas Falta)":
                    tipo_final = medida_disciplinar

            data_ocorrencia = st.date_input("Data", value=date.today(), key="data_ocorrencia", format="DD/MM/YYYY")
            dias = st.number_input("Quantidade de dias", min_value=1, max_value=365, value=1, key="av_dias")
            cid = st.text_input("CID (opcional)", key="av_cid").strip().upper()
            motivo = st.text_area("Observação / Descrição", max_chars=500, key="av_motivo").strip()

            if st.button("Salvar ocorrência avulsa", key="btn_salvar_avulso"):
                pessoa = base[base["matricula"] == matricula].iloc[0]
                novo = {
                    "registro_id": str(uuid.uuid4()), "matricula": matricula, "funcionario": pessoa["funcionario"],
                    "setor": pessoa["setor"], "data": data_ocorrencia.isoformat(), "tipo": tipo_final, "dias": int(dias),
                    "cid": cid, "motivo": motivo, "origem": "Avulso",
                }
                if salvar_entidade("faltas", pd.concat([faltas, pd.DataFrame([novo])], ignore_index=True)):
                    registrar_historico(matricula, pessoa["funcionario"], f"Registro: {tipo_final}", f"{motivo or tipo_final} em {data_ocorrencia:%d/%m/%Y}.", autor)
                    st.success(f"'{tipo_final}' registrada com sucesso e gravada no histórico!")
                    st.rerun()

    with aba_historico:
        historico = filtrar_setor(faltas, setor).sort_values("data", ascending=False)
        tabela = tabela_exibicao(historico, ["data", "funcionario", "setor", "tipo", "dias", "cid", "motivo", "origem"])
        st.dataframe(tabela, use_container_width=True, hide_index=True)
        bloco_exportacao("historico_faltas", tabela)


def tela_epi(colaboradores: pd.DataFrame, epis: pd.DataFrame, setor: str, autor: str) -> None:
    st.subheader("Solicitação & Entrega de EPI")
    aba_entrega, aba_solicitacao = st.tabs(["Registro de Entrega", "Solicitação ao RH"])
    base = filtrar_setor(colaboradores, setor)
    opcoes, mapa = opcoes_colaboradores(base, apenas_ativos=True)

    with aba_entrega:
        if not opcoes:
            st.info("Não há colaboradores ativos para o filtro atual.")
        else:
            with st.form("form_epi_entrega", clear_on_submit=True):
                matricula = st.selectbox("Colaborador", opcoes, format_func=lambda valor: mapa[valor], key="epi_ent_mat")
                c1, c2, c3 = st.columns(3)
                epi = c1.selectbox("EPI", ("Camiseta", "Bota de segurança", "Luvas", "Óculos", "Protetor auricular", "Outro"))
                detalhe = c2.text_input("Tamanho/Detalhe", max_chars=100)
                data_entrega = c3.date_input("Data da entrega", value=date.today(), format="DD/MM/YYYY")
                if st.form_submit_button("Registrar Entrega"):
                    pessoa = base[base["matricula"] == matricula].iloc[0]
                    novo = {
                        "entrega_id": str(uuid.uuid4()), "matricula": matricula, "funcionario": pessoa["funcionario"],
                        "setor": pessoa["setor"], "data": data_entrega.isoformat(), "epi": epi,
                        "detalhe_tamanho": detalhe, "responsavel": autor, "tipo_registro": "Entrega",
                    }
                    if salvar_entidade("epis", pd.concat([epis, pd.DataFrame([novo])], ignore_index=True)):
                        registrar_historico(matricula, pessoa["funcionario"], "Entrega de EPI", f"{epi}: {detalhe or 'sem detalhe'}.", autor)
                        st.rerun()

        st.markdown("#### Histórico de entregas")
        entregas_df = epis[epis["tipo_registro"] == "Entrega"] if "tipo_registro" in epis.columns else epis
        tabela_ent = tabela_exibicao(filtrar_setor(entregas_df, setor).sort_values("data", ascending=False), ["data", "funcionario", "setor", "epi", "detalhe_tamanho", "responsavel"])
        st.dataframe(tabela_ent, use_container_width=True, hide_index=True)
        bloco_exportacao("relatorio_entregas_epi", tabela_ent)

    with aba_solicitacao:
        if not opcoes:
            st.info("Não há colaboradores ativos para o filtro atual.")
        else:
            with st.form("form_epi_solicitacao", clear_on_submit=True):
                matricula = st.selectbox("Colaborador", opcoes, format_func=lambda valor: mapa[valor], key="epi_sol_mat")
                c1, c2 = st.columns(2)
                epi = c1.selectbox("EPI Solicitado", ("Camiseta", "Bota de segurança", "Luvas", "Óculos", "Protetor auricular", "Outro"), key="sol_epi")
                detalhe = c2.text_input("Motivo / Tamanho necessário", max_chars=150, key="sol_det")
                data_sol = st.date_input("Data da Solicitação", value=date.today(), key="sol_dt", format="DD/MM/YYYY")
                if st.form_submit_button("Gerar Solicitação ao RH"):
                    pessoa = base[base["matricula"] == matricula].iloc[0]
                    novo = {
                        "entrega_id": str(uuid.uuid4()), "matricula": matricula, "funcionario": pessoa["funcionario"],
                        "setor": pessoa["setor"], "data": data_sol.isoformat(), "epi": epi,
                        "detalhe_tamanho": detalhe, "responsavel": autor, "tipo_registro": "Solicitação",
                    }
                    if salvar_entidade("epis", pd.concat([epis, pd.DataFrame([novo])], ignore_index=True)):
                        registrar_historico(matricula, pessoa["funcionario"], "Solicitação de EPI", f"Solicitado ao RH - {epi}: {detalhe}", autor)
                        st.success("Solicitação gerada e encaminhada ao RH com sucesso!")
                        st.rerun()

        st.markdown("#### Solicitações pendentes ao RH")
        sol_df = epis[epis["tipo_registro"] == "Solicitação"] if "tipo_registro" in epis.columns else pd.DataFrame(columns=epis.columns)
        if not sol_df.empty:
            tabela_sol = tabela_exibicao(filtrar_setor(sol_df, setor).sort_values("data", ascending=False), ["data", "funcionario", "setor", "epi", "detalhe_tamanho", "responsavel"])
            st.dataframe(tabela_sol, use_container_width=True, hide_index=True)
            bloco_exportacao("solicitacoes_epi_rh", tabela_sol)
        else:
            st.info("Nenhuma solicitação de EPI registrada.")


def tela_ficha(colaboradores: pd.DataFrame, historico: pd.DataFrame) -> None:
    st.subheader("Ficha individual do colaborador")
    opcoes, mapa = opcoes_colaboradores(colaboradores)
    if not opcoes:
        st.info("Nenhum colaborador cadastrado.")
        return
    matricula = st.selectbox("Colaborador", opcoes, format_func=lambda valor: mapa[valor])
    pessoa = colaboradores[colaboradores["matricula"] == matricula].iloc[0]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Matrícula", pessoa["matricula"])
    c2.metric("Cargo", pessoa["cargo"] or "Não informado")
    c3.metric("Setor", pessoa["setor"] or "Não informado")
    c4.metric("Status", pessoa["status"])
    st.markdown(f"**Admissão:** {formatar_data(pessoa['admissao'])} &nbsp;&nbsp; **Nascimento:** {formatar_data(pessoa['nascimento'])} &nbsp;&nbsp; **Últimas Férias:** {formatar_data(pessoa['ultimas_ferias'])} &nbsp;&nbsp; **Retorno:** {formatar_data(pessoa['data_retorno_ferias'])}")
    eventos = historico[historico["matricula"] == matricula].sort_values("data", ascending=False)
    st.markdown("#### Linha do tempo")
    if eventos.empty:
        st.caption("Nenhum evento registrado para este colaborador.")
    else:
        st.dataframe(tabela_exibicao(eventos, ["data", "tipo_evento", "descricao", "autor"]), use_container_width=True, hide_index=True)


def tela_experiencia(colaboradores: pd.DataFrame, setor: str, autor: str) -> None:
    st.subheader("Controle de experiência (45 / 90 dias) & Contrato de Trabalho")
    hoje = date.today()
    base = filtrar_setor(colaboradores, setor)
    
    pendentes = []
    for _, pessoa in base.iterrows():
        if pessoa["status"] != "Ativo":
            continue
        admissao = para_data(pessoa["admissao"])
        if not admissao:
            continue
        _, dias45, dias90 = classificar_experiencia(admissao, hoje)
        decisao = limpar_texto(pessoa["decisao_experiencia"])
        if dias90 is not None and dias90 >= -10 and decisao.lower() != "efetivado":
            pendentes.append(pessoa)

    if not pendentes:
        st.success("Nenhum colaborador dentro do período de contrato de experiência pendente de avaliação.")
        return

    for pessoa in pendentes:
        admissao = para_data(pessoa["admissao"])
        _, dias45, dias90 = classificar_experiencia(admissao, hoje)
        situacao, _, _ = classificar_experiencia(admissao, hoje)
        
        with st.container(border=True):
            col_info, col_acoes = st.columns([2, 2.5])
            with col_info:
                st.markdown(f"**{pessoa['funcionario']}** (`{pessoa['matricula']}`)")
                st.caption(f"Setor: {pessoa['setor']} | Cargo: {pessoa['cargo']}")
                st.markdown(f"Admissão: **{formatar_data(admissao)}**")
                st.markdown(f"Situação: **{situacao}** (90 dias em {dias90} dias)")
            
            with col_acoes:
                with st.form(f"form_contrato_{pessoa['matricula']}"):
                    decisao_atual = pessoa["decisao_experiencia"] if pessoa["decisao_experiencia"] in ("", "Em avaliação", "Efetivado", "Desligado por quebra de contrato") else ""
                    opcoes_contrato = ("", "Em avaliação", "Efetivado", "Desligado por quebra de contrato")
                    idx_contrato = opcoes_contrato.index(decisao_atual) if decisao_atual in opcoes_contrato else 0
                    
                    nova_decisao = st.selectbox("Contrato de Trabalho", opcoes_contrato, index=idx_contrato, format_func=lambda x: "Selecione uma opção" if x == "" else x)
                    aviso_rh = st.checkbox("✅ Enviar aviso formal de fechamento de contrato ao RH")
                    
                    if st.form_submit_button("Salvar Decisão Contratual"):
                        idx = colaboradores.index[colaboradores["matricula"] == pessoa["matricula"]][0]
                        colaboradores.loc[idx, "decisao_experiencia"] = nova_decisao
                        if nova_decisao == "Desligado por quebra de contrato":
                            colaboradores.loc[idx, "status"] = "Desligado"
                            colaboradores.loc[idx, "data_desligamento"] = date.today().isoformat()
                        
                        if salvar_entidade("colaboradores", colaboradores):
                            msg_aviso = " com aviso formal ao RH" if aviso_rh else ""
                            registrar_historico(pessoa["matricula"], pessoa["funcionario"], "Contrato Experiência", f"Definição: {nova_decisao}{msg_aviso}", autor)
                            st.success("Atualizado com sucesso!")
                            st.rerun()


def gerar_pdf_ferias(titulo, df_escala):
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontSize=13, leading=16,
        textColor=colors.HexColor("#1E3A8A"), spaceAfter=5
    )
    sub_style = ParagraphStyle(
        'SubStyle', parent=styles['Normal'], fontSize=8, leading=11,
        textColor=colors.HexColor("#334155"), spaceAfter=10
    )

    hoje_txt = datetime.now().strftime("%d/%m/%Y às %H:%M")
    elements.append(Paragraph(f"<b>{titulo}</b>", title_style))
    elements.append(Paragraph(
        f"<b>Gerado em:</b> {hoje_txt} | <b>Empresa:</b> Tropical Distribuidora<br/>"
        f"<b>Regras de Negócio:</b> Início aos Domingos | Cota Máx: 2 colabs/mês por setor (Fev-Nov) | Coletivas Livres (Dez-Jan)",
        sub_style
    ))
    elements.append(Spacer(1, 5))

    colunas = list(df_escala.columns)
    table_data = [[Paragraph(f"<b><font size=7.5>{col}</font></b>", styles['Normal']) for col in colunas]]
    
    for _, linha in df_escala.iterrows():
        row_data = []
        for item in linha:
            val_str = str(item) if pd.notnull(item) else ""
            row_data.append(Paragraph(f"<font size=7>{val_str}</font>", styles['Normal']))
        table_data.append(row_data)

    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E2E8F0")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor("#1E293B")),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#94A3B8")),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    
    elements.append(t)
    doc.build(elements)
    pdf_out = buffer.getvalue()
    buffer.close()
    return pdf_out


def converter_df_para_excel(df_exp):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_exp.to_excel(writer, index=False, sheet_name='Escala_Ferias')
    return output.getvalue()


def ajustar_para_domingo(data_val):
    if isinstance(data_val, str):
        data_val = pd.to_datetime(data_val, dayfirst=True, errors='coerce').date()
    if not isinstance(data_val, date) or pd.isnull(data_val):
        return date.today()
    dias_para_domingo = (6 - data_val.weekday()) % 7
    return data_val + timedelta(days=dias_para_domingo)


def tela_ferias_escala(colaboradores: pd.DataFrame, autor: str) -> None:
    st.subheader("🏖️ Módulo de Pré-Agendamento & Simulação Flexível de Férias")

    hoje = date.today()
    df_ativos = colaboradores[colaboradores['status'].isin(['Ativo', 'Férias'])].copy()

    if df_ativos.empty:
        st.info("Não há colaboradores ativos para escalonamento.")
        return

    for col_req in ['aprovacao_rh', 'fracionamento', 'escala_confirmada', 'data_pre_agendada']:
        if col_req not in colaboradores.columns:
            colaboradores[col_req] = 'Pendente' if col_req != 'escala_confirmada' else False

    lista_temp = []
    regulares_cnt = 0
    atencao_cnt = 0
    vencidos_cnt = 0

    for idx, r in df_ativos.iterrows():
        adm = para_data(r.get('admissao'))
        ult_ferias = para_data(r.get('ultimas_ferias'))
        data_base = ult_ferias if ult_ferias else adm

        if data_base:
            anos = (hoje - data_base).days // 365
            inicio_aq = data_base + timedelta(days=365 * max(0, anos))
            fim_aq = inicio_aq + timedelta(days=365)
            limite_conc = fim_aq + timedelta(days=365)
            dias_restantes = (limite_conc - hoje).days

            if dias_restantes <= 0:
                vencidos_cnt += 1
            elif dias_restantes <= 60:
                atencao_cnt += 1
            else:
                regulares_cnt += 1

            lista_temp.append({
                'idx': idx, 'colab': r, 'data_base': data_base,
                'limite_conc': limite_conc, 'dias_restantes': dias_restantes, 'ult_ferias': ult_ferias
            })

    lista_temp = sorted(lista_temp, key=lambda x: x['limite_conc'])

    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric("🟢 Férias Regulares", regulares_cnt)
    c_m2.metric("🟡 Atenção (Próximos 60d)", atencao_cnt)
    c_m3.metric("🚨 Vencidos / Risco Multa", vencidos_cnt)

    st.markdown("---")

    tab_pre, tab_resumo, tab_export = st.tabs([
        "📅 Pré-Agendamento & Reorganização Flexível", 
        "📋 Escala Consolidada Auditável", 
        "📥 Impressão e Exportação"
    ])

    ocupacao_setor_mes = {}
    dados_escala = []
    alterou = False

    with tab_pre:
        st.markdown("##### 🗓️ Escolha ou simule a data do colaborador. O sistema ajusta automaticamente para o primeiro Domingo e valida a capacidade do setor:")

        for ordem_prio, item in enumerate(lista_temp, start=1):
            idx = item['idx']
            r = item['colab']
            limite_conc = item['limite_conc']
            ult_ferias = item['ult_ferias']
            setor = r.get('setor', 'Geral')

            data_sugerida_inicial = max(hoje + timedelta(days=30), limite_conc - timedelta(days=60))
            data_algoritmo = ajustar_para_domingo(data_sugerida_inicial)

            dt_pre_salva = r.get('data_pre_agendada')
            if dt_pre_salva and str(dt_pre_salva).strip() != "":
                data_efetiva = para_data(dt_pre_salva) or data_algoritmo
            else:
                data_efetiva = data_algoritmo

            adm_dt = para_data(r.get('admissao'))
            dt_adm_str = adm_dt.strftime('%d/%m/%Y') if adm_dt else 'N/A'
            dt_ult_str = ult_ferias.strftime('%d/%m/%Y') if ult_ferias else 'Não Registrada'

            c_info, c_dt, c_status = st.columns([2.2, 1.3, 1.5])

            with c_info:
                st.markdown(f"👤 **{r['funcionario']}** (Prio #{ordem_prio}) | Setor: **{setor}** | Cargo: {r.get('cargo', 'N/A')}")
                st.caption(f"Admissão: **{dt_adm_str}** | Últs Férias: **{dt_ult_str}** | Limite Concessivo: **{limite_conc.strftime('%d/%m/%Y')}**")

            with c_dt:
                nova_dt_input = st.date_input(
                    "Pré-Agendar para:",
                    value=data_efetiva,
                    key=f"pre_dt_{idx}",
                    format="DD/MM/YYYY"
                )
                data_inicio_domingo = ajustar_para_domingo(nova_dt_input)
                if data_inicio_domingo != data_efetiva:
                    colaboradores.at[idx, 'data_pre_agendada'] = data_inicio_domingo.isoformat()
                    alterou = True

            chave_mes = (setor, data_inicio_domingo.strftime("%Y-%m"))
            qtd_agendada = ocupacao_setor_mes.get(chave_mes, 0)
            is_férias_coletivas = data_inicio_domingo.month in [12, 1]
            ocupacao_setor_mes[chave_mes] = qtd_agendada + 1

            data_aviso_rh = data_inicio_domingo - timedelta(days=30)

            with c_status:
                if data_inicio_domingo > limite_conc:
                    st.error("🚨 Ultrapassa Limite Legal!")
                elif qtd_agendada >= 2 and not is_férias_coletivas:
                    st.warning(f"⚠️ Cota Excedida ({qtd_agendada + 1}º no mês)")
                else:
                    st.success(f"✅ Domingo: **{data_inicio_domingo.strftime('%d/%m/%Y')}**")

                st.caption(f"Aviso RH até: **{data_aviso_rh.strftime('%d/%m/%Y')}**")

            c_frac, c_aprov, c_conf = st.columns([1.5, 1.5, 1])

            with c_frac:
                frac_atual = r.get('fracionamento') if pd.notnull(r.get('fracionamento')) else '30 Dias Corridos'
                opcoes_frac = ['30 Dias Corridos', '15 + 15 Dias', '20 + 10 Dias']
                idx_f = opcoes_frac.index(frac_atual) if frac_atual in opcoes_frac else 0
                novo_frac = st.selectbox("Fracionamento", opcoes_frac, index=idx_f, key=f"frac_{idx}")
                if novo_frac != frac_atual:
                    colaboradores.at[idx, 'fracionamento'] = novo_frac
                    alterou = True

            with c_aprov:
                aprov_atual = r.get('aprovacao_rh') if pd.notnull(r.get('aprovacao_rh')) else 'Pendente'
                opcoes_aprov = ['Pendente', 'Pré-Agendado RH', 'Aprovado RH', 'Em Análise', 'Rejeitado']
                idx_a = opcoes_aprov.index(aprov_atual) if aprov_atual in opcoes_aprov else 0
                nova_aprov = st.selectbox("Status RH", opcoes_aprov, index=idx_a, key=f"aprov_{idx}")
                if nova_aprov != aprov_atual:
                    colaboradores.at[idx, 'aprovacao_rh'] = nova_aprov
                    alterou = True

            with c_conf:
                conf_atual = bool(r.get('escala_confirmada')) if pd.notnull(r.get('escala_confirmada')) else False
                nova_conf = st.checkbox("Confirmar?", value=conf_atual, key=f"conf_{idx}")
                if nova_conf != conf_atual:
                    colaboradores.at[idx, 'escala_confirmada'] = nova_conf
                    alterou = True

            st.divider()

            cota_txt = f"Vaga {qtd_agendada + 1}/2 no Setor" if not is_férias_coletivas else "Férias Coletivas (Sem Cota)"

            dados_escala.append({
                "Prio": f"#{ordem_prio}",
                "Matrícula": r.get('matricula', 'N/A'),
                "Funcionário": r['funcionario'],
                "Setor": setor,
                "Cargo": r.get('cargo', 'N/A'),
                "Admissão": dt_adm_str,
                "Últimas Férias": dt_ult_str,
                "Início Férias (Domingo)": data_inicio_domingo.strftime('%d/%m/%Y'),
                "Prazo Aviso RH": data_aviso_rh.strftime('%d/%m/%Y'),
                "Limite Concessivo": limite_conc.strftime('%d/%m/%Y'),
                "Status Cota Setor": cota_txt,
                "Fracionamento": colaboradores.at[idx, 'fracionamento'],
                "Status RH": colaboradores.at[idx, 'aprovacao_rh'],
                "Confirmado": "SIM" if colaboradores.at[idx, 'escala_confirmada'] else "NÃO"
            })

    if alterou:
        salvar_entidade("colaboradores", colaboradores, mostrar_feedback=False)
        st.success("✅ Pré-agendamentos e alterações salvos com sucesso!")
        st.rerun()

    df_escala_final = pd.DataFrame(dados_escala)

    with tab_resumo:
        st.markdown("##### 📋 Visão Consolidada da Escala de Férias Reorganizada")
        st.dataframe(df_escala_final, use_container_width=True, hide_index=True)

    with tab_export:
        st.markdown("##### 📥 Exportar Relatório com Datas Pré-Agendadas")
        c_down1, c_down2 = st.columns(2)
        
        with c_down1:
            st.download_button(
                label="📥 Baixar Escala em Excel (.xlsx)",
                data=converter_df_para_excel(df_escala_final),
                file_name=f"escala_inteligente_ferias_{hoje.strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_excel_pre"
            )
            
        with c_down2:
            pdf_bytes_out = gerar_pdf_ferias(
                "Relatório - Escala Inteligente de Férias e Pré-Agendamento RH",
                df_escala_final
            )
            st.download_button(
                label="🖨️ Baixar PDF para Impressão",
                data=pdf_bytes_out,
                file_name=f"escala_inteligente_ferias_{hoje.strftime('%d_%m_%Y')}.pdf",
                mime="application/pdf",
                key="btn_pdf_pre"
            )


def tela_indicadores(colaboradores: pd.DataFrame, faltas: pd.DataFrame, setor: str) -> None:
    st.subheader("Indicadores de frequência e absenteísmo")
    c1, c2 = st.columns(2)
    inicio = c1.date_input("Início do período", value=date.today().replace(day=1), format="DD/MM/YYYY")
    fim = c2.date_input("Fim do período", value=date.today(), format="DD/MM/YYYY")
    if fim < inicio:
        st.error("O fim do período precisa ser igual ou posterior ao início.")
        return
    base_faltas = filtrar_setor(faltas, setor).copy()
    base_faltas["data_dt"] = pd.to_datetime(base_faltas["data"], errors="coerce")
    periodo = base_faltas[(base_faltas["data_dt"] >= pd.Timestamp(inicio)) & (base_faltas["data_dt"] <= pd.Timestamp(fim))]
    ausencias = periodo[~periodo["tipo"].isin(["Folga Concedida"])]
    dias_ausentes = int(ausencias["dias"].sum()) if not ausencias.empty else 0
    quadro_medio = len(filtrar_setor(colaboradores, setor).query("status == 'Ativo'"))
    dias_calendario = (fim - inicio).days + 1
    taxa = (dias_ausentes / (quadro_medio * dias_calendario) * 100) if quadro_medio and dias_calendario else 0
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Dias de ausência", dias_ausentes)
    m2.metric("Atestados", int((periodo["tipo"] == "Atestado Médico").sum()))
    m3.metric("Faltas injustificadas", int((periodo["tipo"] == "Falta Injustificada").sum()))
    m4.metric("Taxa no período", f"{taxa:.2f}%")


def tela_colaboradores(colaboradores: pd.DataFrame, autor: str) -> None:
    if not exigir_admin():
        return
    st.subheader("Gestão de colaboradores")
    aba_novo, aba_editar = st.tabs(["Novo colaborador", "Editar / desligar"])
    with aba_novo:
        with st.form("novo_colaborador", clear_on_submit=True):
            c1, c2 = st.columns(2)
            matricula = c1.text_input("Matrícula").strip()
            nome = c2.text_input("Nome completo").strip()
            c3, c4 = st.columns(2)
            setor = c3.text_input("Setor").strip()
            cargo = c4.text_input("Cargo").strip()
            c5, c6, c7, c8 = st.columns(4)
            admissao = c5.date_input("Admissão", value=date.today(), format="DD/MM/YYYY")
            nascimento = c6.date_input("Nascimento", value=date(1990, 1, 1), format="DD/MM/YYYY")
            ultimas_ferias = c7.date_input("Últimas Férias (opcional)", value=None, format="DD/MM/YYYY")
            status = c8.selectbox("Status", STATUS_COLABORADOR)
            
            if st.form_submit_button("Cadastrar"):
                if not matricula or not nome or not setor:
                    st.error("Matrícula, nome e setor são obrigatórios.")
                elif (colaboradores["matricula"] == limpar_matricula(matricula)).any():
                    st.error("Já existe um colaborador com esta matrícula.")
                else:
                    novo = {
                        "matricula": limpar_matricula(matricula), "funcionario": nome, "setor": setor, "cargo": cargo,
                        "admissao": admissao.isoformat(), "nascimento": nascimento.isoformat(), "status": status,
                        "ultimas_ferias": ultimas_ferias.isoformat() if ultimas_ferias else "",
                        "data_retorno_ferias": "", "decisao_experiencia": "", "data_desligamento": "",
                    }
                    atualizada = pd.concat([colaboradores, pd.DataFrame([novo])], ignore_index=True)
                    if salvar_entidade("colaboradores", atualizada):
                        registrar_historico(novo["matricula"], nome, "Cadastro", "Colaborador cadastrado.", autor)
                        st.rerun()

    with aba_editar:
        opcoes, mapa = opcoes_colaboradores(colaboradores)
        if not opcoes:
            st.info("Nenhum colaborador disponível.")
            return
        matricula = st.selectbox("Colaborador para editar", opcoes, format_func=lambda valor: mapa[valor])
        indice = colaboradores.index[colaboradores["matricula"] == matricula][0]
        pessoa = colaboradores.loc[indice]
        with st.form("editar_colaborador"):
            c1, c2 = st.columns(2)
            nova_matricula = c1.text_input("Matrícula", value=pessoa["matricula"]).strip()
            nome = c2.text_input("Nome completo", value=pessoa["funcionario"]).strip()
            c3, c4 = st.columns(2)
            setor = c3.text_input("Setor", value=pessoa["setor"]).strip()
            cargo = c4.text_input("Cargo", value=pessoa["cargo"]).strip()
            
            c5, c6, c7 = st.columns(3)
            admissao = c5.date_input("Admissão", value=para_data(pessoa["admissao"]) or date.today(), format="DD/MM/YYYY")
            ultimas_ferias_atual = para_data(pessoa["ultimas_ferias"])
            ult_ferias = c6.date_input("Últimas Férias", value=ultimas_ferias_atual if ultimas_ferias_atual else None, format="DD/MM/YYYY")
            
            status_atual = pessoa["status"] if pessoa["status"] in STATUS_COLABORADOR else "Ativo"
            status = c7.selectbox("Status", STATUS_COLABORADOR, index=STATUS_COLABORADOR.index(status_atual))
            
            dt_deslig = para_data(pessoa["data_desligamento"]) or date.today()
            data_desligamento = st.date_input("Data do desligamento", value=dt_deslig, format="DD/MM/YYYY")
            
            if st.form_submit_button("Atualizar"):
                nova_matricula = limpar_matricula(nova_matricula)
                duplicada = (colaboradores["matricula"] == nova_matricula) & (colaboradores.index != indice)
                if not nova_matricula or not nome or not setor:
                    st.error("Matrícula, nome e setor são obrigatórios.")
                elif duplicada.any():
                    st.error("A matrícula informada já pertence a outro colaborador.")
                else:
                    colaboradores.loc[indice, ["matricula", "funcionario", "setor", "cargo", "admissao", "ultimas_ferias", "status", "data_desligamento"]] = [
                        nova_matricula, nome, setor, cargo, admissao.isoformat(),
                        ult_ferias.isoformat() if ult_ferias else "", status,
                        data_desligamento.isoformat() if status == "Desligado" else "",
                    ]
                    if salvar_entidade("colaboradores", colaboradores):
                        evento = "Desligamento" if status == "Desligado" else "Atualização cadastral"
                        registrar_historico(nova_matricula, nome, evento, f"Status atualizado para {status}.", autor)
                        st.rerun()


def tela_usuarios(usuarios: pd.DataFrame) -> None:
    if not exigir_admin():
        return
    st.subheader("Gestão de usuários")
    aba_novo, aba_lista = st.tabs(["Novo usuário", "Lista / Editar"])
    
    with aba_novo:
        with st.form("novo_usuario", clear_on_submit=True):
            usuario = st.text_input("Login", help="Será convertido para letras minúsculas.").strip().lower()
            senha = st.text_input("Senha inicial", type="password")
            perfil = st.selectbox("Perfil", ("Admin", "Gestor"))
            if st.form_submit_button("Criar usuário"):
                if not usuario or not senha:
                    st.error("Login e senha são obrigatórios.")
                elif (usuarios["usuario"] == usuario).any():
                    st.error("Já existe um usuário com este login.")
                else:
                    novo = {"usuario": usuario, "senha": senha, "perfil": perfil}
                    if salvar_entidade("usuarios", pd.concat([usuarios, pd.DataFrame([novo])], ignore_index=True)):
                        st.success("Usuário criado com sucesso!")
                        st.rerun()
                        
    with aba_lista:
        if usuarios.empty:
            st.info("Nenhum usuário cadastrado.")
            return
            
        selecionado = st.selectbox("Selecione o usuário para editar/excluir", usuarios["usuario"].tolist())
        if selecionado:
            indice = usuarios.index[usuarios["usuario"] == selecionado][0]
            reg = usuarios.loc[indice]
            perfil_atual = reg.get("perfil", "Admin")
            if perfil_atual not in ("Admin", "Gestor"):
                perfil_atual = "Admin"
            
            with st.form("form_edicao_usuario"):
                st.markdown(f"**Editando usuário:** `{selecionado}`")
                novo_login = st.text_input("Nome de usuário (Login)", value=selecionado).strip().lower()
                nova_senha = st.text_input("Nova senha (deixe em branco para não alterar)", type="password")
                novo_perfil = st.selectbox("Perfil", ("Admin", "Gestor"), index=("Admin", "Gestor").index(perfil_atual))
                
                c1, c2 = st.columns(2)
                salvar_edicao = c1.form_submit_button("Salvar alterações")
                excluir = c2.form_submit_button("Excluir usuário")
                
                if salvar_edicao:
                    duplicado = (usuarios["usuario"] == novo_login) & (usuarios.index != indice)
                    if not novo_login:
                        st.error("O login não pode ficar em branco.")
                    elif duplicado.any():
                        st.error("Já existe outro usuário com este login.")
                    else:
                        usuarios.loc[indice, "usuario"] = novo_login
                        if nova_senha.strip():
                            usuarios.loc[indice, "senha"] = nova_senha.strip()
                        usuarios.loc[indice, "perfil"] = novo_perfil
                        salvar_entidade("usuarios", usuarios)
                        st.success("Usuário atualizado com sucesso!")
                        st.rerun()
                        
                if excluir:
                    if len(usuarios) <= 1:
                        st.error("Você não pode excluir o único usuário restante do sistema.")
                    elif selecionado == st.session_state.get("usuario"):
                        st.error("Você não pode excluir seu próprio usuário logado no momento.")
                    else:
                        usuarios_atualizado = usuarios.drop(indice).reset_index(drop=True)
                        salvar_entidade("usuarios", usuarios_atualizado)
                        st.success("Usuário excluído com sucesso!")
                        st.rerun()


def tela_importacao(colaboradores: pd.DataFrame) -> None:
    if not exigir_admin():
        return
    st.subheader("Importar nova base de colaboradores")
    arquivo = st.file_uploader("Planilha .xlsx", type=["xlsx"])
    if not arquivo:
        return
    try:
        bruto = pd.read_excel(arquivo, dtype=object)
        importado, _ = normalizar_entidade(bruto, "colaboradores")
    except Exception as erro:
        st.error(f"Não foi possível processar a planilha: {erro}")
        return
    confirmar = st.checkbox("Confirmo que desejo importar e atualizar a base.")
    if st.button("Importar e atualizar base", disabled=not confirmar):
        combinado = colaboradores.set_index("matricula")
        atualizacoes = importado.set_index("matricula")
        combinado.update(atualizacoes)
        novos = atualizacoes.loc[~atualizacoes.index.isin(combinado.index)]
        resultado = pd.concat([combinado, novos]).reset_index()
        if salvar_entidade("colaboradores", resultado):
            st.success("Importação concluída com sucesso!")
            st.rerun()


def tela_assistente_ia(colaboradores: pd.DataFrame, faltas: pd.DataFrame, epis: pd.DataFrame, historico: pd.DataFrame) -> None:
    st.subheader("Assistente IA para DP e Gestão")
    chave = segredo("GEMINI_API_KEY")
    
    historico_chat = st.session_state.setdefault("historico_ia", [])
    for mensagem in historico_chat:
        with st.chat_message(mensagem["role"]):
            st.markdown(mensagem["content"])
            
    pergunta = st.chat_input("Pergunte sobre faltas, férias, escala ou dados da equipe")
    if pergunta:
        historico_chat.append({"role": "user", "content": pergunta})
        with st.chat_message("user"):
            st.markdown(pergunta)
            
        texto_lower = pergunta.lower()
        resposta_texto = ""
        hoje = date.today()
        
        if any(m in texto_lower for m in ["outubro", "novembro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "dezembro", "janeiro", "posso liberar", "cota", "vagas"]):
            meses_map = {
                "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4,
                "maio": 5, "junho": 6, "julho": 7, "agosto": 8, "setembro": 9,
                "outubro": 10, "novembro": 11, "dezembro": 12
            }
            mes_alvo = None
            for nome_m, num_m in meses_map.items():
                if nome_m in texto_lower:
                    mes_alvo = num_m
                    break
            
            if not colaboradores.empty:
                ativos = colaboradores[colaboradores["status"].isin(["Ativo", "Férias"])].copy()
                setores_disponiveis = ativos["setor"].dropna().unique()
                
                resposta_texto = f"📊 **Análise de Cota e Vagas para Férias**"
                if mes_alvo:
                    nome_mes_txt = [k for k, v in meses_map.items() if v == mes_alvo][0].capitalize()
                    resposta_texto += f" **(Mês de referência: {nome_mes_txt})**:\n\n"
                    
                    is_coletiva = mes_alvo in [12, 1]
                    if is_coletiva:
                        resposta_texto += "ℹ️ *Atenção:* Dezembro e Janeiro são meses de Férias Coletivas livres (sem limite estrito de cota por setor).\n\n"
                    else:
                        resposta_texto += "ℹ️ *Regra:* Limite padrão de **2 colaboradores por mês por setor**.\n\n"

                    ocupacao_setor = {s: [] for s in setores_disponiveis}
                    for _, r in ativos.iterrows():
                        dt_pre = para_data(r.get("data_pre_agendada"))
                        if dt_pre and dt_pre.month == mes_alvo:
                            s_colab = r.get("setor", "Geral")
                            if s_colab in ocupacao_setor:
                                ocupacao_setor[s_colab].append(r['funcionario'])

                    for setor_item, alocados in ocupacao_setor.items():
                        vagas_restantes = "Ilimitadas" if is_coletiva else max(0, 2 - len(alocados))
                        resposta_texto += f"- **Setor {setor_item}:** {len(alocados)} agendado(s) | Vagas restantes: **{vagas_restantes}**\n"
                        if alocados:
                            resposta_texto += f"  * Já escalados: {', '.join(alocados)}\n"
                else:
                    resposta_texto += "Informe o mês desejado (ex: 'quantos posso liberar em outubro?') para calcularmos as vagas por setor com base na Escala Inteligente.\n"
            else:
                resposta_texto = "Não há dados de colaboradores cadastrados para cruzar com a escala."

        elif "quem está em férias" in texto_lower or "quem esta em ferias" in texto_lower or "em férias" in texto_lower or "em ferias" in texto_lower:
            if not colaboradores.empty:
                em_ferias = colaboradores[colaboradores["status"] == "Férias"]
                resposta_texto = f"🏖️ **Colaboradores Atualmente em Férias:**\n"
                if not em_ferias.empty:
                    for _, r in em_ferias.iterrows():
                        ret = formatar_data(r.get("data_retorno_ferias"))
                        resposta_texto += f"- *{r['funcionario']}* ({r['setor']}) — Retorno previsto: **{ret}**\n"
                else:
                    resposta_texto += "Nenhum colaborador com status de férias ativado no momento."
            else:
                resposta_texto = "Sem dados de colaboradores."

        elif "falta" in texto_lower or "faltas" in texto_lower or "atestado" in texto_lower or "ausên" in texto_lower or "advertência" in texto_lower or "suspensão" in texto_lower:
            if not faltas.empty:
                total_faltas = len(faltas[~faltas["tipo"].isin(["Folga Concedida"])])
                atestados = len(faltas[faltas["tipo"] == "Atestado Médico"])
                injustificadas = len(faltas[faltas["tipo"] == "Falta Injustificada"])
                advertencias = len(faltas[faltas["tipo"] == "Advertência Escrita"])
                suspensoes = len(faltas[faltas["tipo"] == "Suspensão"])
                
                resposta_texto = f"📋 **Relatório de Ocorrências e Disciplinar:**\n"
                resposta_texto += f"- Total de registros: **{total_faltas}**\n"
                resposta_texto += f"- Atestados médicos: **{atestados}**\n"
                resposta_texto += f"- Faltas injustificadas: **{injustificadas}**\n"
                resposta_texto += f"- Advertências escritas: **{advertencias}**\n"
                resposta_texto += f"- Suspensões: **{suspensoes}**\n"
            else:
                resposta_texto = "Nenhuma ocorrência registrada."

        elif "equipe" in texto_lower or "colaboradores" in texto_lower or "entraram" in texto_lower or "ativos" in texto_lower:
            if not colaboradores.empty:
                total = len(colaboradores)
                ativos = len(colaboradores[colaboradores["status"] == "Ativo"])
                resposta_texto = f"👥 **Quadro Geral:** {total} cadastrados ({ativos} ativos)."
            else:
                resposta_texto = "Base vazia."

        else:
            if chave:
                try:
                    import google.generativeai as genai
                    genai.configure(api_key=chave)
                    modelo = genai.GenerativeModel("gemini-1.5-pro")
                    resp_ai = modelo.generate_content("Responda em português, de forma objetiva sobre DP e CLT: " + pergunta)
                    resposta_texto = getattr(resp_ai, "text", "Sem resposta.")
                except Exception as e:
                    resposta_texto = f"Erro ao consultar a IA: {e}"
            else:
                resposta_texto = (
                    f"💡 **Assistente de Gestão & DP:** Analisei sua solicitação (*'{pergunta}'*). "
                    "Posso cruzar dados de faltas, advertências, suspensões, férias ou escala inteligente."
                )

        with st.chat_message("assistant"):
            st.markdown(resposta_texto)
        historico_chat.append({"role": "assistant", "content": resposta_texto})


def aplicar_estilo() -> None:
    st.markdown("""
    <style>
      .stApp { background: #0E1117; color: #F8FAFC; }
      [data-testid="stSidebar"] { background: #163A2A; }
      [data-testid="stSidebar"] * { color: #F8FAFC !important; }
      div.stButton > button, div[data-testid="stFormSubmitButton"] > button {
        background: #F97316; color: white; border: 0; border-radius: 8px; font-weight: 700;
      }
    </style>
    """, unsafe_allow_html=True)


def main() -> None:
    aplicar_estilo()
    if not tela_login():
        return

    colaboradores, fonte = carregar_entidade("colaboradores")
    usuarios, _ = carregar_entidade("usuarios")
    faltas, _ = carregar_entidade("faltas")
    epis, _ = carregar_entidade("epis")
    historico, _ = carregar_entidade("historico")

    st.session_state["fonte_colaboradores"] = fonte
    nome = st.session_state["usuario"]
    perfil = st.session_state["perfil"]
    
    modulos = list(TODOS_MODULOS) if perfil == "Admin" else [m for m in TODOS_MODULOS if m not in MODULOS_ADMIN]

    st.sidebar.title("🍊 Gestão & DP")
    st.sidebar.caption(f"{nome} · {perfil}")
    st.sidebar.caption(f"Fonte atual: **{fonte}**")
    if st.sidebar.button("Sair"):
        encerrar_sessao()
    alterar_minha_senha()

    setores = ["Todos os setores"] + sorted([valor for valor in colaboradores["setor"].dropna().unique() if limpar_texto(valor)])
    setor = st.sidebar.selectbox("Filtrar por setor", setores)
    menu = st.sidebar.radio("Navegação", modulos)

    st.title("Painel de Gestão & DP")

    if menu == "Dashboard & Alertas":
        tela_dashboard(colaboradores, faltas, setor)
    elif menu == "Assistente IA (DP & Gestão)":
        tela_assistente_ia(colaboradores, faltas, epis, historico)
    elif menu == "Chamada & Faltas do Dia":
        tela_chamada(colaboradores, faltas, setor, nome)
    elif menu == "Solicitação & Entrega de EPI":
        tela_epi(colaboradores, epis, setor, nome)
    elif menu == "Ficha Individual do Colaborador":
        tela_ficha(colaboradores, historico)
    elif menu == "Controle de Experiência (45/90 dias)":
        tela_experiencia(colaboradores, setor, nome)
    elif menu == "Escala Inteligente de Férias & Folga":
        tela_ferias_escala(colaboradores, nome)
    elif menu == "Colaboradores em Férias":
        ferias = filtrar_setor(colaboradores[colaboradores["status"] == "Férias"], setor)
        st.subheader("Colaboradores em férias")
        st.dataframe(tabela_exibicao(ferias, ["matricula", "funcionario", "setor", "cargo", "ultimas_ferias", "data_retorno_ferias"]), use_container_width=True, hide_index=True)
    elif menu == "Indicadores de Frequência & Absenteísmo":
        tela_indicadores(colaboradores, faltas, setor)
    elif menu == "Aniversariantes do Mês":
        st.subheader("Aniversariantes do mês")
        meses_nomes = {
            1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
            7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
        }
        mes_atual = date.today().month
        mes_escolhido = st.selectbox("Mês", range(1, 13), index=mes_atual - 1, format_func=lambda numero: meses_nomes[numero])
        
        aniversariantes = filtrar_setor(colaboradores, setor).copy()
        if not aniversariantes.empty and "nascimento" in aniversariantes.columns:
            tabela_aniv = []
            for _, r in aniversariantes.iterrows():
                dt = para_data(r.get("nascimento"))
                if dt and dt.month == mes_escolhido:
                    tabela_aniv.append({
                        "dia_ordem": dt.day,
                        "Dia": f"{dt.day:02d}/{dt.month:02d}",
                        "Funcionário": r["funcionario"],
                        "Setor": r["setor"],
                        "Cargo": r["cargo"]
                    })
            if tabela_aniv:
                df_aniv = pd.DataFrame(tabela_aniv).sort_values("dia_ordem").drop(columns=["dia_ordem"])
                st.dataframe(df_aniv, use_container_width=True, hide_index=True)
            else:
                st.info("Nenhum aniversariante neste mês para o setor selecionado.")
        else:
            st.info("Nenhum registro de colaboradores encontrado.")
            
    elif menu == "Cadastrar / Editar Colaborador":
        tela_colaboradores(colaboradores, nome)
    elif menu == "Criar / Gerenciar Usuários":
        tela_usuarios(usuarios)
    elif menu == "Importar Nova Base":
        tela_importacao(colaboradores)


if __name__ == "__main__":
    main()
