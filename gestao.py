"""Painel de Gestão & DP — versão corrigida, blindada e preparada para o Supabase.

Execute com: streamlit run gestao_corrigido.py
Antes do primeiro uso, copie .streamlit/secrets.example.toml para
.streamlit/secrets.toml e aplique supabase_schema.sql, caso use Supabase.
"""

from __future__ import annotations

import base64
import binascii
import hashlib
import hmac
import io
import os
import re
import secrets
import shutil
import unicodedata
import uuid
from datetime import date, datetime, timedelta
from html import escape
from pathlib import Path
from typing import Any, Callable, Optional

import pandas as pd
import streamlit as st

try:
    from supabase import Client, create_client
except ImportError:  # Permite executar em modo local, sem Supabase instalado.
    Client = Any  # type: ignore[misc,assignment]
    create_client = None


# ---------------------------------------------------------------------------
# Configuração e contratos de dados
# ---------------------------------------------------------------------------
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
        "data_desligamento",
    ],
    "usuarios": ["usuario", "nome", "email", "senha_hash", "perfil", "modulos", "telefone", "ativo"],
    "faltas": [
        "registro_id", "matricula", "funcionario", "setor", "data", "tipo", "dias",
        "cid", "motivo", "origem",
    ],
    "epis": ["entrega_id", "matricula", "funcionario", "setor", "data", "epi", "detalhe_tamanho", "responsavel"],
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
    "Escala de Férias",
    "Colaboradores em Férias",
    "Indicadores de Frequência & Absenteísmo",
    "Aniversariantes do Mês",
    "Cadastrar / Editar Colaborador",
    "Criar / Gerenciar Usuários",
    "Importar Nova Base",
)

MODULOS_ADMIN = {"Criar / Gerenciar Usuários", "Importar Nova Base"}
PBKDF2_ITERACOES = 600_000

ALIAS_COLUNAS = {
    "colaboradores": {
        "matricula": "matricula", "matrícula": "matricula", "funcionario": "funcionario",
        "funcionário": "funcionario", "nome": "funcionario", "setor": "setor", "cargo": "cargo",
        "admissao": "admissao", "admissão": "admissao", "dt_adm": "admissao",
        "nascimento": "nascimento", "data_nascimento": "nascimento", "ultimas_ferias": "ultimas_ferias",
        "últimas_férias": "ultimas_ferias", "data_retorno_ferias": "data_retorno_ferias",
        "retorno_ferias": "data_retorno_ferias", "status": "status",
        "decisao_experiencia": "decisao_experiencia", "decisão_experiência": "decisao_experiencia",
        "data_desligamento": "data_desligamento",
    },
    "usuarios": {
        "usuario": "usuario", "usuário": "usuario", "login": "usuario", "nome": "nome",
        "email": "email", "e_mail": "email", "senha_hash": "senha_hash", "senha": "senha_legada",
        "perfil": "perfil", "modulos": "modulos", "módulos": "modulos", "telefone": "telefone", "ativo": "ativo",
    },
    "faltas": {
        "registro_id": "registro_id", "matricula": "matricula", "matrícula": "matricula",
        "funcionario": "funcionario", "funcionário": "funcionario", "setor": "setor", "data": "data",
        "tipo": "tipo", "dias": "dias", "cid": "cid", "motivo": "motivo", "origem": "origem",
    },
    "epis": {
        "entrega_id": "entrega_id", "matricula": "matricula", "matrícula": "matricula",
        "funcionario": "funcionario", "funcionário": "funcionario", "setor": "setor", "data": "data",
        "epi": "epi", "detalhe_tamanho": "detalhe_tamanho", "responsavel": "responsavel", "responsável": "responsavel",
    },
    "historico": {
        "evento_id": "evento_id", "matricula": "matricula", "matrícula": "matricula",
        "funcionario": "funcionario", "funcionário": "funcionario", "data": "data",
        "tipo_evento": "tipo_evento", "descricao": "descricao", "descrição": "descricao", "autor": "autor",
    },
}


def segredo(nome: str, padrao: str = "") -> str:
    """Lê uma configuração sem interromper o modo local quando secrets não existe."""
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
    return None if pd.isna(convertido) else convertido.date()


def data_iso(valor: object) -> Optional[str]:
    convertido = para_data(valor)
    return convertido.isoformat() if convertido else None


def formatar_data(valor: object) -> str:
    convertido = para_data(valor)
    return convertido.strftime("%d/%m/%Y") if convertido else "—"


def valor_bool(valor: object, padrao: bool = True) -> bool:
    texto = limpar_texto(valor).lower()
    if texto in {"false", "0", "nao", "não", "inativo"}:
        return False
    if texto in {"true", "1", "sim", "ativo"}:
        return True
    return padrao


def hash_senha(senha: str) -> str:
    if not senha:
        raise ValueError("A senha não pode ficar vazia.")
    salt = secrets.token_bytes(16)
    derivada = hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), salt, PBKDF2_ITERACOES)
    return "pbkdf2_sha256${}${}${}".format(
        PBKDF2_ITERACOES,
        base64.urlsafe_b64encode(salt).decode("ascii"),
        base64.urlsafe_b64encode(derivada).decode("ascii"),
    )


def verificar_senha(senha: str, senha_hash: str) -> bool:
    try:
        algoritmo, iteracoes, salt_b64, digest_b64 = senha_hash.split("$", 3)
        if algoritmo != "pbkdf2_sha256":
            return False
        derivada = hashlib.pbkdf2_hmac(
            "sha256", senha.encode("utf-8"), base64.urlsafe_b64decode(salt_b64), int(iteracoes)
        )
        return hmac.compare_digest(base64.urlsafe_b64encode(derivada).decode("ascii"), digest_b64)
    except (ValueError, TypeError, binascii.Error):
        return False


def validar_senha(senha: str) -> list[str]:
    erros: list[str] = []
    if len(senha) < 12:
        erros.append("A senha deve ter pelo menos 12 caracteres.")
    if not re.search(r"[a-z]", senha):
        erros.append("Inclua ao menos uma letra minúscula.")
    if not re.search(r"[A-Z]", senha):
        erros.append("Inclua ao menos uma letra maiúscula.")
    if not re.search(r"\d", senha):
        erros.append("Inclua ao menos um número.")
    if not re.search(r"[^A-Za-z0-9]", senha):
        erros.append("Inclua ao menos um caractere especial.")
    return erros


# ---------------------------------------------------------------------------
# Persistência: Supabase como prioridade absoluta + Excel local de segurança
# ---------------------------------------------------------------------------
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
    migrou_senha = False

    for coluna in COLUNAS[entidade]:
        if coluna not in df.columns:
            if coluna in {"dias"}:
                df[coluna] = 0
            elif coluna in {"ativo"}:
                df[coluna] = True
            elif coluna in {"status"}:
                df[coluna] = "Ativo"
            else:
                df[coluna] = ""

    if entidade == "colaboradores":
        for coluna in ("matricula", "funcionario", "setor", "cargo", "status", "decisao_experiencia"):
            df[coluna] = df[coluna].map(limpar_texto)
        df["matricula"] = df["matricula"].map(limpar_matricula)
        df["status"] = df["status"].replace({"Ferias": "Férias", "ferias": "Férias"}).replace("", "Ativo")
        for coluna in ("admissao", "nascimento", "ultimas_ferias", "data_retorno_ferias", "data_desligamento"):
            df[coluna] = df[coluna].map(data_iso)

    elif entidade == "usuarios":
        if "senha_legada" in df.columns:
            for indice, senha_legada in df["senha_legada"].items():
                if not limpar_texto(df.at[indice, "senha_hash"]) and limpar_texto(senha_legada):
                    df.at[indice, "senha_hash"] = hash_senha(limpar_texto(senha_legada))
                    migrou_senha = True
        for coluna in ("usuario", "nome", "email", "senha_hash", "perfil", "modulos", "telefone"):
            df[coluna] = df[coluna].map(limpar_texto)
        df["usuario"] = df["usuario"].str.lower()
        df["email"] = df["email"].str.lower()
        df["perfil"] = df["perfil"].replace("", "Gestor")
        df["ativo"] = df["ativo"].map(valor_bool)

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

    return df[COLUNAS[entidade]].copy(), migrou_senha


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


def converter_para_registros(df: pd.DataFrame) -> list[dict[str, Any]]:
    registros: list[dict[str, Any]] = []
    for registro in df.to_dict(orient="records"):
        limpo: dict[str, Any] = {}
        for chave, valor in registro.items():
            if pd.isna(valor) or limpar_texto(valor) == "":
                limpo[chave] = None
            elif isinstance(valor, (pd.Timestamp, datetime, date)):
                limpo[chave] = data_iso(valor)
            elif isinstance(valor, bool):
                limpo[chave] = valor
            else:
                limpo[chave] = valor
        registros.append(limpo)
    return registros


def carregar_entidade(entidade: str) -> tuple[pd.DataFrame, str]:
    # Prioridade absoluta: tenta carregar direto do Supabase
    cliente = obter_supabase()
    if cliente is not None:
        try:
            resposta = cliente.table(entidade).select("*").execute()
            dados = getattr(resposta, "data", None)
            if dados is not None and len(dados) > 0:
                df, migrou_senha = normalizar_entidade(pd.DataFrame(dados), entidade)
                if migrou_senha:
                    salvar_entidade(entidade, df, mostrar_feedback=False)
                return df, "Supabase"
        except Exception as erro:
            st.session_state["erro_supabase"] = f"Aviso: Usando base local temporariamente (Supabase instável: {erro})"

    # Fallback para o Excel local apenas se o Supabase falhar ou estiver vazio
    df, migrou_senha = normalizar_entidade(ler_excel(entidade), entidade)
    if migrou_senha:
        salvar_excel_atomico(df, entidade)
    return df, "Excel local (Fallback)"


def salvar_entidade(entidade: str, df: pd.DataFrame, mostrar_feedback: bool = True) -> bool:
    df_normalizado, _ = normalizar_entidade(df, entidade)
    
    # 1. Salva na nuvem (Supabase) primeiro para garantir a integridade da verdade principal
    cliente = obter_supabase()
    if cliente is not None:
        try:
            registros = converter_para_registros(df_normalizado)
            if registros:
                cliente.table(entidade).upsert(registros, on_conflict=CHAVES_PRIMARIAS[entidade]).execute()
        except Exception as erro:
            st.error(f"Erro crítico ao salvar no Supabase. Alteração cancelada para evitar perda de dados. Detalhe: {erro}")
            return False

    # 2. Salva o backup local em Excel apenas após o sucesso na nuvem
    try:
        salvar_excel_atomico(df_normalizado, entidade)
    except Exception as erro:
        st.warning(f"Salvo no Supabase, mas houve falha ao atualizar o arquivo local de segurança: {erro}")

    if mostrar_feedback:
        st.success("Dados salvos com segurança na nuvem e localmente.")
    return True


def registrar_historico(matricula: str, funcionario: str, tipo: str, descricao: str, autor: str) -> None:
    historico, _ = carregar_entidade("historico")
    novo = {
        "evento_id": str(uuid.uuid4()), "matricula": limpar_matricula(matricula),
        "funcionario": limpar_texto(funcionario), "data": datetime.now().isoformat(timespec="minutes"),
        "tipo_evento": limpar_texto(tipo), "descricao": limpar_texto(descricao), "autor": limpar_texto(autor),
    }
    salvar_entidade("historico", pd.concat([historico, pd.DataFrame([novo])], ignore_index=True), False)


# ---------------------------------------------------------------------------
# Sessão, autenticação e autorização
# ---------------------------------------------------------------------------
def iniciar_estado_sessao() -> None:
    padroes = {
        "autenticado": False, "usuario": "", "nome_usuario": "", "perfil": "",
        "modulos": [], "fonte_colaboradores": "", "tentativas_login": 0,
    }
    for chave, valor in padroes.items():
        st.session_state.setdefault(chave, valor)


def provisionar_admin_inicial(usuarios: pd.DataFrame) -> pd.DataFrame:
    if not usuarios.empty:
        return usuarios
    usuario = segredo("BOOTSTRAP_ADMIN_USER").lower().strip()
    senha = segredo("BOOTSTRAP_ADMIN_PASSWORD")
    email = segredo("BOOTSTRAP_ADMIN_EMAIL")
    if not usuario or not senha or not email:
        st.error("Nenhum usuário existe. Configure BOOTSTRAP_ADMIN_USER, BOOTSTRAP_ADMIN_EMAIL e BOOTSTRAP_ADMIN_PASSWORD em secrets.toml.")
        st.stop()
    erros = validar_senha(senha)
    if erros:
        st.error("A senha inicial não atende à política: " + " ".join(erros))
        st.stop()
    novo = pd.DataFrame([{
        "usuario": usuario, "nome": "Administrador Inicial", "email": email.lower(),
        "senha_hash": hash_senha(senha), "perfil": "Admin", "modulos": ",".join(TODOS_MODULOS),
        "telefone": "", "ativo": True,
    }])
    if salvar_entidade("usuarios", novo, False):
        st.success("Administrador inicial criado com sucesso.")
    return novo


def tela_login() -> bool:
    iniciar_estado_sessao()
    if st.session_state["autenticado"]:
        return True

    st.title("🔒 Acesso Restrito — Painel de Gestão & DP")
    st.caption("Use uma conta cadastrada.")
    usuarios, _ = carregar_entidade("usuarios")
    usuarios = provisionar_admin_inicial(usuarios)

    with st.form("form_login"):
        identificador = st.text_input("E-mail ou usuário").strip().lower()
        senha = st.text_input("Senha", type="password")
        enviar = st.form_submit_button("Entrar")

    if enviar:
        if st.session_state["tentativas_login"] >= 5:
            st.error("Muitas tentativas nesta sessão. Atualize a página antes de tentar novamente.")
            return False
        candidato = usuarios[(usuarios["usuario"] == identificador) | (usuarios["email"] == identificador)]
        if candidato.empty or not bool(candidato.iloc[0]["ativo"]) or not verificar_senha(senha, candidato.iloc[0]["senha_hash"]):
            st.session_state["tentativas_login"] += 1
            st.error("E-mail/usuário ou senha incorretos.")
            return False

        registro = candidato.iloc[0]
        modulos = [modulo.strip() for modulo in limpar_texto(registro["modulos"]).split(",") if modulo.strip()]
        st.session_state.update({
            "autenticado": True, "usuario": registro["usuario"], "nome_usuario": registro["nome"],
            "perfil": registro["perfil"], "modulos": list(TODOS_MODULOS) if registro["perfil"] == "Admin" else modulos,
            "tentativas_login": 0,
        })
        st.rerun()
    return False


def encerrar_sessao() -> None:
    for chave in ("autenticado", "usuario", "nome_usuario", "perfil", "modulos", "tentativas_login"):
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
            if posicao.empty or not verificar_senha(atual, usuarios.loc[posicao[0], "senha_hash"]):
                st.error("A senha atual está incorreta.")
            elif nova != confirmar:
                st.error("A confirmação não corresponde à nova senha.")
            else:
                erros = validar_senha(nova)
                if erros:
                    st.error(" ".join(erros))
                else:
                    usuarios.loc[posicao[0], "senha_hash"] = hash_senha(nova)
                    salvar_entidade("usuarios", usuarios)


# ---------------------------------------------------------------------------
# Utilitários de regras de negócio e relatórios
# ---------------------------------------------------------------------------
def filtrar_setor(df: pd.DataFrame, setor: str) -> pd.DataFrame:
    if setor == "Todos os setores" or df.empty:
        return df.copy()
    return df[df["setor"] == setor].copy()


def tabela_exibicao(df: pd.DataFrame, campos: list[str]) -> pd.DataFrame:
    rotulos = {
        "matricula": "Matrícula", "funcionario": "Funcionário", "setor": "Setor", "cargo": "Cargo",
        "nome": "Nome", "usuario": "Usuário", "email": "E-mail", "perfil": "Perfil", "telefone": "Telefone", "ativo": "Ativo",
        "admissao": "Admissão", "nascimento": "Nascimento", "status": "Status", "ultimas_ferias": "Últimas férias",
        "data_retorno_ferias": "Retorno das férias", "data": "Data", "tipo": "Tipo", "dias": "Dias",
        "cid": "CID", "motivo": "Motivo", "origem": "Origem", "epi": "EPI", "detalhe_tamanho": "Detalhe/Tamanho",
        "responsavel": "Responsável", "tipo_evento": "Tipo de evento", "descricao": "Descrição", "autor": "Autor",
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


# ---------------------------------------------------------------------------
# Módulos da aplicação
# ---------------------------------------------------------------------------
def tela_dashboard(colaboradores: pd.DataFrame, faltas: pd.DataFrame, setor: str) -> None:
    hoje = date.today()
    base = filtrar_setor(colaboradores, setor)
    faltas_base = filtrar_setor(faltas, setor)
    ativos = base[base["status"] == "Ativo"]
    em_ferias = base[base["status"] == "Férias"]
    afastados = base[base["status"] == "Afastado"]
    ocorrencias_hoje = faltas_base[faltas_base["data"].map(para_data) == hoje] if not faltas_base.empty else faltas_base
    ausencias_hoje = ocorrencias_hoje[ocorrencias_hoje["tipo"] != "Folga Concedida"] if not ocorrencias_hoje.empty else ocorrencias_hoje

    st.subheader("Painel geral de indicadores")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total no quadro", len(base))
    c2.metric("Ativos", len(ativos))
    c3.metric("Em férias", len(em_ferias))
    c4.metric("Afastados", len(afastados))
    c5.metric("Ausências hoje", len(ausencias_hoje))

    retornos = em_ferias.copy()
    retornos["retorno"] = retornos["data_retorno_ferias"].map(para_data)
    retornos = retornos[retornos["retorno"].notna()]
    if not retornos.empty:
        for _, pessoa in retornos.iterrows():
            dias = (pessoa["retorno"] - hoje).days
            if dias <= 2:
                mensagem = f"{pessoa['funcionario']} — retorno previsto em {formatar_data(pessoa['retorno'])}."
                if dias < 0:
                    st.error("Retorno de férias vencido: " + mensagem)
                elif dias == 0:
                    st.warning("Retorno de férias hoje: " + mensagem)
                else:
                    st.info(f"Retorno de férias em {dias} dia(s): " + mensagem)

    aniversariantes = base[base["nascimento"].map(lambda valor: (para_data(valor) or date(1900, 1, 1)).strftime("%m-%d") == hoje.strftime("%m-%d"))]
    if not aniversariantes.empty:
        st.success("Aniversariantes de hoje: " + ", ".join(aniversariantes["funcionario"].tolist()))


def tela_chamada(colaboradores: pd.DataFrame, faltas: pd.DataFrame, setor: str, autor: str) -> None:
    st.subheader("Chamada diária e ocorrências")
    aba_chamada, aba_avulso, aba_historico = st.tabs(["Chamada diária", "Lançamento avulso", "Histórico"])
    base = filtrar_setor(colaboradores, setor)
    ativos = base[base["status"] == "Ativo"].copy()
    termos_lideranca = r"gerente|supervisor|encarregado|coordenador|líder|lider"
    operacionais = ativos[~ativos["cargo"].str.lower().str.contains(termos_lideranca, na=False)].copy()

    with aba_chamada:
        data_chamada = st.date_input("Data da chamada", value=date.today(), key="data_chamada")
        if operacionais.empty:
            st.info("Não há colaboradores operacionais ativos para o filtro atual.")
        else:
            anteriores = faltas[(faltas["data"].map(para_data) == data_chamada) & (faltas["origem"] == "Chamada")]
            anteriores = filtrar_setor(anteriores, setor)
            estados: dict[str, str] = {linha["matricula"]: "Presente" for _, linha in operacionais.iterrows()}
            for _, ocorrencia in anteriores.iterrows():
                estados[ocorrencia["matricula"]] = "Folga" if ocorrencia["tipo"] == "Folga Concedida" else "Ausente"

            with st.form("form_chamada"):
                novo_estado: dict[str, str] = {}
                for _, pessoa in operacionais.iterrows():
                    col_nome, col_status = st.columns([2.2, 1.35])
                    col_nome.markdown(f"**{pessoa['funcionario']}**  \n`{pessoa['matricula']}`")
                    novo_estado[pessoa["matricula"]] = col_status.radio(
                        "Status", ("Presente", "Folga", "Ausente"),
                        index=("Presente", "Folga", "Ausente").index(estados.get(pessoa["matricula"], "Presente")),
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
                    if escolha != "Presente":
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
            with st.form("form_ocorrencia", clear_on_submit=True):
                matricula = st.selectbox("Colaborador", opcoes, format_func=lambda valor: mapa[valor])
                tipo = st.selectbox("Tipo", TIPOS_OCORRENCIA)
                data_ocorrencia = st.date_input("Data", value=date.today(), key="data_ocorrencia")
                dias = st.number_input("Quantidade de dias", min_value=1, max_value=365, value=1)
                cid = st.text_input("CID (opcional)").strip().upper()
                motivo = st.text_area("Observação", max_chars=500).strip()
                if st.form_submit_button("Salvar ocorrência"):
                    pessoa = base[base["matricula"] == matricula].iloc[0]
                    novo = {
                        "registro_id": str(uuid.uuid4()), "matricula": matricula, "funcionario": pessoa["funcionario"],
                        "setor": pessoa["setor"], "data": data_ocorrencia.isoformat(), "tipo": tipo, "dias": int(dias),
                        "cid": cid, "motivo": motivo, "origem": "Avulso",
                    }
                    if salvar_entidade("faltas", pd.concat([faltas, pd.DataFrame([novo])], ignore_index=True)):
                        registrar_historico(matricula, pessoa["funcionario"], "Ocorrência", f"{tipo} em {data_ocorrencia:%d/%m/%Y}.", autor)
                        st.rerun()

    with aba_historico:
        historico = filtrar_setor(faltas, setor).sort_values("data", ascending=False)
        tabela = tabela_exibicao(historico, ["data", "funcionario", "setor", "tipo", "dias", "cid", "motivo", "origem"])
        st.dataframe(tabela, use_container_width=True, hide_index=True)
        bloco_exportacao("historico_faltas", tabela)


def tela_epi(colaboradores: pd.DataFrame, epis: pd.DataFrame, setor: str, autor: str) -> None:
    st.subheader("Solicitação e entrega de EPI")
    base = filtrar_setor(colaboradores, setor)
    opcoes, mapa = opcoes_colaboradores(base, apenas_ativos=True)
    if not opcoes:
        st.info("Não há colaboradores ativos para o filtro atual.")
        return
    with st.form("form_epi", clear_on_submit=True):
        matricula = st.selectbox("Colaborador", opcoes, format_func=lambda valor: mapa[valor])
        c1, c2, c3 = st.columns(3)
        epi = c1.selectbox("EPI", ("Camiseta", "Bota de segurança", "Luvas", "Óculos", "Protetor auricular", "Outro"))
        detalhe = c2.text_input("Tamanho/Detalhe", max_chars=100)
        data_entrega = c3.date_input("Data da entrega", value=date.today())
        if st.form_submit_button("Registrar entrega"):
            pessoa = base[base["matricula"] == matricula].iloc[0]
            novo = {
                "entrega_id": str(uuid.uuid4()), "matricula": matricula, "funcionario": pessoa["funcionario"],
                "setor": pessoa["setor"], "data": data_entrega.isoformat(), "epi": epi,
                "detalhe_tamanho": detalhe, "responsavel": autor,
            }
            if salvar_entidade("epis", pd.concat([epis, pd.DataFrame([novo])], ignore_index=True)):
                registrar_historico(matricula, pessoa["funcionario"], "Entrega de EPI", f"{epi}: {detalhe or 'sem detalhe'}.", autor)
                st.rerun()

    st.markdown("#### Histórico de entregas")
    tabela = tabela_exibicao(filtrar_setor(epis, setor).sort_values("data", ascending=False), ["data", "funcionario", "setor", "epi", "detalhe_tamanho", "responsavel"])
    st.dataframe(tabela, use_container_width=True, hide_index=True)


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
    st.markdown(f"**Admissão:** {formatar_data(pessoa['admissao'])} &nbsp;&nbsp; **Nascimento:** {formatar_data(pessoa['nascimento'])} &nbsp;&nbsp; **Retorno de férias:** {formatar_data(pessoa['data_retorno_ferias'])}")
    eventos = historico[historico["matricula"] == matricula].sort_values("data", ascending=False)
    st.markdown("#### Linha do tempo")
    if eventos.empty:
        st.caption("Nenhum evento registrado para este colaborador.")
    else:
        st.dataframe(tabela_exibicao(eventos, ["data", "tipo_evento", "descricao", "autor"]), use_container_width=True, hide_index=True)


def tela_experiencia(colaboradores: pd.DataFrame, setor: str) -> None:
    st.subheader("Controle de experiência — 45 e 90 dias")
    hoje = date.today()
    base = filtrar_setor(colaboradores, setor)
    base = base[(base["status"] == "Ativo") & (base["decisao_experiencia"].str.lower() != "efetivado")].copy()
    linhas: list[dict[str, Any]] = []
    for _, pessoa in base.iterrows():
        classificacao, dias45, dias90 = classificar_experiencia(para_data(pessoa["admissao"]), hoje)
        linhas.append({
            "Matrícula": pessoa["matricula"], "Funcionário": pessoa["funcionario"], "Setor": pessoa["setor"], "Cargo": pessoa["cargo"],
            "45 dias": formatar_data((para_data(pessoa["admissao"]) + timedelta(days=45)) if para_data(pessoa["admissao"]) else None),
            "90 dias": formatar_data((para_data(pessoa["admissao"]) + timedelta(days=90)) if para_data(pessoa["admissao"]) else None),
            "Situação": classificacao, "Dias p/ 90": dias90 if dias90 is not None else "—",
        })
    if not linhas:
        st.success("Nenhum colaborador pendente de decisão de experiência.")
    else:
        tabela = pd.DataFrame(linhas).sort_values("Dias p/ 90", key=lambda s: pd.to_numeric(s, errors="coerce").fillna(99999))
        st.dataframe(tabela, use_container_width=True, hide_index=True)
        bloco_exportacao("controle_experiencia", tabela)


def tela_ferias(colaboradores: pd.DataFrame, autor: str) -> None:
    st.subheader("Escala de férias")
    opcoes, mapa = opcoes_colaboradores(colaboradores, apenas_ativos=True)
    with st.form("form_ferias", clear_on_submit=True):
        matricula = st.selectbox("Colaborador", opcoes, format_func=lambda valor: mapa[valor]) if opcoes else None
        c1, c2 = st.columns(2)
        inicio = c1.date_input("Início", value=date.today())
        retorno = c2.date_input("Retorno previsto", value=date.today() + timedelta(days=30))
        if st.form_submit_button("Registrar férias"):
            if not matricula:
                st.error("Selecione um colaborador.")
            elif retorno <= inicio:
                st.error("A data de retorno precisa ser posterior à data de início.")
            else:
                indice = colaboradores.index[colaboradores["matricula"] == matricula][0]
                colaboradores.loc[indice, "status"] = "Férias"
                colaboradores.loc[indice, "ultimas_ferias"] = inicio.isoformat()
                colaboradores.loc[indice, "data_retorno_ferias"] = retorno.isoformat()
                if salvar_entidade("colaboradores", colaboradores):
                    registrar_historico(matricula, colaboradores.loc[indice, "funcionario"], "Férias", f"Férias de {inicio:%d/%m/%Y} a {retorno:%d/%m/%Y}.", autor)
                    st.rerun()
    em_ferias = colaboradores[colaboradores["status"] == "Férias"]
    st.dataframe(tabela_exibicao(em_ferias, ["matricula", "funcionario", "setor", "cargo", "ultimas_ferias", "data_retorno_ferias"]), use_container_width=True, hide_index=True)


def tela_indicadores(colaboradores: pd.DataFrame, faltas: pd.DataFrame, setor: str) -> None:
    st.subheader("Indicadores de frequência e absenteísmo")
    c1, c2 = st.columns(2)
    inicio = c1.date_input("Início do período", value=date.today().replace(day=1))
    fim = c2.date_input("Fim do período", value=date.today())
    if fim < inicio:
        st.error("O fim do período precisa ser igual ou posterior ao início.")
        return
    base_faltas = filtrar_setor(faltas, setor).copy()
    base_faltas["data_dt"] = pd.to_datetime(base_faltas["data"], errors="coerce")
    periodo = base_faltas[(base_faltas["data_dt"] >= pd.Timestamp(inicio)) & (base_faltas["data_dt"] <= pd.Timestamp(fim))]
    ausencias = periodo[periodo["tipo"] != "Folga Concedida"]
    dias_ausentes = int(ausencias["dias"].sum()) if not ausencias.empty else 0
    quadro_medio = len(filtrar_setor(colaboradores, setor).query("status == 'Ativo'"))
    dias_calendario = (fim - inicio).days + 1
    taxa = (dias_ausentes / (quadro_medio * dias_calendario) * 100) if quadro_medio and dias_calendario else 0
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Dias de ausência", dias_ausentes)
    m2.metric("Atestados", int((periodo["tipo"] == "Atestado Médico").sum()))
    m3.metric("Faltas injustificadas", int((periodo["tipo"] == "Falta Injustificada").sum()))
    m4.metric("Taxa no período", f"{taxa:.2f}%")
    st.caption("Taxa = dias de ausência ÷ (colaboradores ativos × dias do período). Folgas concedidas não entram no numerador.")
    if not periodo.empty:
        resumo = periodo.groupby("tipo", as_index=False)["dias"].sum().sort_values("dias", ascending=False)
        st.bar_chart(resumo, x="tipo", y="dias", use_container_width=True)


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
            c5, c6, c7 = st.columns(3)
            admissao = c5.date_input("Admissão", value=date.today())
            nascimento = c6.date_input("Nascimento", value=date(1990, 1, 1))
            status = c7.selectbox("Status", STATUS_COLABORADOR)
            if st.form_submit_button("Cadastrar"):
                if not matricula or not nome or not setor:
                    st.error("Matrícula, nome e setor são obrigatórios.")
                elif (colaboradores["matricula"] == limpar_matricula(matricula)).any():
                    st.error("Já existe um colaborador com esta matrícula.")
                else:
                    novo = {
                        "matricula": limpar_matricula(matricula), "funcionario": nome, "setor": setor, "cargo": cargo,
                        "admissao": admissao.isoformat(), "nascimento": nascimento.isoformat(), "status": status,
                        "ultimas_ferias": "", "data_retorno_ferias": "", "decisao_experiencia": "", "data_desligamento": "",
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
            admissao = c5.date_input("Admissão", value=para_data(pessoa["admissao"]) or date.today())
            status = c6.selectbox("Status", STATUS_COLABORADOR, index=STATUS_COLABORADOR.index(pessoa["status"]) if pessoa["status"] in STATUS_COLABORADOR else 0)
            decisao = c7.selectbox("Experiência", ("", "Em avaliação", "Efetivado", "Não efetivado"), index=("", "Em avaliação", "Efetivado", "Não efetivado").index(pessoa["decisao_experiencia"]) if pessoa["decisao_experiencia"] in ("", "Em avaliação", "Efetivado", "Não efetivado") else 0)
            data_desligamento = st.date_input("Data do desligamento", value=para_data(pessoa["data_desligamento"]) or date.today(), disabled=status != "Desligado")
            if st.form_submit_button("Atualizar"):
                nova_matricula = limpar_matricula(nova_matricula)
                duplicada = (colaboradores["matricula"] == nova_matricula) & (colaboradores.index != indice)
                if not nova_matricula or not nome or not setor:
                    st.error("Matrícula, nome e setor são obrigatórios.")
                elif duplicada.any():
                    st.error("A matrícula informada já pertence a outro colaborador.")
                else:
                    colaboradores.loc[indice, ["matricula", "funcionario", "setor", "cargo", "admissao", "status", "decisao_experiencia", "data_desligamento"]] = [
                        nova_matricula, nome, setor, cargo, admissao.isoformat(), status, decisao,
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
    aba_novo, aba_lista = st.tabs(["Novo usuário", "Lista"])
    with aba_novo:
        with st.form("novo_usuario", clear_on_submit=True):
            c1, c2 = st.columns(2)
            nome = c1.text_input("Nome").strip()
            usuario = c2.text_input("Login", help="Será convertido para letras minúsculas.").strip().lower()
            c3, c4 = st.columns(2)
            email = c3.text_input("E-mail").strip().lower()
            senha = c4.text_input("Senha inicial", type="password")
            telefone = st.text_input("Telefone (opcional)").strip()
            perfil = st.selectbox("Perfil", ("Gestor", "Admin"))
            modulos = st.multiselect("Módulos liberados", TODOS_MODULOS, default=list(TODOS_MODULOS) if perfil == "Admin" else ["Dashboard & Alertas", "Chamada & Faltas do Dia"])
            if st.form_submit_button("Criar usuário"):
                if not nome or not usuario or not email:
                    st.error("Nome, login e e-mail são obrigatórios.")
                elif (usuarios["usuario"] == usuario).any() or (usuarios["email"] == email).any():
                    st.error("Já existe um usuário com este login ou e-mail.")
                else:
                    erros = validar_senha(senha)
                    if erros:
                        st.error(" ".join(erros))
                    else:
                        novo = {
                            "usuario": usuario, "nome": nome, "email": email, "senha_hash": hash_senha(senha),
                            "perfil": perfil, "modulos": ",".join(TODOS_MODULOS if perfil == "Admin" else modulos),
                            "telefone": telefone, "ativo": True,
                        }
                        if salvar_entidade("usuarios", pd.concat([usuarios, pd.DataFrame([novo])], ignore_index=True)):
                            st.info("Usuário criado. Compartilhe apenas o login e instrua a pessoa a definir uma senha segura por um canal protegido.")
                            st.rerun()
    with aba_lista:
        st.dataframe(tabela_exibicao(usuarios, ["nome", "usuario", "email", "perfil", "telefone", "ativo"]), use_container_width=True, hide_index=True)


def tela_importacao(colaboradores: pd.DataFrame) -> None:
    if not exigir_admin():
        return
    st.subheader("Importar nova base de colaboradores")
    st.warning("A importação valida a planilha antes de salvar. Nenhum registro existente de outros setores é excluído; matrículas existentes são atualizadas e novos registros são adicionados.")
    arquivo = st.file_uploader("Planilha .xlsx", type=["xlsx"])
    if not arquivo:
        return
    try:
        bruto = pd.read_excel(arquivo, dtype=object)
        importado, _ = normalizar_entidade(bruto, "colaboradores")
    except Exception as erro:
        st.error(f"Não foi possível processar a planilha: {erro}")
        return
    erros: list[str] = []
    obrigatorios = {"matricula": "Matrícula", "funcionario": "Funcionário", "setor": "Setor", "admissao": "Admissão"}
    for campo, rotulo in obrigatorios.items():
        if (importado[campo].map(limpar_texto) == "").any():
            erros.append(f"Existem linhas sem {rotulo}.")
    if importado["matricula"].duplicated().any():
        erros.append("A planilha possui matrículas duplicadas.")
    if importado["admissao"].isna().any() or (importado["admissao"] == "").any():
        erros.append("Existem datas de admissão inválidas ou ausentes.")
    if erros:
        for erro in erros:
            st.error(erro)
        return
    st.markdown("#### Prévia validada")
    st.dataframe(tabela_exibicao(importado.head(20), ["matricula", "funcionario", "setor", "cargo", "admissao", "status"]), use_container_width=True, hide_index=True)
    confirmar = st.checkbox("Confirmo que revisei a prévia e desejo aplicar as atualizações de forma incremental.")
    if st.button("Importar e atualizar base", disabled=not confirmar):
        combinado = colaboradores.set_index("matricula")
        atualizacoes = importado.set_index("matricula")
        combinado.update(atualizacoes)
        novos = atualizacoes.loc[~atualizacoes.index.isin(combinado.index)]
        resultado = pd.concat([combinado, novos]).reset_index()
        if salvar_entidade("colaboradores", resultado):
            st.success(f"Importação concluída com sucesso: {len(importado)} linhas processadas e enviadas ao Supabase.")
            st.rerun()


def tela_assistente_ia() -> None:
    st.subheader("Assistente IA para DP e Gestão")
    chave = segredo("GEMINI_API_KEY")
    if not chave:
        st.info("Assistente desativado. Configure GEMINI_API_KEY em secrets.toml para habilitá-lo.")
        return
    try:
        import google.generativeai as genai
        genai.configure(api_key=chave)
        modelo = genai.GenerativeModel("gemini-1.5-pro")
    except Exception as erro:
        st.error(f"Não foi possível iniciar o assistente: {erro}")
        return
    st.caption("Não inclua CPF, CID, senhas, dados bancários ou outras informações pessoais sensíveis na conversa.")
    historico = st.session_state.setdefault("historico_ia", [])
    for mensagem in historico:
        with st.chat_message(mensagem["role"]):
            st.markdown(mensagem["content"])
    pergunta = st.chat_input("Digite uma pergunta sobre DP e gestão")
    if pergunta:
        historico.append({"role": "user", "content": pergunta})
        with st.chat_message("user"):
            st.markdown(pergunta)
        contexto = "\n".join(f"{m['role']}: {m['content']}" for m in historico[-8:])
        instrucao = "Você é um assistente de DP e gestão. Responda em português, de forma objetiva, sem solicitar dados pessoais sensíveis.\n"
        with st.chat_message("assistant"):
            with st.spinner("Analisando..."):
                try:
                    resposta = modelo.generate_content(instrucao + contexto)
                    texto = getattr(resposta, "text", "Não foi possível gerar uma resposta para esta solicitação.")
                    st.markdown(texto)
                    historico.append({"role": "assistant", "content": texto})
                except Exception as erro:
                    st.error(f"Erro ao consultar o assistente: {erro}")


def aplicar_estilo() -> None:
    st.markdown("""
    <style>
      .stApp { background: #0E1117; color: #F8FAFC; }
      [data-testid="stSidebar"] { background: #163A2A; }
      [data-testid="stSidebar"] * { color: #F8FAFC !important; }
      div.stButton > button, div[data-testid="stFormSubmitButton"] > button {
        background: #F97316; color: white; border: 0; border-radius: 8px; font-weight: 700;
      }
      div.stButton > button:hover, div[data-testid="stFormSubmitButton"] > button:hover { background: #EA580C; color: white; }
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
    nome = st.session_state["nome_usuario"]
    perfil = st.session_state["perfil"]
    modulos = st.session_state["modulos"] or ["Dashboard & Alertas"]
    modulos = [modulo for modulo in modulos if modulo not in MODULOS_ADMIN or perfil == "Admin"]

    st.sidebar.title("🍊 Gestão & DP")
    st.sidebar.caption(f"{nome} · {perfil}")
    st.sidebar.caption(f"Fonte atual: **{fonte}**")
    if st.session_state.get("erro_supabase"):
        st.sidebar.warning(st.session_state.get("erro_supabase"))
    if st.sidebar.button("Sair"):
        encerrar_sessao()
    alterar_minha_senha()

    setores = ["Todos os setores"] + sorted([valor for valor in colaboradores["setor"].dropna().unique() if limpar_texto(valor)])
    setor = st.sidebar.selectbox("Filtrar por setor", setores)
    menu = st.sidebar.radio("Navegação", modulos)

    st.title("Painel de Gestão & DP")
    st.caption("Versão corrigida: prioridade absoluta ao Supabase, salvamento seguro e importação incremental por setor.")

    if menu == "Dashboard & Alertas":
        tela_dashboard(colaboradores, faltas, setor)
    elif menu == "Assistente IA (DP & Gestão)":
        tela_assistente_ia()
    elif menu == "Chamada & Faltas do Dia":
        tela_chamada(colaboradores, faltas, setor, nome)
    elif menu == "Solicitação & Entrega de EPI":
        tela_epi(colaboradores, epis, setor, nome)
    elif menu == "Ficha Individual do Colaborador":
        tela_ficha(colaboradores, historico)
    elif menu == "Controle de Experiência (45/90 dias)":
        tela_experiencia(colaboradores, setor)
    elif menu == "Escala de Férias":
        tela_ferias(colaboradores, nome)
    elif menu == "Colaboradores em Férias":
        ferias = filtrar_setor(colaboradores[colaboradores["status"] == "Férias"], setor)
        st.subheader("Colaboradores em férias")
        st.dataframe(tabela_exibicao(ferias, ["matricula", "funcionario", "setor", "cargo", "ultimas_ferias", "data_retorno_ferias"]), use_container_width=True, hide_index=True)
    elif menu == "Indicadores de Frequência & Absenteísmo":
        tela_indicadores(colaboradores, faltas, setor)
    elif menu == "Aniversariantes do Mês":
        st.subheader("Aniversariantes do mês")
        mes = st.selectbox("Mês", range(1, 13), index=date.today().month - 1, format_func=lambda numero: date(2000, numero, 1).strftime("%B").capitalize())
        aniversariantes = filtrar_setor(colaboradores, setor)
        aniversariantes = aniversariantes[aniversariantes["nascimento"].map(lambda valor: (para_data(valor) or date(1900, 1, 1)).month == mes)]
        aniversariantes = aniversariantes.sort_values("nascimento")
        st.dataframe(tabela_exibicao(aniversariantes, ["nascimento", "funcionario", "setor", "cargo"]), use_container_width=True, hide_index=True)
    elif menu == "Cadastrar / Editar Colaborador":
        tela_colaboradores(colaboradores, nome)
    elif menu == "Criar / Gerenciar Usuários":
        tela_usuarios(usuarios)
    elif menu == "Importar Nova Base":
        tela_importacao(colaboradores)


if __name__ == "__main__":
    main()
