import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
import streamlit as st


def ajustar_para_domingo(data):
  """Garante que a data inicial seja um domingo (weekday == 6)."""
  dias = (6 - data.weekday()) % 7
  return data + datetime.timedelta(days=dias)


def calcular_proxima_elegibilidade(dt_adm, dt_ult_ferias, limite_minimo_inicio):
  """Calcula a próxima data de elegibilidade dentro do ciclo atual/futuro."""
  if pd.isnull(dt_adm):
    return None

  data_base = dt_ult_ferias if pd.notnull(dt_ult_ferias) else dt_adm

  # Avança o ciclo ano a ano até que a data de vencimento atinja ou supere o limite mínimo permitido
  data_elegivel = data_base + relativedelta(years=1)
  while data_elegivel < limite_minimo_inicio:
    data_elegivel += relativedelta(years=1)

  return data_elegivel


def alocar_ferias(
    df_input,
    cota_geral_padrao,
    cota_separacao,
    cota_outros_setores,
    flexibilizar_dez_jan,
    cota_alta_temporada,
    dias_gap_rh=30,
):
  df_sorted = df_input.copy()
  hoje = datetime.date.today()

  # GAP MÍNIMO DO RH: O início das férias precisa ser de no mínimo X dias a partir de hoje
  limite_minimo_inicio = hoje + datetime.timedelta(days=dias_gap_rh)

  # Mapeamento de colunas
  col_nome = (
      "Funcionário"
      if "Funcionário" in df_sorted.columns
      else next(
          (
              c
              for c in df_sorted.columns
              if "func" in str(c).lower() or "nome" in str(c).lower()
          ),
          None,
      )
  )
  col_setor = (
      "Setor"
      if "Setor" in df_sorted.columns
      else next(
          (c for c in df_sorted.columns if "setor" in str(c).lower()), None
      )
  )

  if "dt_adm" in df_sorted.columns:
    col_admissao = "dt_adm"
  elif "Admissão" in df_sorted.columns:
    col_admissao = "Admissão"
  else:
    col_admissao = next(
        (c for c in df_sorted.columns if "admiss" in str(c).lower()), None
    )

  col_ult_ferias = (
      "dt_ult_ferias"
      if "dt_ult_ferias" in df_sorted.columns
      else ("Ultimas_Ferias" if "Ultimas_Ferias" in df_sorted.columns else None)
  )

  if not (col_nome and col_setor and col_admissao):
    return (
        None,
        "As colunas necessárias ('Funcionário', 'Setor' e 'Admissão') não foram"
        " identificadas.",
    )

  # Tratamento das datas
  df_sorted[col_admissao] = pd.to_datetime(
      df_sorted[col_admissao], errors="coerce"
  ).dt.date
  df_sorted = df_sorted.dropna(subset=[col_admissao])

  if col_ult_ferias:
    df_sorted[col_ult_ferias] = pd.to_datetime(
        df_sorted[col_ult_ferias], dayfirst=True, errors="coerce"
    ).dt.date

  # Calcula a elegibilidade respeitando o GAP do RH
  df_sorted["data_elegivel"] = df_sorted.apply(
      lambda r: calcular_proxima_elegibilidade(
          r[col_admissao],
          r[col_ult_ferias] if col_ult_ferias else None,
          limite_minimo_inicio,
      ),
      axis=1,
  )

  df_sorted = df_sorted.dropna(subset=["data_elegivel"])
  df_sorted = df_sorted.sort_values(by="data_elegivel")

  ocupacao = {}
  resultado = []

  for _, row in df_sorted.iterrows():
    data_base = row["data_elegivel"]

    # Início da alocação respeita a elegibilidade e o limite do Gap RH
    inicio_busca = max(data_base, limite_minimo_inicio)
    ano = inicio_busca.year
    mes = inicio_busca.month

    alocado = False
    tentativas = 0

    while not alocado and tentativas < 24:
      chave_mes = (ano, mes)
      if chave_mes not in ocupacao:
        ocupacao[chave_mes] = {"total": 0, "setores": {}}

      dados_mes = ocupacao[chave_mes]

      eh_alta_temporada = flexibilizar_dez_jan and (mes == 12 or mes == 1)
      limite_geral = (
          cota_alta_temporada if eh_alta_temporada else cota_geral_padrao
      )

      setor = str(row[col_setor]).strip()
      limite_setor = (
          cota_separacao
          if setor.lower() in ["separação", "separacao"]
          else cota_outros_setores
      )
      if eh_alta_temporada:
        limite_setor = limite_geral

      qtd_setor = dados_mes["setores"].get(setor, 0)

      if dados_mes["total"] < limite_geral and qtd_setor < limite_setor:
        dados_mes["total"] += 1
        dados_mes["setores"][setor] = qtd_setor + 1

        primeiro_dia_mes = datetime.date(ano, mes, 1)

        if primeiro_dia_mes < inicio_busca:
          data_inicio_sugerida = ajustar_para_domingo(inicio_busca)
        else:
          data_inicio_sugerida = ajustar_para_domingo(primeiro_dia_mes)

        # Se após o ajuste do domingo cair antes do Gap RH, avança para o próximo domingo
        if data_inicio_sugerida < limite_minimo_inicio:
          data_inicio_sugerida += datetime.timedelta(days=7)

        fim_ferias = data_inicio_sugerida + datetime.timedelta(days=30)
        aviso_previo_limite = data_inicio_sugerida - datetime.timedelta(days=30)

        resultado.append({
            "Colaborador": row[col_nome],
            "Setor": setor,
            "Admissão": row[col_admissao].strftime("%d/%m/%Y"),
            "Início Férias (Domingo)": data_inicio_sugerida.strftime("%d/%m/%Y"),
            "Retorno": fim_ferias.strftime("%d/%m/%Y"),
            "Prazo Máx. Aviso RH": aviso_previo_limite.strftime("%d/%m/%Y"),
            "Status Cota Setor": f"{dados_mes['setores'][setor]}/{limite_setor}",
            "Status Cota Mês": f"{dados_mes['total']}/{limite_geral}",
        })
        alocado = True
      else:
        mes += 1
        if mes > 12:
          mes = 1
          ano += 1
        tentativas += 1

  return pd.DataFrame(resultado), None


def renderizar_modulo_ferias(df_base):
  st.title("📅 Planejamento Inteligente de Férias (Ciclo Atual)")
  st.caption(
      "Regras: Início no Domingo (Escala Dom-Qui) | Cota Máxima: 2/mês | Gap RH:"
      " Min. 30 dias | Trava por Setor"
  )

  if df_base is None or df_base.empty:
    st.warning("⚠️ Nenhuma base de dados carregada para simulação.")
    return

  st.sidebar.header("⚙️ Cotas & Prazos")
  gap_rh = st.sidebar.number_input(
      "Gap Mínimo RH (Dias)",
      min_value=15,
      max_value=60,
      value=30,
      step=5,
      key="f_gap",
  )
  cota_geral = st.sidebar.number_input(
      "Cota Geral (Mês)", min_value=1, value=2, key="f_cg"
  )
  cota_sep = st.sidebar.number_input(
      "Máx. Separação", min_value=1, value=2, key="f_cs"
  )
  cota_outros = st.sidebar.number_input(
      "Máx. Outros Setores", min_value=1, value=1, key="f_co"
  )

  st.sidebar.subheader("🏖️ Alta Temporada")
  flexib = st.sidebar.checkbox(
      "Expandir Dez/Jan", value=True, key="f_flex"
  )
  cota_alta = st.sidebar.number_input(
      "Cota Dez/Jan", min_value=2, value=5, key="f_ca"
  )

  if st.button("🔄 Simular Escala de Férias", type="primary"):
    df_res, erro = alocar_ferias(
        df_base, cota_geral, cota_sep, cota_outros, flexib, cota_alta, gap_rh
    )

    if erro:
      st.error(erro)
    else:
      st.session_state["df_res_ferias"] = df_res

  if "df_res_ferias" in st.session_state:
    df_res = st.session_state["df_res_ferias"]

    setores = ["Todos"] + list(df_res["Setor"].unique())
    setor_sel = st.selectbox("🔍 Filtrar visualização por Setor:", setores)

    if setor_sel != "Todos":
      st.dataframe(
          df_res[df_res["Setor"] == setor_sel], use_container_width=True
      )
    else:
      st.dataframe(df_res, use_container_width=True)

    st.subheader("📊 Ocupação das Cotas por Mês")
    resumo_mes = (
        df_res.groupby(["Status Cota Mês", "Setor"])
        .size()
        .unstack(fill_value=0)
    )
    st.bar_chart(resumo_mes)
