import streamlit as st
import pandas as pd
import plotly.express as px
import ferias
from datetime import datetime, date, timedelta
import os
import io

st.set_page_config(page_title="Gestão de Equipe Tropical", page_icon="👥", layout="wide")

# --- SISTEMA DE AUTENTICAÇÃO POR SENHA SEGURA ---
def verificar_senha():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if not st.session_state["autenticado"]:
        st.title("🔒 Acesso Restrito - Gestão de Equipe Tropical")
        st.info("Por razões de segurança, informe a senha de acesso para continuar.")
        
        senha_digitada = st.text_input("Digite a Senha de Acesso:", type="password")
        btn_entrar = st.button("🔑 Entrar no Sistema")
        
        # Tenta buscar dos secrets (nuvem), se não existir usa a senha padrão local '030711'
        try:
            senha_correta = st.secrets["SENHA_ACESSO"]
        except Exception:
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
    ARQUIVO_FALTAS = "faltas.xlsx"

    def carregar_dados():
        if os.path.exists(ARQUIVO_DADOS):
            df = pd.read_excel(ARQUIVO_DADOS)
            df['dt_adm'] = pd.to_datetime(df['Admissão'], dayfirst=True, errors='coerce').dt.date
            df['dt_nasc'] = pd.to_datetime(df['Nascimento'], dayfirst=True, errors='coerce').dt.date
            
            df['Vaga'] = df['Vaga'].astype(str).str.replace('.0', '', regex=False)
            df['Matricula'] = df['Matricula'].astype(str).str.replace('.0', '', regex=False)
            
            if 'Ultimas_Ferias' not in df.columns:
                df['Ultimas_Ferias'] = None
            else:
                df['dt_ult_ferias'] = pd.to_datetime(df['Ultimas_Ferias'], dayfirst=True, errors='coerce').dt.date
                
            if 'Decisao_Experiencia' not in df.columns:
                df['Decisao_Experiencia'] = None
                
            return df
        else:
            st.error(f"Arquivo '{ARQUIVO_DADOS}' não encontrado na pasta atual!")
            return pd.DataFrame()

    def carregar_faltas():
        if os.path.exists(ARQUIVO_FALTAS):
            df_f = pd.read_excel(ARQUIVO_FALTAS)
            df_f['dt_falta'] = pd.to_datetime(df_f['Data'], dayfirst=True, errors='coerce').dt.date
            return df_f
        else:
            return pd.DataFrame(columns=["Matricula", "Funcionário", "Setor", "Data", "Tipo", "Dias", "CID", "Motivo", "dt_falta"])

    def salvar_dados(df):
        cols_salvar = [c for c in df.columns if c not in ['dt_adm', 'dt_nasc', 'dt_ult_ferias', 'exp_45', 'exp_90', 'dias_para_45', 'dias_para_90']]
        df[cols_salvar].to_excel(ARQUIVO_DADOS, index=False)

    def salvar_faltas(df_f):
        cols_salvar = [c for c in df_f.columns if c != 'dt_falta']
        df_f[cols_salvar].to_excel(ARQUIVO_FALTAS, index=False)

    def converter_df_para_excel(df_exp):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_exp.to_excel(writer, index=False, sheet_name='Relatorio')
        return output.getvalue()

    @st.dialog("📋 Lista Detalhada do Quadro")
    def exibir_modal_detalhes(titulo, df_detalhes):
        st.subheader(titulo)
        if df_detalhes.empty:
            st.info("Nenhum colaborador nesta situação.")
        else:
            st.dataframe(df_detalhes, use_container_width=True)

    df = carregar_dados()
    df_faltas = carregar_faltas()
    hoje = date.today()

    if st.sidebar.button("🚪 Sair / Desconectar"):
        st.session_state["autenticado"] = False
        st.rerun()

    st.title("👥 Gestão de Equipe Tropical")

    if not df.empty:
        aniversariantes_hoje = df[df['dt_nasc'].apply(lambda d: d.month == hoje.month and d.day == hoje.day if pd.notnull(d) else False)]
        if not aniversariantes_hoje.empty:
            for _, colab in aniversariantes_hoje.iterrows():
                st.balloons()
                st.success(f"🎉 **HOJE É ANIVERSÁRIO DE:** {colab['Funcionário']} ({colab['Cargo']} - Setor: {colab['Setor']})! Parabéns!")

        # --- NAVEGAÇÃO E FILTRO DE SETOR ---
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
            "Gestão de Férias",
            "Escala Inteligente de Férias",
            "Faltas & Folgas",
            "Aniversariantes do Mês", 
            "Cadastrar / Editar Colaborador",
            "📥 Importar Nova Base"
        ])

        # Lógica de Experiência
        df_exp = df_filtrado.copy()
        df_exp['exp_45'] = df_exp['dt_adm'].apply(lambda d: d + timedelta(days=45) if pd.notnull(d) else None)
        df_exp['exp_90'] = df_exp['dt_adm'].apply(lambda d: d + timedelta(days=90) if pd.notnull(d) else None)
        df_exp['dias_para_45'] = df_exp['exp_45'].apply(lambda d: (d - hoje).days if pd.notnull(d) else 999)
        df_exp['dias_para_90'] = df_exp['exp_90'].apply(lambda d: (d - hoje).days if pd.notnull(d) else 999)

        df_apenas_exp = df_exp[(df_exp['Status'] == 'Ativo') & (df_exp['dias_para_90'] >= 0) & (df_exp['Decisao_Experiencia'] != 'Efetivado')].copy()

        if menu == "Dashboard & Alertas":
            st.subheader(f"⚠️ Painel Geral de Alertas - {setor_selecionado}")
            
            vagas_abertas = df_filtrado[df_filtrado['Status'] == 'Desligado (Quebra Experiencia)']
            if not vagas_abertas.empty:
                st.error(f"🚨 **ALERTA DE REPOSIÇÃO DE QUADRO:** Existem {len(vagas_abertas)} vaga(s) aberta(s) por quebra de contrato de experiência!")
                cols_vaga = st.columns(min(len(vagas_abertas), 3))
                for i, (_, v) in enumerate(vagas_abertas.iterrows()):
                    with cols_vaga[i % 3]:
                        st.info(f"📌 **VAGA ABERTA:** {v['Cargo']}\n\n**Setor:** {v['Setor']}\n\n**Ex-colaborador:** {v['Funcionário']}")

            c1, c2, c3, c4, c5, c6 = st.columns(6)
            
            c1.metric("Total Quadro", len(df_filtrado))
            if c1.button("🔍 Ver Quadro", key="btn_quadro"):
                df_det = df_filtrado[['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Status', 'Admissão']].copy()
                exibir_modal_detalhes("Quadro Geral de Colaboradores", df_det)
                
            df_ativos = df_filtrado[df_filtrado['Status'] == 'Ativo']
            c2.metric("Ativos", len(df_ativos))
            if c2.button("🔍 Ver Ativos", key="btn_ativos"):
                df_det = df_ativos[['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Admissão']].copy()
                exibir_modal_detalhes("Colaboradores Ativos no Quadro", df_det)

            df_ferias_st = df_filtrado[df_filtrado['Status'] == 'Férias']
            c3.metric("Em Férias", len(df_ferias_st))
            if c3.button("🔍 Ver Férias", key="btn_ferias"):
                df_det = df_ferias_st[['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Ultimas_Ferias']].copy()
                exibir_modal_detalhes("Colaboradores em Gozo de Férias", df_det)

            df_afastados = df_filtrado[df_filtrado['Status'].isin(['Atestado', 'Afastado', 'INSS'])]
            c4.metric("Atest./Afast./INSS", len(df_afastados))
            if c4.button("🔍 Ver Afastados", key="btn_afastados"):
                df_det = df_afastados[['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Status', 'Contato']].copy()
                exibir_modal_detalhes("Colaboradores Afastados / Atestado / INSS", df_det)

            faltas_mes = df_faltas_filtrado[df_faltas_filtrado['dt_falta'].apply(lambda d: d.month == hoje.month and d.year == hoje.year if pd.notnull(d) else False)] if not df_faltas_filtrado.empty else pd.DataFrame()
            c5.metric("Ocorrências (Mês)", len(faltas_mes))
            if c5.button("🔍 Ver Faltas/Atest", key="btn_faltas_m"):
                df_det = faltas_mes[['Data', 'Funcionário', 'Setor', 'Tipo', 'Dias', 'CID', 'Motivo']].copy() if not faltas_mes.empty else pd.DataFrame()
                exibir_modal_detalhes(f"Ocorrências e Lançamentos de {hoje.strftime('%m/%Y')}", df_det)

            niver_mes = df_filtrado[df_filtrado['dt_nasc'].apply(lambda d: d.month == hoje.month and d.year == hoje.year if pd.notnull(d) else False)]
            c6.metric("Aniversariantes", len(niver_mes))
            if c6.button("🔍 Ver Aniversár.", key="btn_niver_m"):
                df_det = niver_mes[['Nascimento', 'Funcionário', 'Setor', 'Cargo']].copy()
                exibir_modal_detalhes(f"Aniversariantes do Mês ({hoje.strftime('%m/%Y')})", df_det)

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

            st.markdown("---")
            col_exp, col_ferias = st.columns(2)

            with col_exp:
                st.subheader("🔔 Decisão de Experiência (Próximos 10 dias)")
                alerta_45 = df_apenas_exp[(df_apenas_exp['dias_para_45'] >= 0) & (df_apenas_exp['dias_para_45'] <= 10)]
                alerta_90 = df_apenas_exp[(df_apenas_exp['dias_para_90'] >= 0) & (df_apenas_exp['dias_para_90'] <= 10)]

                if alerta_45.empty and alerta_90.empty:
                    st.info("Nenhum contrato de experiência vencendo nos próximos 10 dias.")
                else:
                    for idx, r in alerta_45.iterrows():
                        st.warning(f"⏳ **45 DIAS:** {r['Funcionário']} ({r['Cargo']} - {r['Setor']})\nVence em {r['exp_45'].strftime('%d/%m/%Y')} (Faltam {r['dias_para_45']} dias)")
                        b1, b2 = st.columns(2)
                        if b1.button("✅ Efetivar (45d)", key=f"ef_45_{idx}"):
                            df.at[idx, 'Decisao_Experiencia'] = 'Efetivado'
                            salvar_dados(df)
                            st.success(f"{r['Funcionário']} efetivado com sucesso!")
                            st.rerun()
                        if b2.button("🚫 Quebra de Contrato", key=f"qb_45_{idx}"):
                            df.at[idx, 'Status'] = 'Desligado (Quebra Experiencia)'
                            df.at[idx, 'Decisao_Experiencia'] = 'Quebrado'
                            salvar_dados(df)
                            st.error(f"Contrato encerrado. Vaga liberada no quadro!")
                            st.rerun()

                    for idx, r in alerta_90.iterrows():
                        st.error(f"🚨 **90 DIAS (FINAL):** {r['Funcionário']} ({r['Cargo']} - {r['Setor']})\nVence em {r['exp_90'].strftime('%d/%m/%Y')} (Faltam {r['dias_para_90']} dias)")
                        b1, b2 = st.columns(2)
                        if b1.button("✅ Efetivar Definitivo", key=f"ef_90_{idx}"):
                            df.at[idx, 'Decisao_Experiencia'] = 'Efetivado'
                            salvar_dados(df)
                            st.success(f"{r['Funcionário']} efetivado definitivamente!")
                            st.rerun()
                        if b2.button("🚫 Quebra de Contrato", key=f"qb_90_{idx}"):
                            df.at[idx, 'Status'] = 'Desligado (Quebra Experiencia)'
                            df.at[idx, 'Decisao_Experiencia'] = 'Quebrado'
                            salvar_dados(df)
                            st.error(f"Contrato encerrado. Vaga liberada no quadro!")
                            st.rerun()

            with col_ferias:
                st.subheader("🏖️ Alertas de Férias Pendentes")
                alertas_ferias = []
                for idx, r in df_filtrado[df_filtrado['Status'] == 'Ativo'].iterrows():
                    adm = r['dt_adm']
                    ult_ferias = r.get('dt_ult_ferias') if 'dt_ult_ferias' in r else None
                    data_base = ult_ferias if pd.notnull(ult_ferias) else adm
                    
                    if pd.notnull(data_base):
                        anos = (hoje - data_base).days // 365
                        if anos >= 1:
                            inicio_aquisitivo = data_base + timedelta(days=365 * (anos - 1))
                            fim_aquisitivo = inicio_aquisitivo + timedelta(days=365)
                            limite_concessivo = fim_aquisitivo + timedelta(days=365)
                            dias_restantes = (limite_concessivo - hoje).days
                            if dias_restantes <= 60:
                                alertas_ferias.append((idx, r['Funcionário'], r['Setor'], limite_concessivo, dias_restantes))
                
                if not alertas_ferias:
                    st.info("Nenhum colaborador com risco imediato de férias vencidas.")
                else:
                    for idx, nome, setor, lim, dias in alertas_ferias:
                        c_info, c_btn = st.columns([3, 1])
                        with c_info:
                            st.error(f"⚠️ **{nome}** ({setor})\nLimite: {lim.strftime('%d/%m/%Y')} (Faltam {dias} dias)")
                        with c_btn:
                            if st.button("✅ Dar Baixa", key=f"baixa_{idx}"):
                                df.at[idx, 'Ultimas_Ferias'] = hoje.strftime('%d/%m/%Y')
                                salvar_dados(df)
                                st.success(f"Férias baixadas para {nome}!")
                                st.rerun()

        elif menu == "Controle de Experiência (45/90 dias)":
            st.subheader(f"📋 Colaboradores Atualmente em Período de Experiência - {setor_selecionado}")
            if df_apenas_exp.empty:
                st.success("Nenhum colaborador em período de experiência neste setor.")
            else:
                df_exibir = df_apenas_exp[['Matricula', 'Funcionário', 'Setor', 'Cargo', 'Admissão']].copy()
                df_exibir['Vencimento 45 Dias'] = df_apenas_exp['exp_45'].apply(lambda d: d.strftime('%d/%m/%Y') if pd.notnull(d) else "")
                df_exibir['Faltam (45d)'] = df_apenas_exp['dias_para_45'].apply(lambda d: f"{d} dias" if d >= 0 else "Já passou")
                df_exibir['Vencimento 90 Dias'] = df_apenas_exp['exp_90'].apply(lambda d: d.strftime('%d/%m/%Y') if pd.notnull(d) else "")
                df_exibir['Faltam (90d)'] = df_apenas_exp['dias_para_90'].apply(lambda d: f"{d} dias")
                st.dataframe(df_exibir, use_container_width=True)
                
                st.download_button(
                    label="📥 Baixar Tabela em Excel",
                    data=converter_df_para_excel(df_exibir),
                    file_name=f"experiencia_{setor_selecionado.lower().replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        elif menu == "Gestão de Férias":
            st.subheader(f"🏖️ Controle de Períodos Aquisitivos e Concessivos - {setor_selecionado}")
            lista_ferias = []
            for idx, r in df_filtrado[df_filtrado['Status'].isin(['Ativo', 'Férias'])].iterrows():
                adm = r['dt_adm']
                ult_ferias = pd.to_datetime(r.get('Ultimas_Ferias'), dayfirst=True, errors='coerce').date() if pd.notnull(r.get('Ultimas_Ferias')) else None
                data_base = ult_ferias if ult_ferias else adm
                
                if pd.notnull(data_base):
                    anos = (hoje - data_base).days // 365
                    inicio_aq = data_base + timedelta(days=365 * max(0, anos))
                    fim_aq = inicio_aq + timedelta(days=365)
                    limite_conc = fim_aq + timedelta(days=365)
                    dias_limite = (limite_conc - hoje).days
                    
                    status_ferias = "✅ Regular"
                    if r['Status'] == 'Férias':
                        status_ferias = "🏖️ Em Gozo de Férias"
                    elif dias_limite <= 60:
                        status_ferias = "⚠️ Atenção (Próximo do Limite)"
                    elif dias_limite <= 0:
                        status_ferias = "🚨 Vencido / Multa"
                        
                    lista_ferias.append({
                        "Matrícula": r['Matricula'],
                        "Funcionário": r['Funcionário'],
                        "Setor": r['Setor'],
                        "Admissão": r['Admissão'],
                        "Status Atual": r['Status'],
                        "Últimas Férias": r.get('Ultimas_Ferias', 'Nunca registradas'),
                        "Limite Concessivo": limite_conc.strftime('%d/%m/%Y'),
                        "Situação Férias": status_ferias
                    })
            df_ferias_exibir = pd.DataFrame(lista_ferias)
            st.dataframe(df_ferias_exibir, use_container_width=True)
            
            if not df_ferias_exibir.empty:
                st.download_button(
                    label="📥 Baixar Controle de Férias em Excel",
                    data=converter_df_para_excel(df_ferias_exibir),
                    file_name=f"controle_ferias_{setor_selecionado.lower().replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        elif menu == "Escala Inteligente de Férias":
            ferias.renderizar_modulo_ferias(df_filtrado)

        elif menu == "Faltas & Folgas":
            st.subheader(f"📌 Lançamento & Gestão de Faltas e Folgas - {setor_selecionado}")
            
            tab_cad, tab_hist = st.tabs(["➕ Lançar Ocorrência / Folga", "📋 Histórico de Lançamentos"])
            
            with tab_cad:
                colabs_ativos = df_filtrado[df_filtrado['Status'].isin(['Ativo', 'Férias'])]
                if colabs_ativos.empty:
                    st.warning("Nenhum colaborador encontrado para este setor.")
                else:
                    lista_nomes = sorted(colabs_ativos['Funcionário'].unique())
                    
                    with st.form("form_falta"):
                        st.markdown("##### 👤 Colaborador & Tipo de Ocorrência")
                        
                        nome_colab = st.selectbox(
                            "Digite as primeiras letras para buscar o colaborador:", 
                            lista_nomes
                        )
                        
                        c_t1, c_t2, c_t3 = st.columns([1.5, 1, 1])
                        tipo_falta = c_t1.selectbox("Tipo de Ocorrência", ["Falta Injustificada", "Atestado Médico", "Folga Concedida"])
                        data_falta = c_t2.date_input("Data do Ocorrido", value=hoje, format="DD/MM/YYYY")
                        dias_falta = c_t3.number_input("Qtd. de Dias", min_value=1, max_value=60, value=1, step=1)
                        
                        st.markdown("##### 📄 Informações Complementares")
                        c_c1, c_c2 = st.columns([1, 2])
                        
                        cid_input = c_c1.text_input("Código CID (Se atestado)", value="")
                        motivo_obs = c_c2.text_input("Observação / Motivo Geral", value="")
                        
                        btn_salvar = st.form_submit_button("💾 Salvar Lançamento")
                        
                        if btn_salvar and nome_colab:
                            d_colab = colabs_ativos[colabs_ativos['Funcionário'] == nome_colab].iloc[0]
                            
                            cid_final = "-"
                            if tipo_falta == "Atestado Médico":
                                cid_final = cid_input.strip().upper() if cid_input.strip() else "NÃO INFORMADO"
                            
                            nova_ocorrencia = {
                                "Matricula": str(d_colab['Matricula']),
                                "Funcionário": nome_colab,
                                "Setor": d_colab['Setor'],
                                "Data": data_falta.strftime('%d/%m/%Y'),
                                "Tipo": tipo_falta,
                                "Dias": dias_falta,
                                "CID": cid_final,
                                "Motivo": motivo_obs,
                                "dt_falta": data_falta
                            }
                            df_faltas = pd.concat([df_faltas, pd.DataFrame([nova_ocorrencia])], ignore_index=True)
                            salvar_faltas(df_faltas)
                            st.success(f"Ocorrência de **{tipo_falta}** ({dias_falta} dia(s)) lançada para **{nome_colab}** com sucesso!")
                            st.rerun()

            with tab_hist:
                if df_faltas_filtrado.empty:
                    st.info("Nenhum registro cadastrado até o momento.")
                else:
                    st.markdown("### 📊 Total de Dias por Tipo de Ocorrência")
                    resumo = df_faltas_filtrado.groupby(['Funcionário', 'Tipo'])['Dias'].sum().unstack(fill_value=0).reset_index()
                    st.dataframe(resumo, use_container_width=True)
                    
                    st.markdown("---")
                    st.markdown("### 📋 Histórico Detalhado")
                    df_exibir_f = df_faltas_filtrado[['Data', 'Funcionário', 'Setor', 'Tipo', 'Dias', 'CID', 'Motivo']].copy()
                    st.dataframe(df_exibir_f, use_container_width=True)
                    
                    st.download_button(
                        label="📥 Baixar Histórico de Ocorrências em Excel",
                        data=converter_df_para_excel(df_exibir_f),
                        file_name=f"historico_ausencias_{setor_selecionado.lower().replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        elif menu == "Aniversariantes do Mês":
            st.subheader(f"🎂 Aniversariantes do Mês - {setor_selecionado}")
            mes_sel = st.selectbox("Selecione o Mês", range(1, 13), index=hoje.month - 1)
            df_niver = df_filtrado[df_filtrado['dt_nasc'].apply(lambda d: d.month == mes_sel if pd.notnull(d) else False)].copy()
            if not df_niver.empty:
                df_niver['dia'] = df_niver['dt_nasc'].apply(lambda d: d.day)
                df_niver = df_niver.sort_values(by='dia')
                for _, r in df_niver.iterrows():
                    st.write(f"🎈 **Dia {r['dia']:02d}** - {r['Funcionário']} ({r['Cargo']} - Setor: {r['Setor']})")
            else:
                st.info("Nenhum aniversariante neste setor/mês.")

        elif menu == "Cadastrar / Editar Colaborador":
            st.subheader("👥 Gestão do Cadastro de Colaboradores")
            
            tab_cad_novo, tab_edit_colab = st.tabs(["➕ Cadastrar Novo Colaborador", "✏️ Editar / Atualizar Cadastro Existente"])
            
            lista_status = ["Ativo", "Férias", "Atestado", "Afastado", "INSS", "Desligado", "Desligado (Quebra Experiencia)"]

            with tab_cad_novo:
                with st.form("form_cad"):
                    c1, c2, c3 = st.columns(3)
                    vaga = c1.text_input("Vaga (Digite o Número)", value="1")
                    matricula = c2.text_input("Matrícula (Digite o Número/Código)", value="")
                    nome = c3.text_input("Nome Completo")
                    
                    c4, c5, c6 = st.columns(3)
                    adm = c4.date_input("Data de Admissão", value=hoje, format="DD/MM/YYYY")
                    nasc = c5.date_input("Data de Nascimento", value=date(1995, 1, 1), format="DD/MM/YYYY")
                    contato = c6.text_input("Contato / Telefone")
                    
                    c7, c8, c9 = st.columns(3)
                    lista_cargos = ["Separador", "Conferente", "Auxiliar de Carregamento", "Gerente de Separação", "Supervisor de Seção", "Assistente Supervisor Master", "Assistente Supervisor", "Administrativo"]
                    lista_setores_cad = ["Separação", "Conferência", "Carregamento", "Liderança", "Limpeza", "Administrativo"]
                    cargo = c7.selectbox("Cargo", lista_cargos)
                    setor = c8.selectbox("Setor", lista_setores_cad)
                    status = c9.selectbox("Status Inicial", ["Ativo", "Atestado", "Afastado", "INSS", "Férias"])
                    salvar = st.form_submit_button("💾 Salvar Registro")
                    
                    if salvar and nome:
                        novo_reg = {
                            "Vaga": vaga.strip(), "Matricula": matricula.strip(), "Funcionário": nome.strip(),
                            "Admissão": adm.strftime('%d/%m/%Y'), "Nascimento": nasc.strftime('%d/%m/%Y'),
                            "Contato": contato.strip(), "Cargo": cargo, "Status": status, "Setor": setor,
                            "Ultimas_Ferias": None, "Decisao_Experiencia": None
                        }
                        df_novo = pd.concat([df, pd.DataFrame([novo_reg])], ignore_index=True)
                        salvar_dados(df_novo)
                        st.success(f"Colaborador **{nome}** cadastrado com sucesso!")
                        st.rerun()

            with tab_edit_colab:
                lista_todos_colabs = sorted(df['Funcionário'].dropna().unique())
                colab_sel = st.selectbox("Selecione o Colaborador para Editar:", lista_todos_colabs)
                
                if colab_sel:
                    idx_colab = df[df['Funcionário'] == colab_sel].index[0]
                    dados_c = df.loc[idx_colab]
                    
                    with st.form("form_edit"):
                        st.info(f"Editando dados de **{dados_c['Funcionário']}** (Matrícula Atual: {dados_c['Matricula']})")
                        
                        e1, e2, e3 = st.columns(3)
                        e_vaga = e1.text_input("Vaga", value=str(dados_c['Vaga']) if pd.notnull(dados_c['Vaga']) else "1")
                        e_matr = e2.text_input("Matrícula", value=str(dados_c['Matricula']) if pd.notnull(dados_c['Matricula']) else "")
                        e_nome = e3.text_input("Nome Completo", value=str(dados_c['Funcionário']))
                        
                        dt_adm_val = dados_c['dt_adm'] if pd.notnull(dados_c['dt_adm']) else hoje
                        dt_nasc_val = dados_c['dt_nasc'] if pd.notnull(dados_c['dt_nasc']) else date(1995, 1, 1)
                        
                        e4, e5, e6 = st.columns(3)
                        e_adm = e4.date_input("Data de Admissão", value=dt_adm_val, format="DD/MM/YYYY")
                        e_nasc = e5.date_input("Data de Nascimento", value=dt_nasc_val, format="DD/MM/YYYY")
                        e_contato = e6.text_input("Contato / Telefone", value=str(dados_c['Contato']) if pd.notnull(dados_c['Contato']) else "")
                        
                        lista_cargos = ["Separador", "Conferente", "Auxiliar de Carregamento", "Gerente de Separação", "Supervisor de Seção", "Assistente Supervisor Master", "Assistente Supervisor", "Administrativo"]
                        lista_setores_cad = ["Separação", "Conferência", "Carregamento", "Liderança", "Limpeza", "Administrativo"]
                        
                        idx_cargo = lista_cargos.index(dados_c['Cargo']) if dados_c['Cargo'] in lista_cargos else 0
                        idx_setor = lista_setores_cad.index(dados_c['Setor']) if dados_c['Setor'] in lista_setores_cad else 0
                        
                        idx_status = lista_status.index(dados_c['Status']) if dados_c['Status'] in lista_status else 0
                        
                        e7, e8, e9 = st.columns(3)
                        e_cargo = e7.selectbox("Cargo", lista_cargos, index=idx_cargo)
                        e_setor = e8.selectbox("Setor", lista_setores_cad, index=idx_setor)
                        e_status = e9.selectbox("Status Atual", lista_status, index=idx_status)
                        
                        btn_atualizar = st.form_submit_button("✏️ Atualizar Dados do Colaborador")
                        
                        if btn_atualizar:
                            df.loc[idx_colab, 'Vaga'] = str(e_vaga.strip())
                            df.loc[idx_colab, 'Matricula'] = str(e_matr.strip())
                            df.loc[idx_colab, 'Funcionário'] = str(e_nome.strip())
                            df.loc[idx_colab, 'Admissão'] = e_adm.strftime('%d/%m/%Y')
                            df.loc[idx_colab, 'Nascimento'] = e_nasc.strftime('%d/%m/%Y')
                            df.loc[idx_colab, 'Contato'] = str(e_contato.strip())
                            df.loc[idx_colab, 'Cargo'] = e_cargo
                            df.loc[idx_colab, 'Setor'] = e_setor
                            df.loc[idx_colab, 'Status'] = e_status
                            
                            salvar_dados(df)
                            st.success(f"Dados de **{e_nome}** atualizados com sucesso!")
                            st.rerun()

        elif menu == "📥 Importar Nova Base":
            st.subheader("📥 Atualizar Base Geral de Colaboradores (.xlsx)")
            st.info("Faça o upload de um novo arquivo Excel para substituir ou atualizar a planilha `equipe.xlsx` principal.")
            
            arquivo_upload = st.file_uploader("Arraste ou selecione o arquivo .xlsx", type=["xlsx"])
            if arquivo_upload is not None:
                if st.button("Confirmar e Substituir Base"):
                    try:
                        df_novo_up = pd.read_excel(arquivo_upload)
                        df_novo_up.to_excel(ARQUIVO_DADOS, index=False)
                        st.success("Nova base importada com sucesso! Atualizando sistema...")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao processar arquivo Excel: {e}")
