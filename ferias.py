import io
from datetime import datetime, date, timedelta
import pandas as pd
import streamlit as st

# --- GERADOR DE RELATÓRIO EM PDF PARA O MÓDULO DE FÉRIAS ---
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
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=13,
        leading=16,
        textColor=colors.HexColor("#1E3A8A"),
        spaceAfter=5
    )
    
    sub_style = ParagraphStyle(
        'SubStyle',
        parent=styles['Normal'],
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#334155"),
        spaceAfter=10
    )

    hoje_txt = datetime.now().strftime("%d/%m/%Y às %H:%M")
    elements.append(Paragraph(f"<b>{titulo}</b>", title_style))
    elements.append(Paragraph(
        f"<b>Gerado em:</b> {hoje_txt} | <b>Empresa:</b> Tropical Distribuidora<br/>"
        f"<b>Regras de Negócio:</b> Início aos Domingos | Cota Máx: 2 colabs/mês por setor (Fev-Nov) | Coletivas Livres (Dez-Jan) | Pré-Agendamento Auditável.",
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

def renderizar_modulo_ferias(df):
    st.subheader("🏖️ Módulo de Pré-Agendamento & Simulação Flexível de Férias")

    if df.empty:
        st.warning("Nenhum dado de colaborador disponível para gerar a escala.")
        return

    hoje = date.today()
    df_ativos = df[df['Status'].isin(['Ativo', 'Férias'])].copy()

    if df_ativos.empty:
        st.info("Não há colaboradores ativos para escalonamento.")
        return

    # Inicialização de colunas operacionais
    if 'Aprovacao_RH' not in df.columns:
        df['Aprovacao_RH'] = 'Pendente'
    if 'Fracionamento' not in df.columns:
        df['Fracionamento'] = '30 Dias Corridos'
    if 'Escala_Confirmada' not in df.columns:
        df['Escala_Confirmada'] = False
    if 'Data_Pre_Agendada' not in df.columns:
        df['Data_Pre_Agendada'] = None

    lista_temp = []
    regulares_cnt = 0
    atencao_cnt = 0
    vencidos_cnt = 0

    for idx, r in df_ativos.iterrows():
        adm = r.get('dt_adm')
        ult_ferias = pd.to_datetime(r.get('Ultimas_Ferias'), dayfirst=True, errors='coerce').date() if pd.notnull(r.get('Ultimas_Ferias')) else None
        data_base = ult_ferias if pd.notnull(ult_ferias) else adm

        if pd.notnull(data_base):
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
                'idx': idx,
                'colab': r,
                'data_base': data_base,
                'limite_conc': limite_conc,
                'dias_restantes': dias_restantes,
                'ult_ferias': ult_ferias
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
            setor = r.get('Setor', 'Geral')

            # Sugestão inicial do algoritmo
            data_sugerida_inicial = max(hoje + timedelta(days=30), limite_conc - timedelta(days=60))
            data_algoritmo = ajustar_para_domingo(data_sugerida_inicial)

            # Data salva ou editada manualmente
            dt_pre_salva = r.get('Data_Pre_Agendada')
            if pd.notnull(dt_pre_salva) and str(dt_pre_salva).strip() != "":
                data_efetiva = pd.to_datetime(dt_pre_salva, dayfirst=True, errors='coerce').date()
                if pd.isnull(data_efetiva):
                    data_efetiva = data_algoritmo
            else:
                data_efetiva = data_algoritmo

            dt_adm_str = r['dt_adm'].strftime('%d/%m/%Y') if pd.notnull(r.get('dt_adm')) else (r.get('Admissão', 'N/A'))
            dt_ult_str = ult_ferias.strftime('%d/%m/%Y') if pd.notnull(ult_ferias) else 'Não Registrada'

            c_info, c_dt, c_status = st.columns([2.2, 1.3, 1.5])

            with c_info:
                st.write(f"👤 **{r['Funcionário']}** (Prio #{ordem_prio}) | Setor: **{setor}** | Cargo: {r.get('Cargo', 'N/A')}")
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
                    df.at[idx, 'Data_Pre_Agendada'] = data_inicio_domingo.strftime('%d/%m/%Y')
                    alterou = True

            # Validação de Capacidade e Alertas Legais
            chave_mes = (setor, data_inicio_domingo.strftime("%Y-%m"))
            qtd_agendada = ocupacao_setor_mes.get(chave_mes, 0)
            is_férias_coletivas = data_inicio_domingo.month in [12, 1]
            ocupacao_setor_mes[chave_mes] = qtd_agendada + 1

            data_aviso_rh = data_inicio_domingo - timedelta(days=30)

            with c_status:
                if data_inicio_domingo > limite_conc:
                    st.error("🚨 Ultrapassa Limite Legal!")
                elif qtd_agendada >= 2 and not is_férias_coletivas:
                    st.warning(f"⚠️ Cota Excedida no Setor ({qtd_agendada + 1}º em {data_inicio_domingo.strftime('%m/%Y')})")
                else:
                    st.success(f"✅ Início no Domingo: **{data_inicio_domingo.strftime('%d/%m/%Y')}**")

                st.caption(f"Emissão Aviso RH até: **{data_aviso_rh.strftime('%d/%m/%Y')}**")

            c_frac, c_aprov, c_conf = st.columns([1.5, 1.5, 1])

            with c_frac:
                frac_atual = r.get('Fracionamento') if pd.notnull(r.get('Fracionamento')) else '30 Dias Corridos'
                opcoes_frac = ['30 Dias Corridos', '15 + 15 Dias', '20 + 10 Dias']
                idx_f = opcoes_frac.index(frac_atual) if frac_atual in opcoes_frac else 0
                novo_frac = st.selectbox("Fracionamento", opcoes_frac, index=idx_f, key=f"frac_{idx}")
                if novo_frac != frac_atual:
                    df.at[idx, 'Fracionamento'] = novo_frac
                    alterou = True

            with c_aprov:
                aprov_atual = r.get('Aprovacao_RH') if pd.notnull(r.get('Aprovacao_RH')) else 'Pendente'
                opcoes_aprov = ['Pendente', 'Pré-Agendado RH', 'Aprovado RH', 'Em Análise', 'Rejeitado']
                idx_a = opcoes_aprov.index(aprov_atual) if aprov_atual in opcoes_aprov else 0
                nova_aprov = st.selectbox("Status RH", opcoes_aprov, index=idx_a, key=f"aprov_{idx}")
                if nova_aprov != aprov_atual:
                    df.at[idx, 'Aprovacao_RH'] = nova_aprov
                    alterou = True

            with c_conf:
                conf_atual = bool(r.get('Escala_Confirmada')) if pd.notnull(r.get('Escala_Confirmada')) else False
                nova_conf = st.checkbox("Confirmar?", value=conf_atual, key=f"conf_{idx}")
                if nova_conf != conf_atual:
                    df.at[idx, 'Escala_Confirmada'] = nova_conf
                    alterou = True

            st.divider()

            cota_txt = f"Vaga {qtd_agendada + 1}/2 no Setor" if not is_férias_coletivas else "Férias Coletivas (Sem Cota)"

            dados_escala.append({
                "Prio": f"#{ordem_prio}",
                "Matrícula": r.get('Matricula', 'N/A'),
                "Funcionário": r['Funcionário'],
                "Setor": setor,
                "Cargo": r.get('Cargo', 'N/A'),
                "Admissão": dt_adm_str,
                "Últimas Férias": dt_ult_str,
                "Início Férias (Domingo)": data_inicio_domingo.strftime('%d/%m/%Y'),
                "Prazo Aviso RH": data_aviso_rh.strftime('%d/%m/%Y'),
                "Limite Concessivo": limite_conc.strftime('%d/%m/%Y'),
                "Status Cota Setor": cota_txt,
                "Fracionamento": df.at[idx, 'Fracionamento'],
                "Status RH": df.at[idx, 'Aprovacao_RH'],
                "Confirmado": "SIM" if df.at[idx, 'Escala_Confirmada'] else "NÃO"
            })

    if alterou:
        cols_salvar = [c for c in df.columns if c not in ['dt_adm', 'dt_nasc', 'dt_nasc_dt', 'dt_ult_ferias', 'exp_45', 'exp_90', 'dias_para_45', 'dias_para_90']]
        df[cols_salvar].to_excel("equipe.xlsx", index=False)
        st.success("✅ Pré-agendamentos e alterações salvos na base!")
        st.rerun()

    df_escala_final = pd.DataFrame(dados_escala)

    with tab_resumo:
        st.markdown("##### 📋 Visão Consolidada da Escala de Férias Reorganizada")
        st.dataframe(df_escala_final, use_container_width=True)

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
            pdf_bytes = gerar_pdf_ferias(
                "Relatório - Escala Inteligente de Férias e Pré-Agendamento RH",
                df_escala_final
            )
            st.download_button(
                label="🖨️ Baixar PDF para Impressão",
                data=pdf_bytes,
                file_name=f"escala_inteligente_ferias_{hoje.strftime('%d_%m_%Y')}.pdf",
                mime="application/pdf",
                key="btn_pdf_pre"
            )
