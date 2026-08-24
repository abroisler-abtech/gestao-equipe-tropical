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
        f"<b>Regras de Negócio:</b> Início aos Domingos | Cota Máx: 2 colabs/mês por setor (Fev-Nov) | Coletivas Livres (Dez-Jan) | Aviso RH 30 dias.",
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
    dias_para_domingo = (6 - data_val.weekday()) % 7
    return data_val + timedelta(days=dias_para_domingo)

def renderizar_modulo_ferias(df):
    st.subheader("🏖️ Painel Inteligente de Gestão de Férias")

    if df.empty:
        st.warning("Nenhum dado de colaborador disponível para gerar a escala.")
        return

    hoje = date.today()
    df_ativos = df[df['Status'].isin(['Ativo', 'Férias'])].copy()

    if df_ativos.empty:
        st.info("Não há colaboradores ativos para escalonamento.")
        return

    # Garante a existência das colunas operacionais no DataFrame
    if 'Aprovacao_RH' not in df.columns:
        df['Aprovacao_RH'] = 'Pendente'
    if 'Fracionamento' not in df.columns:
        df['Fracionamento'] = '30 Dias Corridos'
    if 'Escala_Confirmada' not in df.columns:
        df['Escala_Confirmada'] = False

    # Processamento e ordenação da escala pelo limite concessivo
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

    # --- CARDS DE INDICADORES NO TOPO ---
    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric("🟢 Férias Regulares", regulares_cnt)
    c_m2.metric("🟡 Atenção (Próximos 60d)", atencao_cnt)
    c_m3.metric("🚨 Vencidos / Risco Multa", vencidos_cnt)

    st.markdown("---")

    # Mapeamento da Escala Sugerida
    ocupacao_setor_mes = {}
    dados_tabela_editavel = []

    for ordem_prio, item in enumerate(lista_temp, start=1):
        idx = item['idx']
        r = item['colab']
        limite_conc = item['limite_conc']
        ult_ferias = item['ult_ferias']
        setor = r.get('Setor', 'Geral')

        data_alvo = max(hoje + timedelta(days=30), limite_conc - timedelta(days=60))
        data_inicio = ajustar_para_domingo(data_alvo)

        while True:
            chave_mes = (setor, data_inicio.strftime("%Y-%m"))
            qtd_na_mesma_pasta = ocupacao_setor_mes.get(chave_mes, 0)
            is_ferias_coletivas = data_inicio.month in [12, 1]

            if qtd_na_mesma_pasta < 2 or is_ferias_coletivas:
                ocupacao_setor_mes[chave_mes] = qtd_na_mesma_pasta + 1
                regra_status = f"Vaga {qtd_na_mesma_pasta + 1}/2 Setor" if not is_ferias_coletivas else "Coletivas (Livre)"
                break
            else:
                data_inicio = ajustar_para_domingo(data_inicio + timedelta(days=28))

        data_aviso_rh = data_inicio - timedelta(days=30)
        dt_adm_str = r['dt_adm'].strftime('%d/%m/%Y') if pd.notnull(r.get('dt_adm')) else (r.get('Admissão', 'N/A'))
        dt_ult_str = ult_ferias.strftime('%d/%m/%Y') if pd.notnull(ult_ferias) else 'Não Registrada'

        frac_atual = r.get('Fracionamento') if pd.notnull(r.get('Fracionamento')) else '30 Dias Corridos'
        aprov_atual = r.get('Aprovacao_RH') if pd.notnull(r.get('Aprovacao_RH')) else 'Pendente'
        conf_atual = bool(r.get('Escala_Confirmada')) if pd.notnull(r.get('Escala_Confirmada')) else False

        dados_tabela_editavel.append({
            "ID_Original": idx,
            "Prio": f"#{ordem_prio}",
            "Funcionário": r['Funcionário'],
            "Setor": setor,
            "Cargo": r.get('Cargo', 'N/A'),
            "Admissão": dt_adm_str,
            "Início Férias (Domingo)": data_inicio.strftime('%d/%m/%Y'),
            "Aviso RH Até": data_aviso_rh.strftime('%d/%m/%Y'),
            "Limite Concessivo": limite_conc.strftime('%d/%m/%Y'),
            "Cota Setor": regra_status,
            "Fracionamento": frac_atual,
            "Status RH": aprov_atual,
            "Confirmado": conf_atual
        })

    df_escala_base = pd.DataFrame(dados_tabela_editavel)

    # --- ESTRUTURA EM ABAS (TABS) ---
    tab_edit, tab_resumo, tab_export = st.tabs([
        "✏️ Gestão & Aprovação Rápida RH", 
        "📋 Visão Consolidada da Escala", 
        "📥 Impressão e Exportação"
    ])

    with tab_edit:
        st.markdown("##### ✏️ Altere as opções diretamente na tabela e clique em Salvar:")
        
        df_editor = df_escala_base.drop(columns=["ID_Original"])

        df_editado = st.data_editor(
            df_editor,
            hide_index=True,
            column_config={
                "Fracionamento": st.column_config.SelectboxColumn(
                    "Fracionamento",
                    options=['30 Dias Corridos', '15 + 15 Dias', '20 + 10 Dias'],
                    required=True
                ),
                "Status RH": st.column_config.SelectboxColumn(
                    "Status RH",
                    options=['Pendente', 'Aprovado RH', 'Em Análise', 'Rejeitado'],
                    required=True
                ),
                "Confirmado": st.column_config.CheckboxColumn(
                    "Confirmado?",
                    default=False
                )
            },
            disabled=["Prio", "Funcionário", "Setor", "Cargo", "Admissão", "Início Férias (Domingo)", "Aviso RH Até", "Limite Concessivo", "Cota Setor"],
            use_container_width=True,
            key="editor_ferias"
        )

        if st.button("💾 Salvar Todas as Alterações da Escala", type="primary"):
            for i, row in df_editado.iterrows():
                idx_orig = df_escala_base.loc[i, "ID_Original"]
                df.at[idx_orig, 'Fracionamento'] = row["Fracionamento"]
                df.at[idx_orig, 'Aprovacao_RH'] = row["Status RH"]
                df.at[idx_orig, 'Escala_Confirmada'] = row["Confirmado"]

            cols_salvar = [c for c in df.columns if c not in ['dt_adm', 'dt_nasc', 'dt_nasc_dt', 'dt_ult_ferias', 'exp_45', 'exp_90', 'dias_para_45', 'dias_para_90']]
            df[cols_salvar].to_excel("equipe.xlsx", index=False)
            st.success("✅ Todas as alterações foram salvas com sucesso na base!")
            st.rerun()

    with tab_resumo:
        st.markdown("##### 📋 Tabela Processada da Escala Inteligente")
        
        df_resumo_exibir = df_escala_base.drop(columns=["ID_Original"]).copy()
        df_resumo_exibir["Confirmado"] = df_resumo_exibir["Confirmado"].apply(lambda x: "SIM" if x else "NÃO")
        
        st.dataframe(df_resumo_exibir, use_container_width=True)

    with tab_export:
        st.markdown("##### 📥 Exportar Relatórios Formatados para Impressão ou Auditoria")
        
        df_exp_final = df_escala_base.drop(columns=["ID_Original"]).copy()
        df_exp_final["Confirmado"] = df_exp_final["Confirmado"].apply(lambda x: "SIM" if x else "NÃO")

        c_down1, c_down2 = st.columns(2)
        
        with c_down1:
            st.download_button(
                label="📥 Baixar Escala em Excel (.xlsx)",
                data=converter_df_para_excel(df_exp_final),
                file_name=f"escala_inteligente_ferias_{hoje.strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_down_excel_ferias"
            )
            
        with c_down2:
            pdf_bytes = gerar_pdf_ferias(
                "Relatório - Escala Inteligente de Férias e Aprovação RH",
                df_exp_final
            )
            st.download_button(
                label="🖨️ Baixar PDF para Impressão",
                data=pdf_bytes,
                file_name=f"escala_inteligente_ferias_{hoje.strftime('%d_%m_%Y')}.pdf",
                mime="application/pdf",
                key="btn_down_pdf_ferias"
            )
