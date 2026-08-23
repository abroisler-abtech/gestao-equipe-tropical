import io
from datetime import datetime, date, timedelta
import pandas as pd
import streamlit as st

# --- GERADOR DE RELATÓRIO EM PDF PARA O MÓDULO DE FÉRIAS COM REGRAS NO CABEÇALHO ---
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
        spaceAfter=5,
        alignment=0
    )
    
    subtitle_style = ParagraphStyle(
        'SubtitleStyle',
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
        f"<b>Regras de Negócio Aplicadas:</b> Início no Domingo | Cota Máx: 2 colabs/mês por setor (Fev-Nov) | Coletivas Livres (Dez-Jan) | Aviso RH com 30 dias.",
        subtitle_style
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
    st.subheader("🏖️ Escala Inteligente de Férias com Regras Auditáveis")
    st.info("📌 **Regras Ativas no Relatório:** 1) Início no **Domingo** | 2) Cota de **Máx. 2 colabs/mês por setor** (Fev a Nov) | 3) **Dez/Jan livres** (Coletivas) | 4) **Aviso RH** com 30 dias.")

    if df.empty:
        st.warning("Nenhum dado de colaborador disponível para gerar a escala.")
        return

    hoje = date.today()
    df_ativos = df[df['Status'].isin(['Ativo', 'Férias'])].copy()

    if df_ativos.empty:
        st.info("Não há colaboradores ativos para escalonamento.")
        return

    if 'Aprovacao_RH' not in df.columns:
        df['Aprovacao_RH'] = 'Pendente'
    if 'Fracionamento' not in df.columns:
        df['Fracionamento'] = '30 Dias Corridos'
    if 'Escala_Confirmada' not in df.columns:
        df['Escala_Confirmada'] = False

    # 1. Mapeamento e ordenação por urgência do limite concessivo
    lista_temp = []
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
            
            lista_temp.append({
                'idx': idx,
                'colab': r,
                'data_base': data_base,
                'limite_conc': limite_conc,
                'dias_restantes': dias_restantes,
                'ult_ferias': ult_ferias
            })

    # Ordena rigorosamente pelo limite concessivo mais antigo
    lista_temp = sorted(lista_temp, key=lambda x: x['limite_conc'])

    ocupacao_setor_mes = {}
    lista_escala = []
    alterou_dados = False

    st.markdown("### ⚙️ Programação da Escala & Validação do RH")

    for ordem_prio, item in enumerate(lista_temp, start=1):
        idx = item['idx']
        r = item['colab']
        limite_conc = item['limite_conc']
        dias_restantes = item['dias_restantes']
        ult_ferias = item['ult_ferias']
        setor = r.get('Setor', 'Geral')

        data_alvo = max(hoje + timedelta(days=30), limite_conc - timedelta(days=60))
        data_inicio = ajustar_para_domingo(data_alvo)

        # Processamento de Cota
        while True:
            chave_mes = (setor, data_inicio.strftime("%Y-%m"))
            qtd_na_mesma_pasta = ocupacao_setor_mes.get(chave_mes, 0)
            is_ferias_coletivas = data_inicio.month in [12, 1]

            if qtd_na_mesma_pasta < 2 or is_ferias_coletivas:
                ocupacao_setor_mes[chave_mes] = qtd_na_mesma_pasta + 1
                if is_ferias_coletivas:
                    regra_status = "Coletivas (Sem trava)"
                else:
                    regra_status = f"Vaga {qtd_na_mesma_pasta + 1}/2 Setor (OK)"
                break
            else:
                data_inicio = ajustar_para_domingo(data_inicio + timedelta(days=28))

        data_aviso_rh = data_inicio - timedelta(days=30)
        
        if hoje >= data_aviso_rh and hoje < data_inicio:
            alerta_aviso = "🚨 AVISO HOJE!"
        elif hoje >= data_inicio:
            alerta_aviso = "Em Férias"
        else:
            alerta_aviso = f"Aviso até {data_aviso_rh.strftime('%d/%m/%Y')}"

        dt_adm_str = r['dt_adm'].strftime('%d/%m/%Y') if pd.notnull(r.get('dt_adm')) else (r.get('Admissão', 'N/A'))
        dt_ult_str = ult_ferias.strftime('%d/%m/%Y') if pd.notnull(ult_ferias) else 'Não Registrada'

        c_info, c_frac, c_aprov, c_conf = st.columns([2.3, 1.1, 1.2, 0.8])
        
        with c_info:
            st.write(f"👤 **{r['Funcionário']}** (Prio #{ordem_prio}) | Setor: **{setor}** | Cargo: {r.get('Cargo', 'N/A')}")
            st.caption(f"Admissão: **{dt_adm_str}** | Últs Férias: **{dt_ult_str}** | Limite: **{limite_conc.strftime('%d/%m/%Y')}**")
            st.caption(f"📅 **Início Domingo: {data_inicio.strftime('%d/%m/%Y')}** | Cota: **{regra_status}** | Status Aviso: **{alerta_aviso}**")

        with c_frac:
            frac_atual = r.get('Fracionamento') if pd.notnull(r.get('Fracionamento')) else '30 Dias Corridos'
            opcoes_frac = ['30 Dias Corridos', '15 + 15 Dias', '20 + 10 Dias']
            idx_f = opcoes_frac.index(frac_atual) if frac_atual in opcoes_frac else 0
            novo_frac = st.selectbox("Fracionamento", opcoes_frac, index=idx_f, key=f"frac_{idx}")
            if novo_frac != frac_atual:
                df.at[idx, 'Fracionamento'] = novo_frac
                alterou_dados = True

        with c_aprov:
            aprov_atual = r.get('Aprovacao_RH') if pd.notnull(r.get('Aprovacao_RH')) else 'Pendente'
            opcoes_aprov = ['Pendente', 'Aprovado RH', 'Em Análise', 'Rejeitado']
            idx_a = opcoes_aprov.index(aprov_atual) if aprov_atual in opcoes_aprov else 0
            nova_aprov = st.selectbox("Status RH", opcoes_aprov, index=idx_a, key=f"aprov_{idx}")
            if nova_aprov != aprov_atual:
                df.at[idx, 'Aprovacao_RH'] = nova_aprov
                alterou_dados = True

        with c_conf:
            conf_atual = bool(r.get('Escala_Confirmada')) if pd.notnull(r.get('Escala_Confirmada')) else False
            nova_conf = st.checkbox("Confirmar", value=conf_atual, key=f"conf_{idx}")
            if nova_conf != conf_atual:
                df.at[idx, 'Escala_Confirmada'] = nova_conf
                alterou_dados = True

        st.divider()

        lista_escala.append({
            "Prio": f"#{ordem_prio}",
            "Matrícula": r.get('Matricula', 'N/A'),
            "Funcionário": r['Funcionário'],
            "Setor": setor,
            "Cargo": r.get('Cargo', 'N/A'),
            "Admissão": dt_adm_str,
            "Últimas Férias": dt_ult_str,
            "Início Férias (Domingo)": data_inicio.strftime('%d/%m/%Y'),
            "Prazo Aviso RH": data_aviso_rh.strftime('%d/%m/%Y'),
            "Limite Concessivo": limite_conc.strftime('%d/%m/%Y'),
            "Status Cota Setor": regra_status,
            "Fracionamento": df.at[idx, 'Fracionamento'],
            "Status RH": df.at[idx, 'Aprovacao_RH'],
            "Confirmado": "SIM" if df.at[idx, 'Escala_Confirmada'] else "NÃO"
        })

    if alterou_dados:
        cols_salvar = [c for c in df.columns if c not in ['dt_adm', 'dt_nasc', 'dt_nasc_dt', 'dt_ult_ferias', 'exp_45', 'exp_90', 'dias_para_45', 'dias_para_90']]
        df[cols_salvar].to_excel("equipe.xlsx", index=False)
        st.success("Alterações e confirmações salvas!")
        st.rerun()

    df_escala = pd.DataFrame(lista_escala)

    if not df_escala.empty:
        st.markdown("### 📋 Tabela Auditável da Escala de Férias")
        st.dataframe(df_escala, use_container_width=True)

        st.markdown("---")
        st.markdown("### 📥 Exportar Relatório com Regras")
        
        c_down1, c_down2 = st.columns(2)
        
        with c_down1:
            st.download_button(
                label="📥 Baixar Escala em Excel (.xlsx)",
                data=converter_df_para_excel(df_escala),
                file_name=f"escala_inteligente_ferias_{hoje.strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with c_down2:
            pdf_bytes = gerar_pdf_ferias(
                "Relatório - Escala Inteligente de Férias e Aprovação RH",
                df_escala
            )
            st.download_button(
                label="🖨️ Baixar PDF para Impressão",
                data=pdf_bytes,
                file_name=f"escala_inteligente_ferias_{hoje.strftime('%d_%m_%Y')}.pdf",
                mime="application/pdf"
            )
