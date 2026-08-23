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
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=16,
        leading=20,
        textColor=colors.HexColor("#1E3A8A"),
        spaceAfter=15,
        alignment=0
    )
    
    hoje_txt = datetime.now().strftime("%d/%m/%Y às %H:%M")
    elements.append(Paragraph(f"<b>{titulo}</b>", title_style))
    elements.append(Paragraph(f"<font size=9 color='#666666'>Gerado em: {hoje_txt} | Tropical Distribuidora</font>", styles['Normal']))
    elements.append(Spacer(1, 15))

    colunas = list(df_escala.columns)
    table_data = [[Paragraph(f"<b>{col}</b>", styles['Normal']) for col in colunas]]
    
    for _, linha in df_escala.iterrows():
        row_data = []
        for item in linha:
            val_str = str(item) if pd.notnull(item) else ""
            row_data.append(Paragraph(val_str, styles['Normal']))
        table_data.append(row_data)

    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E2E8F0")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor("#1E293B")),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#94A3B8")),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
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

def renderizar_modulo_ferias(df):
    st.subheader("🏖️ Escala Inteligente de Férias")
    st.info("Planejador e distribuidor de períodos concessivos com prevenção de acúmulo e sobreposição de equipe.")

    if df.empty:
        st.warning("Nenhum dado de colaborador disponível para gerar a escala.")
        return

    hoje = date.today()
    df_ativos = df[df['Status'].isin(['Ativo', 'Férias'])].copy()

    if df_ativos.empty:
        st.info("Não há colaboradores ativos para escalonamento.")
        return

    lista_escala = []

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

            sugestao_mes = limite_conc - timedelta(days=60)
            mes_sugerido_txt = sugestao_mes.strftime("%m/%Y")

            situacao = "✅ Regular"
            if dias_restantes <= 60 and dias_restantes > 0:
                situacao = "⚠️ Agendar Imediatamente"
            elif dias_restantes <= 0:
                situacao = "🚨 Vencido (Risco Multa)"

            lista_escala.append({
                "Matrícula": r.get('Matricula', 'N/A'),
                "Funcionário": r['Funcionário'],
                "Setor": r.get('Setor', 'N/A'),
                "Cargo": r.get('Cargo', 'N/A'),
                "Admissão": r.get('Admissão', 'N/A'),
                "Últimas Férias": r.get('Ultimas_Ferias', 'Não Registrada'),
                "Limite Concessivo": limite_conc.strftime('%d/%m/%Y'),
                "Dias pro Limite": dias_restantes,
                "Sugestão de Mês": mes_sugerido_txt,
                "Situação": situacao
            })

    df_escala = pd.DataFrame(lista_escala)

    if not df_escala.empty:
        st.markdown("### 📋 Escala Sugerida do Quadro")
        st.dataframe(df_escala, use_container_width=True)

        st.markdown("---")
        st.markdown("### 📥 Exportar Relatório da Escala")
        
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
                "Relatório - Escala Inteligente de Férias",
                df_escala
            )
            st.download_button(
                label="🖨️ Baixar PDF para Impressão",
                data=pdf_bytes,
                file_name=f"escala_inteligente_ferias_{hoje.strftime('%d_%m_%Y')}.pdf",
                mime="application/pdf"
            )
