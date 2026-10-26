import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import urllib.parse

# Importación de python-docx para la generación del informe médico profesional
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA Y MARCA PERSONAL (Neuromarketing Médico)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Lic. Carlos Pacheco | Fisiología & Salud Vascular",
    page_icon="🫀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados para interfaz médica de alto impacto
st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(135deg, #0A2540 0%, #1A365D 100%);
        color: white;
        padding: 2rem;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .slogan-tag {
        color: #00E6C3;
        font-weight: 600;
        font-size: 1.1rem;
        margin-top: 0.5rem;
    }
    .callout-card {
        background-color: #F8F9FA;
        border-left: 5px solid #0A2540;
        padding: 1.2rem;
        border-radius: 6px;
        margin: 1rem 0;
    }
    .payment-box {
        background-color: #F1F3F5;
        padding: 1.5rem;
        border-radius: 10px;
        border: 1px solid #E9ECEF;
    }
    .stButton>button {
        background-color: #0A2540;
        color: white;
        border-radius: 8px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #00E6C3;
        color: #0A2540;
    }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# FUNCIONES AUXILIARES PARA GENERAR DOCUMENTO DOCX (python-docx)
# -----------------------------------------------------------------------------
def set_cell_background(cell, fill_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin_name, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{margin_name}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def generate_docx_report(patient_name, age, pas, activity, risk_score, risk_category, recommendation):
    doc = Document()
    
    # Márgenes de 1 pulgada
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
        # Encabezado
        header = section.header
        hp = header.paragraphs[0]
        hp.text = "LIC. CARLOS PACHECO | EVALUACIÓN DE SALUD VASCULAR Y ENDOTELIAL"
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hp.style.font.name = 'Arial'
        hp.style.font.size = Pt(8.5)
        hp.style.font.color.rgb = RGBColor(108, 117, 125)

    NAVY_HEX = "0A2540"
    ALT_GRAY_HEX = "F8F9FA"
    NAVY = RGBColor(10, 37, 64)

    # Estilos de Título
    h1 = doc.styles['Heading 1']
    h1.font.name = 'Arial'; h1.font.size = Pt(18); h1.font.bold = True; h1.font.color.rgb = NAVY
    h1.paragraph_format.space_before = Pt(16); h1.paragraph_format.space_after = Pt(6)

    doc.add_heading('Informe Fisiológico de Riesgo Vascular', level=1)
    
    # Datos del Paciente
    p_info = doc.add_paragraph()
    p_info.add_run("Paciente: ").bold = True
    p_info.add_run(f"{patient_name}  |  ")
    p_info.add_run("Edad: ").bold = True
    p_info.add_run(f"{age} años\n")
    p_info.add_run("Presión Arterial Sistólica: ").bold = True
    p_info.add_run(f"{pas} mmHg  |  ")
    p_info.add_run("Nivel de Actividad: ").bold = True
    p_info.add_run(f"{activity}")

    # Tabla de Resultados
    table = doc.add_table(rows=3, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    headers = ["Indicador", "Resultado Evaluado"]
    data = [
        ["Categoría de Riesgo Vascular", risk_category],
        ["Puntaje Hemodinámico Estimado", f"{risk_score} / 100"]
    ]
    
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        set_cell_margins(cell)
        set_cell_background(cell, NAVY_HEX)
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
                
    for row_idx, row_data in enumerate(data, start=1):
        for col_idx, text in enumerate(row_data):
            cell = table.cell(row_idx, col_idx)
            cell.text = text
            set_cell_margins(cell)
            if row_idx % 2 == 0:
                set_cell_background(cell, ALT_GRAY_HEX)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    
    # Recomendación del Especialista
    p_rec = doc.add_paragraph()
    p_rec.add_run("Prescripción de Ejercicio Seguro & Recomendación:").bold = True
    p_rec.runs[0].font.color.rgb = NAVY
    doc.add_paragraph(recommendation)

    # Pie institucional
    p_footer = doc.add_paragraph()
    p_footer.add_run("\nLic. Carlos Pacheco | Fisiólogo & Especialista en Protección Endotelial\nContacto: pachecohx@gmail.com | WhatsApp: +58 412-0582064").italic = True

    # Guardar en buffer de memoria
    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# -----------------------------------------------------------------------------
# ENCABEZADO Y MARCA PERSONAL
# -----------------------------------------------------------------------------
st.markdown("""
    <div class="main-header">
        <h1 style="margin:0; font-size:2.5rem;">Lic. Carlos Pacheco</h1>
        <h3 style="margin:0; font-weight:400; color:#E2E8F0;">Fisiólogo & Salud Vascular a tu Alcance</h3>
        <p class="slogan-tag">Especialista en Prescripción de Ejercicios Seguros y Protección Endotelial</p>
    </div>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# NAVEGACIÓN Y PESTAÑAS INTERACTIVAS
# -----------------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "🫀 Evaluación & Calculadora", 
    "💼 Servicios Médicos", 
    "💳 Pasarela de Pagos", 
    "📚 Blog de Fisiología"
])

# -----------------------------------------------------------------------------
# TAB 1: CALCULADORA DE RIESGO VASCULAR
# -----------------------------------------------------------------------------
with tab1:
    st.header("Calculadora Interactiva de Riesgo Endotelial y Presión Arterial")
    st.write("Completa tus datos para obtener una estratificación fisiológica inmediata y generar tu informe médico.")

    col1, col2 = st.columns([1, 1])

    with col1:
        patient_name = st.text_input("Nombre Completo del Paciente", value="Carlos Rodríguez")
        age = st.number_input("Edad (años)", min_value=18, max_value=100, value=54)
        pas = st.slider("Presión Arterial Sistólica (mmHg)", min_value=90, max_value=200, value=135)
        activity = st.selectbox("Frecuencia de Actividad Física", ["Sedentario", "1-2 días/semana (Leve)", "3-4 días/semana (Moderado)", "5+ días/semana (Avanzado)"])
        antecedents = st.multiselect("Factores de Riesgo / Antecedentes", ["Hipertensión Diagnosticada", "Ateromas / Placa Carotídea", "Diabetes / Resistencia a Insulina", "Tabaquismo", "Estrés Crónico"])

    # Cálculo del algoritmo fisiológico
    risk_score = 20
    if pas > 130: risk_score += 25
    if pas > 150: risk_score += 20
    if age > 50: risk_score += 15
    if activity == "Sedentario": risk_score += 15
    risk_score += len(antecedents) * 10
    risk_score = min(risk_score, 100)

    if risk_score < 40:
        risk_cat = "Bajo Riesgo (Endotelio Saludable)"
        color_code = "#28A745"
        recommendation = "Tu estado hemodinámico es estable. Se recomienda un plan de mantenimiento físico aeróbico de 150 min/semana para preservar la flexibilidad arterial."
    elif risk_score < 70:
        risk_cat = "Riesgo Moderado (Rigidez Arterial Detectada)"
        color_code = "#FFC107"
        recommendation = "Se observa elevación de tensión y resistencia vascular. Requiere prescripción de ejercicio dosificado para estimular la liberación de óxido nítrico endotelial sin picos de presión."
    else:
        risk_cat = "Alto Riesgo (Vulnerabilidad Vascular Elevada)"
        color_code = "#DC3545"
        recommendation = "Cuidado: Existe alto estrés de cizallamiento y presión sistólica. Está contraindicado el ejercicio de alta intensidad sin supervisión. Se requiere un protocolo personalizado de protección endotelial de urgencia."

    with col2:
        st.subheader("Resultados del Perfil Hemodinámico")
        
        # Gráfico dinámico Gauge Plot
        fig = go.Figure(go.Indicator(
            mode = "gauge+number",
            value = risk_score,
            domain = {'x': [0, 1], 'y': [0, 1]},
            title = {'text': f"Puntaje de Riesgo: {risk_cat}"},
            gauge = {
                'axis': {'range': [0, 100]},
                'bar': {'color': color_code},
                'steps': [
                    {'range': [0, 40], 'color': "#E8F5E9"},
                    {'range': [40, 70], 'color': "#FFFDE7"},
                    {'range': [70, 100], 'color': "#FFEBEE"}
                ]
            }
        ))
        fig.update_layout(height=280, margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig, use_container_width=True)

        st.markdown(f"""
            <div class="callout-card">
                <h4 style="margin:0; color:#0A2540;">Recomendación del Fisiólogo:</h4>
                <p style="margin-top:0.5rem;">{recommendation}</p>
            </div>
        """, unsafe_allow_html=True)

        # Generar informe en formato Word (docx)
        docx_file = generate_docx_report(patient_name, age, pas, activity, risk_score, risk_cat, recommendation)
        
        st.download_button(
            label="📄 Descargar Informe Médico Oficial (.docx)",
            data=docx_file,
            file_name=f"Informe_Vascular_{patient_name.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        # Botón de consulta inmediata por WhatsApp
        msg_encoded = urllib.parse.quote(f"Hola Lic. Carlos Pacheco, realicé la evaluación en su plataforma para {patient_name} ({age} años). Presión Sistólica: {pas} mmHg. Riesgo: {risk_cat}. Deseo coordinar una consulta.")
        wa_url = f"https://wa.me/584120582064?text={msg_encoded}"
        st.markdown(f'<a href="{wa_url}" target="_blank"><button style="width:100%; padding:10px; background-color:#25D366; color:white; border:none; border-radius:8px; font-weight:bold; cursor:pointer;">📱 Enviar Evaluación por WhatsApp (+58 412-0582064)</button></a>', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# TAB 2: SERVICIOS MÉDICOS OFRECIDOS
# -----------------------------------------------------------------------------
with tab2:
    st.header("Servicios Especializados de Fisiología & Salud Vascular")
    st.write("Programas diseñados con base científica y neuromarketing médico para garantizar resultados sostenibles.")

    s_col1, s_col2, s_col3 = st.columns(3)

    with s_col1:
        st.subheader("🩸 Prescripción de Ejercicio Seguro")
        st.write("Diseño de rutina personalizada para personas con hipertensión, cardiopatías o ateromas carotídeos. Control hemodinámico estricto.")
        st.caption("Inversión: $45 USD / Mensual")
        if st.button("Reservar Prescripción", key="btn1"):
            st.info("Utiliza la pestaña 'Pasarela de Pagos' o contáctanos por WhatsApp para formalizar tu cita.")

    with s_col2:
        st.subheader("🛡️ Protocolo de Protección Endotelial")
        st.write("Evaluación integral del riesgo arterial, ajuste de dosis de esfuerzo físico y asesoría para reducir el estrés de cizallamiento.")
        st.caption("Inversión: $60 USD / Sesión")
        if st.button("Reservar Protección", key="btn2"):
            st.info("Utiliza la pestaña 'Pasarela de Pagos' o contáctanos por WhatsApp para formalizar tu cita.")

    with s_col3:
        st.subheader("📊 Consultoría Fisiológica Integral")
        st.write("Seguimiento continuo de parámetros hemodinámicos, horario óptimo de entrenamiento adaptado a medicación y nutrición celular.")
        st.caption("Inversión: $90 USD / Plan Trimestral")
        if st.button("Reservar Consultoría", key="btn3"):
            st.info("Utiliza la pestaña 'Pasarela de Pagos' o contáctanos por WhatsApp para formalizar tu cita.")

# -----------------------------------------------------------------------------
# TAB 3: PASARELA DE PAGOS (Pago Móvil & ACH Xpress)
# -----------------------------------------------------------------------------
with tab3:
    st.header("Pasarela de Pagos Directa & Segura")
    st.write("Selecciona tu método de pago preferido para formalizar tus servicios desde cualquier lugar de Latinoamérica.")

    pay_col1, pay_col2 = st.columns(2)

    with pay_col1:
        st.markdown("""
            <div class="payment-box">
                <h3 style="color:#0A2540; margin-top:0;">🇻🇪 Pago Móvil (Venezuela)</h3>
                <p><b>Banco:</b> 0102 - Banco de Venezuela</p>
                <p><b>Teléfono:</b> 0412-0582064</p>
                <p><b>Cédula de Identidad:</b> V-15.642.481</p>
                <p><b>Titular:</b> Carlos Pacheco</p>
                <small style="color:#6C757D;">Transacción inmediata sin comisiones adicionales.</small>
            </div>
        """, unsafe_allow_html=True)

    with pay_col2:
        st.markdown("""
            <div class="payment-box">
                <h3 style="color:#0A2540; margin-top:0;">🌎 Transferencia Internacional ACH Xpress</h3>
                <p><b>Titular de la Cuenta:</b> Rosa Pacheco</p>
                <p><b>Número de Cuenta:</b> 3-002-05387291-65</p>
                <p><b>Tipo de Cuenta:</b> Cuenta de Ahorro</p>
                <p><b>Dominio Asociado:</b> pachecoazuajerosamaria.com</p>
                <small style="color:#6C757D;">Exclusivo para transacciones ACH Xpress internacionales.</small>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("📩 Confirmación de Pago")
    st.write("Una vez realizada tu transferencia o Pago Móvil, envía tu comprobante por correo o WhatsApp para activar tu plan:")
    st.write("📧 **Correo Oficial:** pachecohx@gmail.com")
    st.write("📱 **WhatsApp Soporte Directo:** +58 412-0582064")

# -----------------------------------------------------------------------------
# TAB 4: BLOG DE FISIOLOGÍA VASCULAR
# -----------------------------------------------------------------------------
with tab4:
    st.header("Blog de Fisiología Vascular & Longevidad")
    st.write("Artículos informativos para proteger tu sistema circulatorio y potenciar tu rendimiento de forma segura.")

    art1, art2 = st.columns(2)

    with art1:
        st.subheader("1. El Horario Óptimo para Entrenar si Padeces de Tensión Alta")
        st.write("Entrenar en horas de la tarde (6:00 pm a 7:00 pm) suele brindar la mejor cobertura farmacológica acumulada para prevenir picos hipertensivos de esfuerzo...")
        st.button("Leer Artículo Completo ->", key="art1_btn")

    with art2:
        st.subheader("2. Protección Endotelial: ¿Por qué la Intensidad Excesiva daña tus Arterias?")
        st.write("El estrés de cizallamiento en arterias con ateromas carotídeos puede desestabilizar la placa. Aprende a controlar la velocidad de onda de pulso...")
        st.button("Leer Artículo Completo ->", key="art2_btn")

# Footer
st.markdown("---")
st.markdown("<p style='text-align:center; color:#6C757D;'>© 2026 Lic. Carlos Pacheco - Todos los derechos reservados. Plataforma de Salud Vascular & Protección Endotelial.</p>", unsafe_allow_html=True)
