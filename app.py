import streamlit as st
from pydantic import BaseModel, Field
from typing import List
import os
from dotenv import load_dotenv
import pypdf
import docx
import zipfile
import io
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from google import genai

load_dotenv()

class CriterioEvaluado(BaseModel):
    criterio: str = Field(description="Descripción del criterio evaluado.")
    ponderacion: float = Field(description="Peso del criterio en porcentaje (Suma total = 100).")
    porcentaje_logrado: float = Field(description="Porcentaje de logro alcanzado (0.0 a 100.0).")
    puntaje_obtenido: float = Field(description="Puntaje obtenido (ponderacion * porcentaje_logrado / 100).")
    justificacion: str = Field(description="Explicación detallada basada en el código.")

class ReporteEvaluacion(BaseModel):
    nota_final: float = Field(description="Nota en escala chilena de 1.0 a 7.0 (Exigencia del 60% para el 4.0).")
    total_ponderado: float = Field(description="Porcentaje total sumando todos los puntajes obtenidos (0.0 a 100.0).")
    estado: str = Field(description="Indicar 'Aprobado' si la nota es >= 4.0, de lo contrario 'Reprobado'.")
    retroalimentacion_personalizada: str = Field(description="Párrafo extenso detallando el desempeño general y errores específicos.")
    evaluacion_detallada: List[CriterioEvaluado] = Field(description="Lista de criterios evaluados formando la tabla de rúbrica.")

def crear_pdf_bytes(reporte: ReporteEvaluacion) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    story = []
    
    styles = getSampleStyleSheet()
    
    style_title = ParagraphStyle(name='DocTitle', parent=styles['Heading1'], fontSize=18, leading=22, textColor=colors.HexColor('#0F172A'), spaceAfter=15)
    style_subtitle = ParagraphStyle(name='DocSubTitle', parent=styles['Heading2'], fontSize=13, leading=16, textColor=colors.HexColor('#1E3A8A'), spaceBefore=12, spaceAfter=8)
    style_body = ParagraphStyle(name='DocBody', parent=styles['Normal'], fontSize=10, leading=14, textColor=colors.HexColor('#334155'))
    style_table_header = ParagraphStyle(name='THeader', parent=styles['Normal'], fontSize=9, leading=11, textColor=colors.white, fontName='Helvetica-Bold')
    style_table_cell = ParagraphStyle(name='TCell', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor('#1E293B'))
    
    story.append(Paragraph("📋 INFORME DE EVALUACIÓN AUTOMÁTICA CON IA", style_title))
    story.append(Paragraph("<b>Sistema:</b> Krino Insight - Duoc UC (Clon Educativo)", style_body))
    story.append(Spacer(1, 15))
    
    resumen_data = [
        [
            Paragraph(f"<b>Nota Final:</b> {reporte.nota_final:.1f}", style_body),
            Paragraph(f"<b>Resultado Final:</b> {reporte.estado}", style_body),
            Paragraph(f"<b>Total Ponderado:</b> {reporte.total_ponderado}%", style_body)
        ]
    ]
    t_resumen = Table(resumen_data, colWidths=[170, 170, 170])
    t_resumen.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F1F5F9')),
        ('PADDING', (0,0), (-1,-1), 10),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
    ]))
    story.append(t_resumen)
    story.append(Spacer(1, 15))
    
    story.append(Paragraph("Retroalimentación Personalizada:", style_subtitle))
    story.append(Paragraph(reporte.retroalimentacion_personalizada, style_body))
    story.append(Spacer(1, 15))
    
    story.append(Paragraph("Evaluación Detallada por Criterios:", style_subtitle))
    
    tabla_pdf_datos = [[
        Paragraph("Criterio", style_table_header),
        Paragraph("Pond.", style_table_header),
        Paragraph("% Logro", style_table_header),
        Paragraph("Punt. Obtenida", style_table_header),
        Paragraph("Justificación", style_table_header)
    ]]
    
    for crit in reporte.evaluacion_detallada:
        tabla_pdf_datos.append([
            Paragraph(crit.criterio, style_table_cell),
            Paragraph(f"{crit.ponderacion}%", style_table_cell),
            Paragraph(f"{crit.porcentaje_logrado}%", style_table_cell),
            Paragraph(f"{crit.puntaje_obtenido}%", style_table_cell),
            Paragraph(crit.justificacion, style_table_cell)
        ])
        
    t_detalle = Table(tabla_pdf_datos, colWidths=[120, 45, 50, 50, 245])
    t_detalle.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 6),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
    ]))
    
    for i in range(1, len(tabla_pdf_datos)):
        if i % 2 == 0:
            t_detalle.setStyle(TableStyle([('BACKGROUND', (0, i), (-1, i), colors.HexColor('#F8FAFC'))]))
            
    story.append(t_detalle)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

api_key = st.secrets.get("GEMINI_API_KEY") if "GEMINI_API_KEY" in st.secrets else os.getenv("GEMINI_API_KEY")

if not api_key:
    st.warning("Por favor, configura tu GEMINI_API_KEY para poder iniciar la evaluación.")
else:
    try:
        client = genai.Client(api_key=api_key)
    except Exception as e:
        st.error(f"Error al inicializar Google GenAI: {e}")

def extraer_texto_requerimientos(archivo) -> str:
    nombre_archivo = archivo.name
    if nombre_archivo.endswith('.txt') or nombre_archivo.endswith('.md'):
        return archivo.read().decode("utf-8", errors="ignore")
    elif nombre_archivo.endswith('.pdf'):
        lector_pdf = pypdf.PdfReader(archivo)
        texto_completo = ""
        for pagina in lector_pdf.pages:
            texto_pagina = pagina.extract_text()
            if texto_pagina: 
                texto_completo += texto_pagina + "\n"
        return texto_completo
    elif nombre_archivo.endswith('.docx'):
        documento = docx.Document(archivo)
        texto_completo = ""
        for parrafo in documento.paragraphs:
            texto_completo += parrafo.text + "\n"
        return texto_completo
    return ""

def extraer_codigo_de_lista_py(lista_archivos) -> str:
    codigo_combinado = ""
    for archivo in lista_archivos:
        if archivo.name.endswith('.py'):
            contenido = archivo.read().decode("utf-8", errors="ignore")
            codigo_combinado += f"\n\n--- ARCHIVO: {archivo.name} ---\n{contenido}"
    return codigo_combinado

def extraer_codigo_de_zip(archivo_zip) -> str:
    codigo_combinado = ""
    with zipfile.ZipFile(io.BytesIO(archivo_zip.read())) as carpeta_zip:
        for nombre_interno in carpeta_zip.namelist():
            if nombre_interno.endswith('.py') and not nombre_interno.startswith('__MACOSX') and not os.path.basename(nombre_interno).startswith('.'):
                with carpeta_zip.open(nombre_interno) as archivo_interno:
                    contenido = archivo_interno.read().decode("utf-8", errors="ignore")
                    codigo_combinado += f"\n\n--- ARCHIVO: {nombre_interno} ---\n{contenido}"
    return codigo_combinado

def evaluar_proyecto_con_gemini(requerimientos_texto: str, codigo_texto: str) -> ReporteEvaluacion:
    prompt_sistema = (
        "Eres un Sistema de Evaluación Inteligente para una universidad chilena. "
        "Evalúa el código entregado contrastándolo estrictamente con los requerimientos. "
        "REGLAS DE CALIFICACIÓN: "
        "1. Genera una lista de criterios de evaluación basados en los requerimientos. La suma de las 'ponderaciones' de todos los criterios debe ser exactamente 100. "
        "2. Calcula el 'Total Ponderado' (0 a 100%) sumando los puntajes obtenidos. "
        "3. Convierte ese porcentaje a la Escala de Notas de Chile (1.0 a 7.0), considerando que un 60% de logro equivale a un 4.0. "
        "4. Redacta una 'Retroalimentación Personalizada' constructiva y profesional."
    )
    prompt_usuario = f"REQUERIMIENTOS A EVALUAR:\n{requerimientos_texto}\n\nCÓDIGO FUENTE COMPLETO DEL PROYECTO:\n{codigo_texto}"

    response = client.models.generate_content(
        model="gemini-1.5-pro",
        contents=prompt_usuario,
        config={
            "response_mime_type": "application/json",
            "response_schema": ReporteEvaluacion,
            "system_instruction": prompt_sistema,
            "temperature": 0.0
        }
    )
    return response.parsed

st.set_page_config(page_title="Sistema de Evaluación Inteligente", layout="wide")
st.title("🎓 Sistema de Evaluación Inteligente")
st.markdown("Sube la pauta de evaluación y el código del estudiante.")

col1, col2 = st.columns(2)
with col1:
    archivo_req = st.file_uploader("1. Pauta / Requerimientos (.txt, .pdf, .docx)", type=["txt", "md", "pdf", "docx"])
with col2:
    metodo_carga = st.radio("2. Código Fuente", ["Archivos .py sueltos", "Carpeta .zip"])
    texto_codigo = ""
    if metodo_carga == "Archivos .py sueltos":
        archivos_py = st.file_uploader("Selecciona archivos (.py)", type=["py"], accept_multiple_files=True)
        if archivos_py: 
            texto_codigo = extraer_codigo_de_lista_py(archivos_py)
    else:
        archivo_zip = st.file_uploader("Sube el archivo (.zip)", type=["zip"])
        if archivo_zip: 
            texto_codigo = extraer_codigo_de_zip(archivo_zip)

st.divider()

if st.button("🚀 Iniciar Revisión Automática", type="primary"):
    if 'client' in globals() and archivo_req and texto_codigo:
        texto_req = extraer_texto_requerimientos(archivo_req)
        
        if not texto_req.strip() or not texto_codigo.strip():
            st.error("Error al leer los archivos. Asegúrate de que no estén vacíos.")
        else:
            with st.spinner("Analizando código y construyendo rúbrica de evaluación..."):
                try:
                    reporte = evaluar_proyecto_con_gemini(texto_req, texto_codigo)
                    
                    st.success("Evaluación generada con éxito.")
                    
                    met1, met2, met3 = st.columns(3)
                    met1.metric(label="Nota Final", value=f"{reporte.nota_final:.1f}")
                    met2.metric(label="Resultado Final", value=reporte.estado)
                    met3.metric(label="Total Ponderado", value=f"{reporte.total_ponderado}%")
                    
                    with st.spinner("Preparando archivo PDF descagable..."):
                        pdf_data = crear_pdf_bytes(reporte)
                        
                    st.download_button(
                        label="📥 Descargar Reporte Completo en PDF",
                        data=pdf_data,
                        file_name=f"Retroalimentacion_Nota_{reporte.nota_final:.1f}.pdf",
                        mime="application/pdf"
                    )
                    
                    st.subheader("📝 Retroalimentación Personalizada")
                    st.info(reporte.retroalimentacion_personalizada)
                    
                    st.subheader("📊 Evaluación Detallada por Criterios")
                    tabla_datos = []
                    for crit in reporte.evaluacion_detallada:
                        tabla_datos.append({
                            "Criterio": crit.criterio,
                            "Ponderación": f"{crit.ponderacion}%",
                            "% Logrado": f"{crit.porcentaje_logrado}%",
                            "Punt. Obtenida": f"{crit.puntaje_obtenido}%",
                            "Justificación": crit.justificacion
                        })
                    st.table(tabla_datos)
                    
                except Exception as e:
                    st.error(f"Ocurrió un error en la evaluación: {e}")
    else:
        st.warning("Faltan archivos por subir o la API Key no está configurada correctamente.")
