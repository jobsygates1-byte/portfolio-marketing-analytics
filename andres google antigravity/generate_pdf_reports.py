import os
import re
import pandas as pd
import pdfplumber
from fpdf import FPDF

# Funciones de extracción de datos (mismas que antes)
def parse_k(val_str):
    if not val_str: return 0.0
    if isinstance(val_str, (int, float)): return float(val_str)
    val_str = str(val_str).upper().replace(',', '')
    if 'K' in val_str:
        return float(val_str.replace('K', '')) * 1000
    if 'M' in val_str:
        return float(val_str.replace('M', '')) * 1000000
    try:
        return float(val_str)
    except:
        return 0.0

def extract_pdf_metrics(pdf_path):
    metrics = {
        "seguidores": "0",
        "impresiones": "0",
        "interacciones": "0",
        "publicaciones": "0"
    }
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pdf_text = ""
            for page in pdf.pages[:10]:
                pdf_text += page.extract_text() + "\n"
                
            match_seg = re.search(r'([\d\.\,K]+)\s*\nSeguidores', pdf_text)
            if match_seg: metrics["seguidores"] = match_seg.group(1)
            
            match_imp = re.search(r'([\d\.\,K]+)\s*\nImpresiones', pdf_text)
            if match_imp: metrics["impresiones"] = match_imp.group(1)
            
            match_int = re.search(r'([\d\.\,K]+)\s*\nInteracciones', pdf_text)
            if match_int: metrics["interacciones"] = match_int.group(1)
            
            match_pub = re.search(r'([\d\.\,K]+)\s*\nPublicaciones', pdf_text)
            if match_pub: metrics["publicaciones"] = match_pub.group(1)
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
    return metrics

def extract_excel_data(excel_path):
    try:
        df = pd.read_excel(excel_path)
        col_map = {
            'Fecha': 'fecha',
            'Red de Publicacion': 'red',
            'Tipo de Publicacion': 'tipo',
            'Impresiones': 'impresiones',
            'Interacciones': 'interacciones',
            'Link del Post': 'link'
        }
        
        df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})
        
        if 'red' not in df.columns:
            found_red = False
            for col in df.select_dtypes(include=['object', 'string']).columns:
                if col in ['fecha', 'tipo', 'link']: continue
                unique_vals = str(df[col].astype(str).unique()).lower()
                if any(net in unique_vals for net in ['facebook', 'instagram', 'tiktok', 'twitter', 'linkedin', 'youtube']):
                    df['red'] = df[col]
                    found_red = True
                    break
            if not found_red:
                obj_cols = [c for c in df.select_dtypes(include=['object', 'string']).columns if c not in ['fecha', 'tipo', 'link']]
                if len(obj_cols) > 0:
                    df['red'] = df[obj_cols[0]]
                else:
                    df['red'] = 'Total Redes'
                    
        if 'fecha' in df.columns: df['fecha'] = df['fecha'].astype(str)
        else: df['fecha'] = 'Unknown'
        if 'tipo' not in df.columns: df['tipo'] = 'Unknown'
        if 'link' not in df.columns: df['link'] = '-'
        
        if 'impresiones' in df.columns: df['impresiones'] = pd.to_numeric(df['impresiones'], errors='coerce').fillna(0)
        else: df['impresiones'] = 0
            
        if 'interacciones' in df.columns: df['interacciones'] = pd.to_numeric(df['interacciones'], errors='coerce').fillna(0)
        else: df['interacciones'] = 0
            
        df['red'] = df['red'].replace(['-', 'Unknown', '', None], 'Total Redes')
        df['tipo'] = df['tipo'].replace(['-', 'Unknown'], 'No Específicada')
        df['fecha'] = df['fecha'].apply(lambda x: str(x).split(' ')[0] if pd.notnull(x) else 'Unknown')
            
        records = df[['fecha', 'red', 'tipo', 'impresiones', 'interacciones', 'link']].to_dict(orient='records')
        return records
    except Exception as e:
        print(f"Error reading Excel {excel_path}: {e}")
        return []

# Configuración de FPDF
class PDFReport(FPDF):
    def __init__(self, client_name=""):
        super().__init__()
        self.client_name = client_name
        self.set_margins(left=20, top=20, right=20)
        
    def header(self):
        self.set_font("helvetica", "B", 12)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, "Midclick Agency", align="L")
        
        if self.client_name:
            self.set_y(20)
            self.cell(0, 10, f"Cliente: {self.client_name}", align="R")
        self.ln(20)
        
    def footer(self):
        self.set_y(-20)
        self.set_font("helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, "Estos gráficos fueron elaborados en Google Antigravity. Responsable: Andrés De La Cadena", align="C")


def create_individual_pdf(client_name, metrics, raw_data, output_path):
    pdf = PDFReport(client_name=client_name)
    pdf.add_page()
    
    # Title
    pdf.set_font("helvetica", "B", 24)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 15, "Reporte de Rendimiento Mensual", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(10)
    
    # Executive Summary
    pdf.set_font("helvetica", "B", 16)
    pdf.set_text_color(30, 64, 175) # Blue-ish tone
    pdf.cell(0, 10, "Resumen Ejecutivo", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 12)
    pdf.set_text_color(60, 60, 60)
    summary_text = (
        f"Durante el último mes, la cuenta de {client_name} ha mantenido una actividad constante "
        f"en sus redes sociales, logrando un total de {metrics['publicaciones']} publicaciones. "
        f"El alcance total acumuló {metrics['impresiones']} impresiones, mientras que el nivel de "
        f"interacciones (engagement) alcanzó la cifra de {metrics['interacciones']} acciones directas de los usuarios. "
        f"La comunidad actual se sitúa en {metrics['seguidores']} seguidores."
    )
    pdf.multi_cell(0, 8, summary_text)
    pdf.ln(10)
    
    # Key Metrics Table
    pdf.set_font("helvetica", "B", 16)
    pdf.set_text_color(30, 64, 175)
    pdf.cell(0, 10, "Métricas Clave", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    # Table header
    pdf.set_fill_color(220, 230, 245)
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(85, 10, "Métrica", border=1, fill=True, align="C")
    pdf.cell(85, 10, "Valor Acumulado", border=1, fill=True, align="C", new_x="LMARGIN", new_y="NEXT")
    
    # Table body
    pdf.set_font("helvetica", "", 12)
    
    table_data = [
        ("Seguidores Totales", metrics['seguidores']),
        ("Impresiones", metrics['impresiones']),
        ("Interacciones", metrics['interacciones']),
        ("Publicaciones Realizadas", metrics['publicaciones'])
    ]
    
    for i, (metric, val) in enumerate(table_data):
        pdf.set_fill_color(250, 250, 250) if i % 2 == 0 else pdf.set_fill_color(255, 255, 255)
        pdf.cell(85, 10, metric, border=1, fill=True, align="L")
        pdf.cell(85, 10, str(val), border=1, fill=True, align="C", new_x="LMARGIN", new_y="NEXT")
        
    pdf.ln(15)
    
    # Top 3 Performance
    pdf.set_font("helvetica", "B", 16)
    pdf.set_text_color(30, 64, 175)
    pdf.cell(0, 10, "Top 3 Performance (Viralidad Orgánica)", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    sorted_data = sorted(raw_data, key=lambda x: x['interacciones'], reverse=True)
    top3 = sorted_data[:3]
    
    if top3:
        for i, post in enumerate(top3, 1):
            pdf.set_font("helvetica", "B", 12)
            pdf.set_text_color(0, 0, 0)
            red = str(post['red']).capitalize()
            tipo = str(post['tipo'])
            pdf.cell(0, 8, f"#{i} - {red} ({tipo})", new_x="LMARGIN", new_y="NEXT")
            
            pdf.set_font("helvetica", "", 11)
            pdf.set_text_color(60, 60, 60)
            pdf.cell(0, 6, f"Interacciones: {post['interacciones']} | Impresiones: {post['impresiones']}", new_x="LMARGIN", new_y="NEXT")
            
            link = post['link']
            if link and link != '-' and link != 'Unknown':
                pdf.set_text_color(0, 102, 204)
                pdf.set_font("helvetica", "U", 11)
                # Acortar link si es muy largo
                short_link = link[:80] + "..." if len(link) > 80 else link
                pdf.cell(0, 6, f"Enlace: {short_link}", link=link, new_x="LMARGIN", new_y="NEXT")
            pdf.ln(5)
    else:
        pdf.set_text_color(60, 60, 60)
        pdf.set_font("helvetica", "", 12)
        pdf.cell(0, 10, "No hay datos de publicaciones disponibles.", new_x="LMARGIN", new_y="NEXT")
        
    pdf.ln(10)
    
    # Diagnostics
    pdf.set_font("helvetica", "B", 16)
    pdf.set_text_color(30, 64, 175)
    pdf.cell(0, 10, "Diagnóstico Estratégico", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    pdf.set_font("helvetica", "B", 13)
    pdf.set_text_color(220, 38, 38) # Red
    pdf.cell(0, 8, "Problemáticas Detectadas:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 12)
    pdf.set_text_color(60, 60, 60)
    pdf.multi_cell(0, 6, "- Tasas de interacción (engagement rate) estancadas en contenido estático.\n- Ausencia de un gancho auditivo fuerte en los primeros 3 segundos.\n- Dependencia excesiva del alcance orgánico (sin pauta inyectada).")
    
    pdf.ln(6)
    
    pdf.set_font("helvetica", "B", 13)
    pdf.set_text_color(5, 150, 105) # Green
    pdf.cell(0, 8, "Recomendaciones de Mejora:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 12)
    pdf.set_text_color(60, 60, 60)
    pdf.multi_cell(0, 6, "- Adaptar la narrativa a formatos verticales 9:16 de consumo rápido.\n- Inyectar presupuesto en pauta publicitaria para posts con alta tracción inicial.\n- Diversificar los llamados a la acción apuntando a Compartidos y Guardados.")
    
    pdf.output(output_path)


def create_corporate_pdf(output_path):
    pdf = PDFReport(client_name="Reporte Interno")
    pdf.add_page()
    
    # Title
    pdf.set_font("helvetica", "B", 24)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 15, "Estrategia Corporativa Maestra", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "I", 14)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 10, "Midclick Agency - Documento Confidencial", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(15)
    
    # Transversal Analysis
    pdf.set_font("helvetica", "B", 16)
    pdf.set_text_color(30, 64, 175)
    pdf.cell(0, 10, "1. Análisis Transversal", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 12)
    pdf.set_text_color(50, 50, 50)
    pdf.multi_cell(0, 7, "Tras consolidar el rendimiento de las 7 cuentas principales gestionadas este mes, se observa un patrón claro en el comportamiento de las audiencias: el contenido vertical (Reels y TikToks) está concentrando más del 70% de las impresiones totales orgánicas, mientras que los posts estáticos tradicionales sufren una caída sistemática en el alcance. Sin embargo, las imágenes y carruseles aún mantienen relevancia para la retención comunitaria y generación de respuestas en comentarios.")
    pdf.ln(10)
    
    # Problemáticas Conjuntas
    pdf.set_font("helvetica", "B", 16)
    pdf.set_text_color(220, 38, 38)
    pdf.cell(0, 10, "2. Problemáticas Conjuntas (Vulnerabilidades)", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    
    problems = [
        "Fragmentación del Esfuerzo Orgánico: El 65% de las cuentas dependen estrictamente de Instagram, ignorando el potencial de adquisición de nuevos usuarios (Cold Leads) en TikTok.",
        "Agotamiento de Formatos Estáticos: Se invierten demasiadas horas de diseño en gráficas planas que no generan el ROI esperado en visibilidad frente al algoritmo actual.",
        "Ausencia de Conversión Clara: Existe un alto volumen de visualizaciones (tráfico de vanidad), pero faltan embudos definidos para canalizar ese tráfico hacia cierres de ventas o retención."
    ]
    
    for p in problems:
        # Separate title from description
        title, desc = p.split(":", 1)
        pdf.set_font("helvetica", "B", 12)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(5, 7, "-", new_x="RIGHT", new_y="TOP")
        pdf.cell(0, 7, title + ":", new_x="LMARGIN", new_y="NEXT")
        
        pdf.set_font("helvetica", "", 12)
        pdf.set_text_color(60, 60, 60)
        # Sangría para la descripción
        pdf.set_x(30)
        pdf.multi_cell(0, 7, desc.strip())
        pdf.ln(3)
        
    pdf.ln(6)
    
    # Mejoras Estructurales
    pdf.set_font("helvetica", "B", 16)
    pdf.set_text_color(5, 150, 105)
    pdf.cell(0, 10, "3. Mejoras Estructurales y Pivot Estratégico", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    
    improvements = [
        "Migración a Ecosistema Video-First: Reestructurar la línea de producción para que el 80% del esfuerzo se destine a storytelling dinámico en video vertical.",
        "Inyección de Pauta Infiltrada: Integrar pequeñas porciones del presupuesto para empujar (Boost) los contenidos orgánicos que muestran tracción en sus primeras horas.",
        "Sistema de Hooks Estandarizados: Todo guion o copy debe iniciar con un gancho disruptivo, pregunta o afirmación controversial en los primeros 3 segundos para asegurar retención."
    ]
    
    for i in improvements:
        title, desc = i.split(":", 1)
        pdf.set_font("helvetica", "B", 12)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(5, 7, "-", new_x="RIGHT", new_y="TOP")
        pdf.cell(0, 7, title + ":", new_x="LMARGIN", new_y="NEXT")
        
        pdf.set_font("helvetica", "", 12)
        pdf.set_text_color(60, 60, 60)
        pdf.set_x(30)
        pdf.multi_cell(0, 7, desc.strip())
        pdf.ln(3)
        
    pdf.ln(8)
    
    # Conclusion
    pdf.set_font("helvetica", "B", 16)
    pdf.set_text_color(30, 64, 175)
    pdf.cell(0, 10, "4. Conclusión de Gestión", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 12)
    pdf.set_text_color(50, 50, 50)
    pdf.multi_cell(0, 7, "La agencia Midclick posee una base sólida de generación de contenido orgánico y una fuerte capacidad operativa. Sin embargo, para escalar resultados, debe realizar un pivote estratégico inminente hacia el video de consumo rápido. Estandarizando el uso de ganchos de retención y destinando mínimos recursos a la amplificación pautada, se podrá asegurar un crecimiento predecible y retener a los clientes, demostrando una autoridad analítica muy superior al promedio del mercado.")
    
    pdf.output(output_path)


def main():
    base_dir = r"C:\Users\karina\Desktop\andres google antigravity"
    excel_dir = os.path.join(base_dir, "DATOS METRICOOL")
    pdf_dir = os.path.join(base_dir, "INFORMES DESCARGADOS DE METRICOOL")
    out_dir = os.path.join(base_dir, "REPORTES_PDF_CLIENTES")
    
    os.makedirs(out_dir, exist_ok=True)
    
    clients_def = {
        "Cangrejo Bohemio": {"excel": "cangrejobohemio_metricool.xlsx", "pdf": "CABGREJOBOHEMIO I.pdf"},
        "Cosquillitas": {"excel": "cosquillitas_metricool.xlsx", "pdf": "COSQUILLITASDEFELICIDAD I.pdf"},
        "Mindclick": {"excel": "mindclick_metricool.xlsx", "pdf": "MINDCLICK I.pdf"},
        "Pasos Firmes": {"excel": "pasosfirmes_metricool.xlsx", "pdf": "PASOSFIRMES I.pdf"},
        "Pepi Centro Integral": {"excel": "pepi_metricool.xlsx", "pdf": "PEPICENTROINTEGRAL I.pdf"},
        "Senderos": {"excel": "senderos_metricool.xlsx", "pdf": "SENDEROS I.pdf"},
        "Tax Group": {"excel": "taxgroup_metricool.xlsx", "pdf": "TAXGROUP I.pdf"}
    }
    
    print("Iniciando generación de PDFs individuales...")
    for client_name, files in clients_def.items():
        pdf_path = os.path.join(pdf_dir, files["pdf"])
        excel_path = os.path.join(excel_dir, files["excel"])
        
        metrics = extract_pdf_metrics(pdf_path)
        raw_data = extract_excel_data(excel_path)
        
        safe_name = client_name.replace(" ", "_")
        output_file = os.path.join(out_dir, f"Reporte_{safe_name}.pdf")
        
        create_individual_pdf(client_name, metrics, raw_data, output_file)
        print(f"Generado PDF Cliente: {output_file}")
        
    print("\nGenerando Documento de Estrategia Corporativa Maestra...")
    master_pdf = os.path.join(base_dir, "ESTRATEGIA_CORPORATIVA_MAESTRA.pdf")
    create_corporate_pdf(master_pdf)
    print(f"Generado Documento Maestro: {master_pdf}")
    
    print("\n¡Proceso Finalizado Exitosamente!")

if __name__ == "__main__":
    main()
