import os
import re
import pandas as pd
import pdfplumber

def parse_k(val_str):
    if not val_str: return 0.0
    val_str = val_str.upper().replace(',', '')
    if 'K' in val_str:
        return float(val_str.replace('K', '')) * 1000
    if 'M' in val_str:
        return float(val_str.replace('M', '')) * 1000000
    try:
        return float(val_str)
    except:
        return 0.0

def is_close(val1, val2, tolerance=0.05):
    if val1 == 0 and val2 == 0: return True
    return abs(val1 - val2) / max(abs(val1), abs(val2)) <= tolerance

def main():
    base_dir = r"C:\Users\karina\Desktop\andres google antigravity"
    excel_dir = os.path.join(base_dir, "DATOS METRICOOL")
    pdf_dir = os.path.join(base_dir, "INFORMES DESCARGADOS DE METRICOOL")
    
    clients = {
        "Cangrejo Bohemio": {"excel": "cangrejobohemio_metricool.xlsx", "pdf": "CABGREJOBOHEMIO I.pdf", "logo": "CANGREJO BOHEMIO LOGO.png"},
        "Cosquillitas": {"excel": "cosquillitas_metricool.xlsx", "pdf": "COSQUILLITASDEFELICIDAD I.pdf", "logo": "COSQUILLITAS LOGO.png"},
        "Mindclick": {"excel": "mindclick_metricool.xlsx", "pdf": "MINDCLICK I.pdf", "logo": "MINDCLICK LOGO.png"},
        "Pasos Firmes": {"excel": "pasosfirmes_metricool.xlsx", "pdf": "PASOSFIRMES I.pdf", "logo": "PASOS FIRMES LOGO.png"},
        "Pepi Centro Integral": {"excel": "pepi_metricool.xlsx", "pdf": "PEPICENTROINTEGRAL I.pdf", "logo": "PEPI LOGO.png"},
        "Senderos": {"excel": "senderos_metricool.xlsx", "pdf": "SENDEROS I.pdf", "logo": "SENDEROS LOGO.png"},
        "Tax Group": {"excel": "taxgroup_metricool.xlsx", "pdf": "TAXGROUP I.pdf", "logo": "TAX GROUP LOGO.png"}
    }
    
    report_lines = []
    report_lines.append("# Reporte Consolidado de Clientes")
    report_lines.append("")
    report_lines.append("| Logo | Cliente | Impresiones PDF | Impresiones Excel | Interacciones PDF | Interacciones Excel | Publicaciones PDF | Publicaciones Excel | ¿Coincide? |")
    report_lines.append("|---|---|---|---|---|---|---|---|---|")
    
    for client_name, files in clients.items():
        pdf_path = os.path.join(pdf_dir, files["pdf"])
        excel_path = os.path.join(excel_dir, files["excel"])
        
        # 1. Read PDF
        pdf_impresiones = "N/A"
        pdf_interacciones = "N/A"
        pdf_publicaciones = "N/A"
        
        with pdfplumber.open(pdf_path) as pdf:
            pdf_text = ""
            # Usually metrics are in the first 10 pages
            for page in pdf.pages[:10]:
                pdf_text += page.extract_text() + "\n"
                
            match_imp = re.search(r'([\d\.\,K]+)\s*\nImpresiones', pdf_text)
            if match_imp: pdf_impresiones = match_imp.group(1)
            
            match_int = re.search(r'([\d\.\,K]+)\s*\nInteracciones', pdf_text)
            if match_int: pdf_interacciones = match_int.group(1)
            
            match_pub = re.search(r'([\d\.\,K]+)\s*\nPublicaciones', pdf_text)
            if match_pub: pdf_publicaciones = match_pub.group(1)
            
        # 2. Read Excel
        df = pd.read_excel(excel_path)
        excel_impresiones = pd.to_numeric(df['Impresiones'], errors='coerce').sum() if 'Impresiones' in df.columns else 0
        excel_interacciones = pd.to_numeric(df['Interacciones'], errors='coerce').sum() if 'Interacciones' in df.columns else 0
        excel_publicaciones = len(df)
        
        # 3. Compare
        pdf_imp_val = parse_k(pdf_impresiones) if pdf_impresiones != "N/A" else 0
        pdf_int_val = parse_k(pdf_interacciones) if pdf_interacciones != "N/A" else 0
        pdf_pub_val = parse_k(pdf_publicaciones) if pdf_publicaciones != "N/A" else 0
        
        match_imp_bool = is_close(pdf_imp_val, excel_impresiones, 0.05)
        match_int_bool = is_close(pdf_int_val, excel_interacciones, 0.05)
        match_pub_bool = is_close(pdf_pub_val, excel_publicaciones, 0.05)
        
        overall_match = match_imp_bool and match_int_bool and match_pub_bool
        status = "✅ Sí" if overall_match else "❌ No"
        
        logo_path = os.path.join("LOGOS CLIENTES", files["logo"]).replace("\\", "/")
        logo_img = f"<img src='{logo_path}' width='50' />"
        
        report_lines.append(f"| {logo_img} | {client_name} | {pdf_impresiones} | {excel_impresiones} | {pdf_interacciones} | {excel_interacciones} | {pdf_publicaciones} | {excel_publicaciones} | {status} |")
        
    report_path = os.path.join(base_dir, "REPORTE_CONSOLIDADO.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))
        
    print("Report generated successfully.")

if __name__ == '__main__':
    main()
