import pandas as pd
import pdfplumber

def check_data():
    excel_path = r"C:\Users\karina\Desktop\andres google antigravity\DATOS METRICOOL\cangrejobohemio_metricool.xlsx"
    df = pd.read_excel(excel_path)
    print("=== EXCEL COLUMNS ===")
    print(df.columns.tolist())
    print("=== EXCEL DATA ===")
    print(df.head())
    
    pdf_path = r"C:\Users\karina\Desktop\andres google antigravity\INFORMES DESCARGADOS DE METRICOOL\CABGREJOBOHEMIO I.pdf"
    print("\n=== PDF DATA ===")
    full_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            full_text.append(page.extract_text())
    
    with open("pdf_text.txt", "w", encoding="utf-8") as f:
        f.write("\n\n--- PAGE BREAK ---\n\n".join(full_text))
    print("Saved to pdf_text.txt")

if __name__ == '__main__':
    check_data()
