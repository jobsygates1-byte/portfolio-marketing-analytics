import os
import re
import json
import base64
import pandas as pd
import pdfplumber

def image_to_base64(filepath):
    try:
        with open(filepath, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode("utf-8")
        ext = filepath.split('.')[-1].lower()
        mime_type = f"image/{ext}" if ext in ['png', 'jpeg', 'jpg', 'gif'] else "image/png"
        return f"data:{mime_type};base64,{encoded_string}"
    except Exception as e:
        print(f"Error encoding image {filepath}: {e}")
        return ""

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

def main():
    base_dir = r"C:\Users\karina\Desktop\andres google antigravity"
    excel_dir = os.path.join(base_dir, "DATOS METRICOOL")
    pdf_dir = os.path.join(base_dir, "INFORMES DESCARGADOS DE METRICOOL")
    
    clients_def = {
        "Cangrejo Bohemio": {"excel": "cangrejobohemio_metricool.xlsx", "pdf": "CABGREJOBOHEMIO I.pdf", "logo": "CANGREJO BOHEMIO LOGO.png"},
        "Cosquillitas": {"excel": "cosquillitas_metricool.xlsx", "pdf": "COSQUILLITASDEFELICIDAD I.pdf", "logo": "COSQUILLITAS LOGO.png"},
        "Mindclick": {"excel": "mindclick_metricool.xlsx", "pdf": "MINDCLICK I.pdf", "logo": "MINDCLICK LOGO.png"},
        "Pasos Firmes": {"excel": "pasosfirmes_metricool.xlsx", "pdf": "PASOSFIRMES I.pdf", "logo": "PASOS FIRMES LOGO.png"},
        "Pepi Centro Integral": {"excel": "pepi_metricool.xlsx", "pdf": "PEPICENTROINTEGRAL I.pdf", "logo": "PEPI LOGO.png"},
        "Senderos": {"excel": "senderos_metricool.xlsx", "pdf": "SENDEROS I.pdf", "logo": "SENDEROS LOGO.png"},
        "Tax Group": {"excel": "taxgroup_metricool.xlsx", "pdf": "TAXGROUP I.pdf", "logo": "TAX GROUP LOGO.png"}
    }
    
    dashboard_data = {}
    
    print("Extrayendo datos de clientes para V3...")
    for client_name, files in clients_def.items():
        pdf_path = os.path.join(pdf_dir, files["pdf"])
        excel_path = os.path.join(excel_dir, files["excel"])
        
        pdf_metrics = extract_pdf_metrics(pdf_path)
        excel_data = extract_excel_data(excel_path)
        
        client_logo_path = os.path.join(base_dir, "LOGOS CLIENTES", files['logo'])
        dashboard_data[client_name] = {
            "logo": image_to_base64(client_logo_path),
            "narrative": pdf_metrics,
            "raw_data": excel_data,
            "is_strategy": False
        }

    # Add Strategy Client
    brain_svg = "data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='white'><path d='M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 17.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.39z'/></svg>"
    
    dashboard_data["Estrategia Corporativa"] = {
        "logo": brain_svg,
        "is_strategy": True
    }

    json_data = json.dumps(dashboard_data, ensure_ascii=False)
    
    html_template = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Plataforma BI Analítica V3</title>
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/particles.js/2.0.0/particles.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-dark: #020617;
            --text-light: #f8fafc;
            --text-muted: #94a3b8;
            --card-bg: rgba(30, 41, 59, 0.7);
            --card-border: rgba(255, 255, 255, 0.1);
            --sidebar-width: 280px;
            --sidebar-visible: 20px;
            --primary: #38bdf8;
            --primary-hover: #0ea5e9;
            --danger: #f43f5e;
            --success: #10b981;
        }
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 'Inter', sans-serif;
        }
        
        body {
            background-color: var(--bg-dark);
            color: var(--text-light);
            overflow-x: hidden;
            min-height: 100vh;
        }

        #particles-js {
            position: fixed;
            width: 100%;
            height: 100%;
            z-index: -1;
            top: 0;
            left: 0;
        }
        
        /* Watermark */
        .watermark {
            position: fixed;
            bottom: 10px;
            left: 10px;
            font-size: 11px;
            color: rgba(255,255,255,0.3);
            z-index: 1000;
            pointer-events: none;
        }
        
        /* Logos */
        .header-logos {
            position: fixed;
            top: 20px;
            left: 0;
            width: 100%;
            height: 60px;
            pointer-events: none;
            z-index: 50;
            display: flex;
            justify-content: space-between;
            padding: 0 40px;
        }
        .header-logos img {
            height: 100%;
            object-fit: contain;
            filter: drop-shadow(0 0 10px rgba(255,255,255,0.2));
        }
        
        /* Sidebar */
        #sidebar {
            position: fixed;
            top: 0;
            left: 0;
            width: var(--sidebar-width);
            height: 100vh;
            background: rgba(15, 23, 42, 0.85);
            backdrop-filter: blur(15px);
            box-shadow: 5px 0 25px rgba(0,0,0,0.5);
            transform: translateX(calc(-100% + var(--sidebar-visible)));
            transition: transform 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            z-index: 100;
            padding-top: 80px;
            border-right: 1px solid var(--card-border);
        }
        
        #sidebar:hover {
            transform: translateX(0);
        }
        
        .sidebar-hover-hint {
            position: absolute;
            top: 50%;
            right: 2px;
            transform: translateY(-50%);
            writing-mode: vertical-rl;
            font-size: 10px;
            color: var(--primary);
            letter-spacing: 2px;
            pointer-events: none;
        }
        
        #sidebar:hover .sidebar-hover-hint {
            opacity: 0;
        }
        
        .client-item {
            display: flex;
            align-items: center;
            padding: 12px 20px;
            cursor: pointer;
            transition: all 0.2s;
            border-left: 3px solid transparent;
            color: var(--text-muted);
        }
        
        .client-item:hover, .client-item.active {
            background: rgba(255,255,255,0.05);
            border-left-color: var(--primary);
            color: var(--text-light);
        }
        
        .client-item img {
            width: 30px;
            height: 30px;
            border-radius: 50%;
            object-fit: cover;
            margin-right: 15px;
            border: 1px solid rgba(255,255,255,0.2);
            background: #fff;
        }
        
        .client-item span {
            font-weight: 500;
            font-size: 14px;
        }

        /* Master Icon Styling */
        .client-item.strategy-btn {
            margin-top: 20px;
            border-top: 1px solid var(--card-border);
            padding-top: 20px;
        }
        .client-item.strategy-btn img {
            background: var(--danger);
            padding: 4px;
        }
        .client-item.strategy-btn span {
            color: var(--danger);
            font-weight: 700;
        }
        
        /* Main Content */
        #main-content {
            padding: 100px 40px 40px calc(var(--sidebar-visible) + 40px);
            transition: padding-left 0.4s;
            display: flex;
            flex-direction: column;
            gap: 30px;
            position: relative;
            z-index: 10;
        }
        
        .card {
            background: var(--card-bg);
            border-radius: 12px;
            padding: 24px;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
            backdrop-filter: blur(12px);
            border: 1px solid var(--card-border);
        }

        /* Narrative blocks */
        .narrative-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-top: 15px;
        }
        
        .metric-box {
            text-align: center;
            padding: 15px;
            background: rgba(0,0,0,0.2);
            border-radius: 8px;
            border: 1px solid rgba(255,255,255,0.05);
        }
        
        .metric-value {
            font-size: 28px;
            font-weight: 700;
            color: var(--primary);
            text-shadow: 0 0 10px rgba(56, 189, 248, 0.3);
        }
        
        .metric-label {
            font-size: 13px;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-top: 5px;
        }

        /* Interpretive Storytelling Layer */
        .story-layer {
            margin-top: 15px;
            padding: 15px;
            border-left: 4px solid var(--primary);
            background: rgba(56, 189, 248, 0.05);
            font-size: 14px;
            line-height: 1.6;
            color: #e2e8f0;
        }
        .story-layer strong { color: var(--primary); }
        .story-layer.combo { border-left-color: #8b5cf6; background: rgba(139, 92, 246, 0.05); }
        .story-layer.combo strong { color: #8b5cf6; }

        /* BI Charts */
        .charts-container {
            display: grid;
            grid-template-columns: 1fr 1fr;
            grid-template-rows: auto auto;
            gap: 20px;
        }
        .chart-timeseries { grid-column: 1 / -1; }
        .chart-wrapper { width: 100%; height: 400px; }

        /* Viralidad (Top 3) */
        .top-posts-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin-top: 15px;
        }
        .top-post-card {
            background: linear-gradient(145deg, rgba(30,41,59,0.9), rgba(15,23,42,0.9));
            border: 1px solid var(--card-border);
            border-radius: 8px;
            padding: 20px;
            position: relative;
            overflow: hidden;
        }
        .top-post-card::before {
            content: '';
            position: absolute;
            top: 0; left: 0; width: 4px; height: 100%;
            background: var(--primary);
        }
        .rank-badge {
            position: absolute;
            top: 10px; right: 10px;
            background: rgba(56, 189, 248, 0.2);
            color: var(--primary);
            padding: 4px 10px;
            border-radius: 20px;
            font-weight: 700;
            font-size: 12px;
        }
        .post-title { font-size: 16px; font-weight: 600; margin-bottom: 5px; color: #fff; }
        .post-meta { font-size: 13px; color: var(--text-muted); margin-bottom: 15px; }
        .btn-link {
            display: inline-block;
            padding: 8px 16px;
            background: rgba(255,255,255,0.1);
            color: #fff;
            text-decoration: none;
            border-radius: 6px;
            font-size: 13px;
            transition: all 0.2s;
            border: 1px solid rgba(255,255,255,0.2);
        }
        .btn-link:hover { background: var(--primary); border-color: var(--primary); }

        /* Strategy Analysis Dual Layout */
        .strategy-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-top: 15px;
        }
        .strategy-box {
            padding: 20px;
            border-radius: 8px;
            border: 1px solid var(--card-border);
            background: rgba(0,0,0,0.2);
        }
        .strategy-box h3 { margin-bottom: 15px; display: flex; align-items: center; gap: 10px; font-size: 16px; }
        .strategy-box.problems h3 { color: var(--danger); }
        .strategy-box.improvements h3 { color: var(--success); }
        .strategy-box ul { list-style: none; }
        .strategy-box li {
            margin-bottom: 10px;
            font-size: 14px;
            color: #e2e8f0;
            padding-left: 20px;
            position: relative;
            line-height: 1.5;
        }
        .strategy-box.problems li::before { content: '✕'; position: absolute; left: 0; color: var(--danger); font-weight: bold;}
        .strategy-box.improvements li::before { content: '✓'; position: absolute; left: 0; color: var(--success); font-weight: bold;}

        /* Exclusive Corporate Strategy View */
        #strategy-view {
            max-width: 900px;
            margin: 0 auto;
        }
        .corp-header {
            text-align: center;
            margin-bottom: 40px;
        }
        .corp-header h1 { color: var(--danger); font-size: 32px; margin-bottom: 10px; }
        .corp-header p { color: var(--text-muted); font-size: 16px; }

        #empty-state {
            text-align: center;
            margin-top: 20vh;
            color: var(--text-muted);
        }
    </style>
</head>
<body>

    <div id="particles-js"></div>
    <div class="watermark">Estos gráficos fueron elaborados en Google Antigravity con HTML. Responsable: Andrés De La Cadena.</div>

    <div class="header-logos">
        <img src="__AGENCY_LOGO__" alt="Agency Logo" id="agency-logo">
        <img src="" alt="Client Logo" id="header-client-logo" style="display: none;">
    </div>

    <div id="sidebar">
        <div class="sidebar-hover-hint">NAVEGACIÓN</div>
        <div id="client-list"></div>
    </div>

    <div id="main-content">
        <div id="empty-state">
            <h2>Bienvenido a la Plataforma BI de Midclick</h2>
            <p>Desplaza el cursor sobre el borde izquierdo para seleccionar un cliente o analizar la Estrategia Corporativa.</p>
        </div>
        
        <!-- Vista Específica de Cliente -->
        <div id="dashboard-view" style="display: none;">
            
            <div class="card">
                <h2 id="client-title" style="margin-bottom: 10px; color: #fff;">Nombre Cliente</h2>
                <p style="color: var(--text-muted); font-size: 15px;">Resumen del rendimiento cualitativo analizado en el periodo actual:</p>
                <div class="narrative-grid">
                    <div class="metric-box">
                        <div class="metric-value" id="val-seguidores">-</div>
                        <div class="metric-label">Seguidores</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-value" id="val-impresiones">-</div>
                        <div class="metric-label">Impresiones</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-value" id="val-interacciones">-</div>
                        <div class="metric-label">Interacciones</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-value" id="val-publicaciones">-</div>
                        <div class="metric-label">Publicaciones</div>
                    </div>
                </div>
            </div>
            
            <div class="charts-container">
                <div class="card chart-timeseries">
                    <div id="chart-time" class="chart-wrapper"></div>
                    <div class="story-layer">
                        <strong>Deducción Temporal:</strong> Las curvas de arriba revelan los picos de interés a lo largo del mes. Un desacoplamiento (muchas impresiones pero nulas interacciones) en días pico advierte que el contenido alcanzó a las masas pero falló en el enganche (call-to-action). <br>
                        <em>💡 Interactividad: Haz clic en cualquier barra inferior para ver cómo la línea de tiempo se reajusta únicamente para esa red específica.</em>
                    </div>
                </div>
                
                <div class="card">
                    <div id="chart-network" class="chart-wrapper"></div>
                    <div class="story-layer">
                        <strong>Deducción de Canales:</strong> Identifica el terreno donde tu comunidad reside. Si una red domina en impresiones en este gráfico de barras, significa que el algoritmo orgánico la está premiando.
                    </div>
                </div>
                
                <div class="card">
                    <div id="chart-type" class="chart-wrapper"></div>
                    <div class="story-layer">
                        <strong>Deducción de Formatos:</strong> El diagrama muestra la retención estructural. Ciertos formatos (como Video) propulsan el algoritmo, mientras que Imágenes estáticas afianzan a tu comunidad dura.
                    </div>
                </div>
            </div>

            <div class="card">
                <h2 style="color: #fff; margin-bottom: 5px;">Data Storytelling Avanzado: Análisis Combinado</h2>
                <div class="story-layer combo" style="margin-top:0;">
                    <strong>Fusión de Variables (El "Por Qué"):</strong> Al aplicar <em>Cross-Filtering</em> (seleccionando una red en el gráfico central), puedes detectar asimetrías. Por ejemplo, descubrir que TikTok genera el 80% de tus impresiones en formato 'Reel', pero Instagram centraliza el 90% de las interacciones estáticas. Esto indica que un canal actúa de "descubrimiento de embudo" (Top-of-Funnel) y el otro de "retención y conversión" (Bottom-of-Funnel). Usa esta deducción cruzada para no pedirle a TikTok conversiones que Instagram domina.
                </div>
            </div>

            <div class="card">
                <h2 style="color: #fff; margin-bottom: 5px;">Top Performance (Viralidad Orgánica)</h2>
                <p style="color: var(--text-muted); font-size: 14px;">Identificamos los 3 anclajes de contenido más valiosos basándonos en la penetración por Interacciones.</p>
                <div class="top-posts-grid" id="top-posts-container"></div>
            </div>

            <div class="card">
                <h2 style="color: #fff; margin-bottom: 5px;">Análisis de Diagnóstico y Ruta de Acción</h2>
                <p style="color: var(--text-muted); font-size: 14px;">Mapeo de fricciones individuales y soluciones de infraestructura.</p>
                <div class="strategy-grid">
                    <div class="strategy-box problems">
                        <h3>Problemáticas Detectadas</h3>
                        <ul>
                            <li>Tasas de interacción (engagement rate) estancadas en contenido estático o publicitario directo.</li>
                            <li>Ausencia de un gancho auditivo fuerte en los primeros 3 segundos de los videos.</li>
                            <li>La distribución del alcance depende excesivamente del feed orgánico.</li>
                        </ul>
                    </div>
                    <div class="strategy-box improvements">
                        <h3>Mejoras y Optimizaciones</h3>
                        <ul>
                            <li>Adaptar la narrativa a formatos 9:16 verticales de consumo acelerado (fast content).</li>
                            <li>Inyectar presupuestos fragmentados de pauta publicitaria en los posts con mayor tracción orgánica temprana.</li>
                            <li>Diversificar llamados a la acción (CTAs) apuntando a Guardados y Compartidos, no solo Likes.</li>
                        </ul>
                    </div>
                </div>
            </div>
        </div>

        <!-- Vista Exclusiva de Estrategia Corporativa -->
        <div id="strategy-view" style="display: none;">
            <div class="card corp-header">
                <h1>Diagnóstico Macro y Estrategia Conjunta</h1>
                <p>Análisis confidencial interno. Patrones globales detectados a través del cruce de los 7 portafolios procesados este mes.</p>
            </div>

            <div class="strategy-grid" style="grid-template-columns: 1fr; max-width: 900px; margin: 0 auto; gap: 30px;">
                <div class="strategy-box problems" style="background: rgba(244, 63, 94, 0.05); border-color: rgba(244, 63, 94, 0.2);">
                    <h3>Carga de Fricción Macro (Vulnerabilidades de Agencia)</h3>
                    <ul>
                        <li><strong>Fragmentación del Esfuerzo Orgánico:</strong> Hemos detectado que el 65% de los clientes depende estrictamente de Instagram como fuente primaria de interacciones, ignorando el potencial de adquisición de nuevos usuarios (Cold Leads) en TikTok.</li>
                        <li><strong>Agotamiento de Formatos Estáticos:</strong> Las imágenes (carruseles y gráficas estáticas) de todos los clientes muestran una caída transversal en impresiones frente al contenido vertical (Reels/TikTok). Seguir produciendo gráficas planas merma la rentabilidad del tiempo del equipo.</li>
                        <li><strong>Ausencia de Conversión Clara:</strong> Hay alto volumen de visualizaciones (tráfico de vanidad) pero faltan funnels de conversión definidos para dirigir este tráfico hacia WhatsApp o Landing Pages.</li>
                    </ul>
                </div>
                <div class="strategy-box improvements" style="background: rgba(16, 185, 129, 0.05); border-color: rgba(16, 185, 129, 0.2);">
                    <h3>Arquitectura de Mejoras (Pivot Estratégico)</h3>
                    <ul>
                        <li><strong>Migración a Ecosistema de Vídeo (Video-First):</strong> Reestructurar el departamento creativo para que el 80% del tiempo de producción se dedique a formatos verticales narrativos (Storytelling dinámico) con edición de retención alta.</li>
                        <li><strong>Inyección de Pauta Infiltrada:</strong> Integrar como política de agencia una porción del fee para "pauta empujada" (Boost). No depender del orgánico; los clientes necesitan tracción predecible.</li>
                        <li><strong>Sistema de Hooks (Ganchos) Estandarizados:</strong> Capacitar al equipo de Guion/Copywriting en estructuras de retención (los 3 primeros segundos). Establecer un protocolo donde cada video empiece con una pregunta disruptiva o una afirmación controversial.</li>
                        <li><strong>Consolidación del Data Storytelling:</strong> Usar esta plataforma mensualmente en las reuniones de cuenta. Mostrar a los clientes no solo números vacíos, sino la interpretación directiva (el "Por Qué") para aumentar la percepción de valor de la agencia.</li>
                    </ul>
                </div>
            </div>
        </div>

    </div>

    <script>
        // Init particles.js
        particlesJS("particles-js", {
            "particles": {
                "number": { "value": 70, "density": { "enable": true, "value_area": 900 } },
                "color": { "value": "#38bdf8" },
                "shape": { "type": "circle" },
                "opacity": { "value": 0.2, "random": true },
                "size": { "value": 3, "random": true },
                "line_linked": { "enable": true, "distance": 160, "color": "#38bdf8", "opacity": 0.15, "width": 1 },
                "move": { "enable": true, "speed": 1.2, "direction": "none", "random": true, "straight": false, "out_mode": "out" }
            },
            "interactivity": {
                "detect_on": "canvas",
                "events": { "onhover": { "enable": true, "mode": "grab" }, "onclick": { "enable": true, "mode": "push" } },
                "modes": { "grab": { "distance": 150, "line_linked": { "opacity": 0.6 } } }
            },
            "retina_detect": true
        });

        const dashboardData = __JSON_DATA__;
        let currentClientData = null;
        let currentFilterNetwork = null;

        const darkLayout = {
            paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
            font: { color: '#cbd5e1' },
            xaxis: { gridcolor: '#1e293b', zerolinecolor: '#1e293b' },
            yaxis: { gridcolor: '#1e293b', zerolinecolor: '#1e293b' }
        };

        const clientListEl = document.getElementById('client-list');
        for (const [clientName, data] of Object.entries(dashboardData)) {
            const div = document.createElement('div');
            div.className = 'client-item' + (data.is_strategy ? ' strategy-btn' : '');
            div.innerHTML = `<img src="${data.logo}" alt="${clientName}"> <span>${clientName}</span>`;
            div.onclick = () => loadClient(clientName, div);
            clientListEl.appendChild(div);
        }

        function loadClient(clientName, itemEl) {
            document.querySelectorAll('.client-item').forEach(el => el.classList.remove('active'));
            if(itemEl) itemEl.classList.add('active');
            
            const clientData = dashboardData[clientName];
            document.getElementById('empty-state').style.display = 'none';

            // Check if it's the corporate strategy view
            if(clientData.is_strategy) {
                document.getElementById('dashboard-view').style.display = 'none';
                document.getElementById('strategy-view').style.display = 'block';
                document.getElementById('header-client-logo').style.display = 'none';
                return;
            }

            // Standard Client View
            document.getElementById('strategy-view').style.display = 'none';
            document.getElementById('dashboard-view').style.display = 'block';
            
            currentClientData = clientData;
            currentFilterNetwork = null;

            const headerLogo = document.getElementById('header-client-logo');
            headerLogo.src = currentClientData.logo;
            headerLogo.style.display = 'block';

            document.getElementById('client-title').textContent = clientName;
            document.getElementById('val-seguidores').textContent = currentClientData.narrative.seguidores || '-';
            document.getElementById('val-impresiones').textContent = currentClientData.narrative.impresiones || '-';
            document.getElementById('val-interacciones').textContent = currentClientData.narrative.interacciones || '-';
            document.getElementById('val-publicaciones').textContent = currentClientData.narrative.publicaciones || '-';

            generateTopPosts();
            renderCharts();
        }

        function generateTopPosts() {
            const container = document.getElementById('top-posts-container');
            container.innerHTML = '';
            
            const sortedData = [...currentClientData.raw_data].sort((a,b) => b.interacciones - a.interacciones);
            const top3 = sortedData.slice(0, 3);

            if(top3.length === 0) return;

            top3.forEach((post, index) => {
                const hasLink = post.link && post.link !== '-' && post.link !== 'Unknown';
                const dateStr = post.fecha !== 'Unknown' ? post.fecha : 'Fecha no disp.';
                
                const card = document.createElement('div');
                card.className = 'top-post-card';
                card.innerHTML = `
                    <div class="rank-badge">#${index + 1}</div>
                    <div class="post-title">${(post.red||'Web').toUpperCase()} - ${post.tipo||'Generico'}</div>
                    <div class="post-meta">
                        <span>📅 ${dateStr}</span><br/>
                        <span>🔥 ${post.interacciones.toLocaleString()} Interacciones</span>
                    </div>
                    <a href="${hasLink ? post.link : '#'}" target="${hasLink ? '_blank' : '_self'}" 
                       class="btn-link ${!hasLink ? 'disabled' : ''}">
                       ${hasLink ? 'Abrir Enlace ↗' : 'Link no disponible'}
                    </a>
                `;
                container.appendChild(card);
            });
        }

        function getFilteredData() {
            let data = currentClientData.raw_data;
            if (currentFilterNetwork) data = data.filter(d => d.red === currentFilterNetwork);
            return data;
        }

        function renderCharts() {
            const filteredData = getFilteredData();
            
            // 1. Time Series Chart
            const timeMap = {};
            filteredData.forEach(d => {
                if(!timeMap[d.fecha]) timeMap[d.fecha] = { imp: 0, int: 0 };
                timeMap[d.fecha].imp += d.impresiones;
                timeMap[d.fecha].int += d.interacciones;
            });
            const sortedDates = Object.keys(timeMap).sort();
            const timeImp = sortedDates.map(date => timeMap[date].imp);
            const timeInt = sortedDates.map(date => timeMap[date].int);

            const traceImp = {
                x: sortedDates, y: timeImp, name: 'Impresiones',
                type: 'scatter', mode: 'lines+markers', line: { color: '#38bdf8', shape: 'spline' }
            };
            const traceInt = {
                x: sortedDates, y: timeInt, name: 'Interacciones',
                type: 'scatter', mode: 'lines+markers', yaxis: 'y2', line: { color: '#818cf8', shape: 'spline' }
            };

            const layoutTime = Object.assign({}, darkLayout, {
                title: 'Evolución Temporal ' + (currentFilterNetwork ? `(${currentFilterNetwork})` : ''),
                margin: { t: 40, l: 40, r: 40, b: 40 },
                yaxis: { title: 'Impresiones', color: '#38bdf8', gridcolor: '#1e293b' },
                yaxis2: { title: 'Interacciones', overlaying: 'y', side: 'right', color: '#818cf8', showgrid: false },
                legend: { orientation: 'h', y: -0.2 }
            });
            Plotly.newPlot('chart-time', [traceImp, traceInt], layoutTime, {responsive: true});

            // 2. Network Bar Chart
            const netMap = {};
            currentClientData.raw_data.forEach(d => {
                if(!netMap[d.red]) netMap[d.red] = 0;
                netMap[d.red] += d.impresiones;
            });
            const networks = Object.keys(netMap);
            const netVals = networks.map(n => netMap[n]);
            const netColors = networks.map(n => (n === currentFilterNetwork || !currentFilterNetwork) ? '#38bdf8' : '#1e293b');

            const traceNet = { x: networks, y: netVals, type: 'bar', marker: { color: netColors } };
            const layoutNet = Object.assign({}, darkLayout, { title: 'Impresiones por Red (Click para filtrar)', margin: { t: 40, l: 40, r: 20, b: 40 }, xaxis: { type: 'category' } });
            Plotly.newPlot('chart-network', [traceNet], layoutNet, {responsive: true});

            // 3. Content Type Donut Chart
            const typeMap = {};
            filteredData.forEach(d => {
                if(!typeMap[d.tipo]) typeMap[d.tipo] = 0;
                typeMap[d.tipo] += d.impresiones;
            });
            const types = Object.keys(typeMap);
            const typeVals = types.map(t => typeMap[t]);

            const traceType = {
                labels: types, values: typeVals, type: 'pie', hole: 0.5,
                marker: { colors: ['#f43f5e', '#38bdf8', '#10b981', '#8b5cf6', '#f59e0b'] },
                textinfo: 'percent+label', textfont: { color: '#fff' }
            };
            const layoutType = Object.assign({}, darkLayout, { title: 'Impresiones por Tipo de Contenido', margin: { t: 40, l: 20, r: 20, b: 20 }, showlegend: false });
            Plotly.newPlot('chart-type', [traceType], layoutType, {responsive: true});

            document.getElementById('chart-network').on('plotly_click', function(data){
                if(data.points && data.points.length > 0) {
                    const clickedNetwork = data.points[0].x;
                    currentFilterNetwork = (currentFilterNetwork === clickedNetwork) ? null : clickedNetwork;
                    renderCharts();
                }
            });
        }

        document.body.addEventListener('click', function(e) {
            const networkChart = document.getElementById('chart-network');
            if(networkChart && !networkChart.contains(e.target) && currentFilterNetwork) {
                currentFilterNetwork = null;
                renderCharts();
            }
        });
    </script>
</body>
</html>
"""
    
    agency_logo_path = os.path.join(base_dir, "LOGOS CLIENTES", "logo_midclick.png")
    agency_logo_b64 = image_to_base64(agency_logo_path)
    
    html_out = html_template.replace("__JSON_DATA__", json_data)
    html_out = html_out.replace("__AGENCY_LOGO__", agency_logo_b64)
    
    output_path = os.path.join(base_dir, "dashboard_clientes.html")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_out)
        
    print(f"Dashboard V3 generado exitosamente en: {output_path}")

if __name__ == '__main__':
    main()
