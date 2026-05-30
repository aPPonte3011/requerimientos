import os
import re
import traceback
from datetime import datetime
from collections import Counter
from flask import Flask, render_template, request, jsonify, session, send_file
from flask_cors import CORS

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "mobilia-prod-key-2024")
CORS(app)  # ✅ Habilita CORS para evitar bloqueos en despliegue

def parsear_fecha(fecha_str):
    fecha_str = fecha_str.replace('-', '/').replace('.', '/')
    for fmt in ['%d/%m/%Y', '%d/%m/%y', '%Y/%m/%d', '%d/%m']:
        try:
            dt = datetime.strptime(fecha_str.strip(), fmt)
            if dt.year < 100: dt = dt.replace(year=dt.year + 2000)
            return dt
        except ValueError: continue
    return None

def procesar_chat(texto_chat, fecha_inicio=None, fecha_fin=None, tipo_inmueble=None):
    patron_palabras = r"\b(requiero|solicito|necesito|compro|se requiere|se necesita|se solicita|se busca)\b"
    lineas = texto_chat.split('\n')
    requerimientos_lista = []
    inmuebles_conteo = []
    agentes_lista = []
    vistos = set()
    patron_mensaje = re.compile(r"^\[?(\d{1,4}[/\-\.]\d{1,2}[/\-\.]\d{2,4}),?\s\d{1,2}:\d{2}(?::\d{2})?\]?\s(?:-\s)?([^:]+):\s(.*)$")

    for linea in lineas:
        match = patron_mensaje.match(linea.strip())
        if match:
            fecha_str, contacto, mensaje = match.groups()
            fecha_obj = parsear_fecha(fecha_str)
            if fecha_inicio and fecha_obj and fecha_obj < fecha_inicio: continue
            if fecha_fin and fecha_obj and fecha_obj > fecha_fin: continue

            if re.search(patron_palabras, mensaje, re.IGNORECASE):
                hash_mensaje = mensaje.strip().lower()
                if hash_mensaje not in vistos:
                    vistos.add(hash_mensaje)
                    msg_lower = mensaje.lower()
                    tipo_detectado = "Otro"
                    if any(p in msg_lower for p in ["casa", "quinta", "chalet", "duplex"]): tipo_detectado = "Casa"
                    elif any(p in msg_lower for p in ["apartamento", "apto", "depto", "ph", "flat"]): tipo_detectado = "Apartamento"
                    elif any(p in msg_lower for p in ["local", "oficina", "consultorio", "comercial"]): tipo_detectado = "Local/Oficina"
                    elif any(p in msg_lower for p in ["terreno", "finca", "lote", "parcela"]): tipo_detectado = "Terreno"
                    elif any(p in msg_lower for p in ["galpon", "galpón", "bodega", "almacén"]): tipo_detectado = "Galpón"

                    if tipo_inmueble and tipo_detectado != tipo_inmueble: continue

                    telefono_match = re.search(r"(\+?\d{2,4}[\s-]?\d{3,4}[\s-]?\d{4,7})", mensaje)
                    telefono = telefono_match.group(1) if telefono_match else "No especificado"

                    requerimientos_lista.append({
                        "fecha": fecha_str, "requerimiento": mensaje,
                        "contacto": contacto, "telefono": telefono,
                        "tipo_inmueble": tipo_detectado
                    })
                    inmuebles_conteo.append(tipo_detectado)
                    agentes_lista.append(contacto)

    return {
        "tabla": requerimientos_lista,
        "estadisticas": dict(Counter(inmuebles_conteo)),
        "ranking": [{"nombre": k, "mensajes": v} for k, v in Counter(agentes_lista).most_common(5)]
    }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file uploaded"}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "Empty filename"}), 400
        
        texto_chat = file.read().decode('utf-8', errors='ignore')
        f_inicio = request.form.get('fecha_inicio')
        f_fin = request.form.get('fecha_fin')
        tipo = request.form.get('tipo_inmueble') or None

        dt_inicio = datetime.strptime(f_inicio, '%Y-%m-%d') if f_inicio else None
        dt_fin = datetime.strptime(f_fin, '%Y-%m-%d').replace(hour=23, minute=59, second=59) if f_fin else None

        resultados = procesar_chat(texto_chat, dt_inicio, dt_fin, tipo)
        session['tabla_datos'] = resultados.get('tabla', [])

        response = jsonify(resultados)
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate'
        return response
    except Exception as e:
        print(f"❌ ERROR EN /upload: {traceback.format_exc()}")
        return jsonify({"error": str(e)}), 500

@app.route('/export/excel')
def export_excel():
    import pandas as pd
    from io import BytesIO
    data = session.get('tabla_datos', [])
    if not data: return jsonify({"error": "No data"}), 404
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Requerimientos')
    output.seek(0)
    return send_file(output, download_name='requerimientos.xlsx', as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/export/pdf')
def export_pdf():
    from weasyprint import HTML
    from io import BytesIO
    data = session.get('tabla_datos', [])
    if not data: return jsonify({"error": "No data"}), 404
    html = """<html><head><meta charset="utf-8"><style>body{font-family:Arial;padding:20px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #ddd;padding:8px;text-align:left}th{background:#3498db;color:#fff}</style></head><body><h1>Reporte Mobilia</h1><table><tr><th>Fecha</th><th>Tipo</th><th>Requerimiento</th><th>Contacto</th></tr>"""
    for row in data:
        html += f"<tr><td>{row['fecha']}</td><td>{row['tipo_inmueble']}</td><td>{row['requerimiento']}</td><td>{row['contacto']}</td></tr>"
    html += "</table></body></html>"
    pdf = BytesIO()
    HTML(string=html).write_pdf(pdf)
    pdf.seek(0)
    return send_file(pdf, download_name='requerimientos.pdf', as_attachment=True, mimetype='application/pdf')

if __name__ == '__main__':
    app.run(debug=False, port=5000)
    
