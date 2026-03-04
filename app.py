import openpyxl
import os
import sys
import sqlite3
import datetime
from flask import Flask, request, jsonify, send_from_directory, render_template
from werkzeug.utils import secure_filename
import pandas as pd

# Importar funciones del script original (asumiendo que está en el mismo directorio)
import analizador_rso as rso

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['INFORME_FOLDER'] = 'static/informes'
app.config['GRAFICA_FOLDER'] = 'static/graficas'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB

# Crear carpetas si no existen
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['INFORME_FOLDER'], exist_ok=True)
os.makedirs(app.config['GRAFICA_FOLDER'], exist_ok=True)

# Ruta principal: sirve el index.html
@app.route('/')
def index():
    return render_template('index.html')

# Endpoint para subir y procesar archivo
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No se envió ningún archivo'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Nombre de archivo vacío'}), 400
    if not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        return jsonify({'error': 'Formato no soportado. Use .xlsx'}), 400

    # Guardar archivo temporal
    filename = secure_filename(file.filename)
    temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(temp_path)

    # Procesar el archivo usando la función del script original
    try:
        # Llamamos a procesar_archivo, pero necesitamos adaptarla para que use nuestras carpetas
        # y devuelva las rutas.
        # Modificaremos el script original para que acepte parámetros de salida.
        # Por ahora, crearemos una función wrapper que llame a las funciones internas.
        resultado = procesar_archivo_web(temp_path, app.config['INFORME_FOLDER'], app.config['GRAFICA_FOLDER'])
        # resultado contendrá: informe_path, grafica_path, datos_resumen (para mostrar en frontend)
        return jsonify(resultado)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
    finally:
        # Opcional: eliminar archivo temporal
        os.remove(temp_path)

# Endpoint para obtener lista de reportes históricos
@app.route('/reportes')
def get_reportes():
    conn = sqlite3.connect(rso.DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        SELECT nombre_archivo, fecha_reporte, spi, cpi, avance_real, avance_planificado
        FROM reportes
        ORDER BY fecha_reporte DESC
    ''')
    rows = cursor.fetchall()
    conn.close()
    reportes = []
    for row in rows:
        reportes.append({
            'nombre_archivo': row[0],
            'fecha_reporte': row[1],
            'spi': row[2],
            'cpi': row[3],
            'avance_real': row[4],
            'avance_planificado': row[5]
        })
    return jsonify(reportes)

# Endpoint para servir gráficas
@app.route('/graficas/<path:filename>')
def serve_grafica(filename):
    return send_from_directory(app.config['GRAFICA_FOLDER'], filename)

# Endpoint para servir informes de texto
@app.route('/informes/<path:filename>')
def serve_informe(filename):
    return send_from_directory(app.config['INFORME_FOLDER'], filename)

# Función wrapper que adapta el procesamiento del script original
def procesar_archivo_web(archivo_path, informe_folder, grafica_folder):
    # Extraer nombre base
    nombre_archivo = os.path.basename(archivo_path)
    nombre_base = os.path.splitext(nombre_archivo)[0]

    # Cargar workbook
    wb = openpyxl.load_workbook(archivo_path, data_only=True)

    # Leer hojas (usando funciones del script original)
    meta, df_actividades, totales, df_full = rso.leer_hoja_rdo(wb)
    try:
        df_curva = rso.leer_hoja_curva(wb)
    except Exception as e:
        print(f"Advertencia: no se pudo leer curva: {e}")
        df_curva = pd.DataFrame(columns=["Fecha", "% Previsto Acumulado", "% Real Acumulado"])

    # Obtener avances
    avance_real = None
    avance_plan = None
    if not df_curva.empty:
        df_r = df_curva[df_curva["% Real Acumulado"] > 0]
        if not df_r.empty:
            avance_real = float(df_r.iloc[-1]["% Real Acumulado"])
        df_p = df_curva[df_curva["Fecha"] <= pd.Timestamp.today()]
        if not df_p.empty:
            avance_plan = float(df_p.iloc[-1]["% Previsto Acumulado"])

    # Generar gráfica (guardar en grafica_folder)
    grafica_filename = f"curva_{nombre_base}.png"
    grafica_path = os.path.join(grafica_folder, grafica_filename)
    if not df_curva.empty:
        rso.generar_grafica(df_curva, nombre_archivo, grafica_path)
    else:
        grafica_path = None

    # Generar informe de texto
    informe = rso.generar_informe(nombre_archivo, meta, df_actividades, totales, df_curva, df_full=df_full)
    informe_filename = f"informe_{nombre_base}.txt"
    informe_path = os.path.join(informe_folder, informe_filename)
    with open(informe_path, "w", encoding="utf-8") as f:
        f.write(informe)

    # Guardar en DB
    rso.guardar_en_db(nombre_archivo, meta, totales, avance_real, avance_plan)

    # Devolver rutas relativas (para que el frontend pueda acceder)
    resultado = {
        'informe_url': f'/informes/{informe_filename}',
        'grafica_url': f'/graficas/{grafica_filename}' if grafica_path else None,
        'nombre_archivo': nombre_archivo,
        'fecha_reporte': str(meta.get('fecha_reporte')),
        'spi': meta.get('spi'),
        'cpi': meta.get('cpi'),
        'avance_real': avance_real,
        'avance_planificado': avance_plan,
        'informe_texto': informe  # opcional: enviar el texto completo para mostrarlo
    }
    return resultado

if __name__ == '__main__':
    app.run(debug=True)