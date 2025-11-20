from flask import Flask, request, send_file
import os
import sys
import warnings
import tempfile
import traceback

# Suprimir advertencias de dependencias opcionales
warnings.filterwarnings('ignore')

# Librerías necesarias para exportar Excel
from openpyxl import Workbook

# Estrategia 1: pdfplumber (más confiable)
try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

# Estrategia 2: tabula-py
try:
    import tabula
    TABULA_AVAILABLE = True
except ImportError:
    TABULA_AVAILABLE = False

app = Flask(__name__)

def normalizar_celda(celda):
    """Normaliza contenido de celda."""
    if celda is None:
        return ""
    celda_str = str(celda).strip()
    celda_str = re.sub(r'\s+', ' ', celda_str)
    return celda_str

def extraer_con_pdfplumber(pdf_path):
    """Extrae tablas usando pdfplumber."""
    if not PDFPLUMBER_AVAILABLE:
        return None

    try:
        print("📊 Intentando extracción con pdfplumber...")
        todas_filas = []

        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()

                for table in tables:
                    if table:
                        # Limpiar y normalizar datos
                        for row in table:
                            fila_limpia = [normalizar_celda(cell) for cell in row]
                            if any(fila_limpia):  # Solo filas no vacías
                                todas_filas.append(fila_limpia)

        if todas_filas:
            print(f"   ✅ pdfplumber encontró {len(todas_filas)} filas")
            return todas_filas
        return None
    except Exception as e:
        print(f"   ⚠️ pdfplumber falló: {e}")
        return None

def extraer_con_tabula(pdf_path):
    """Extrae tablas usando tabula-py."""
    if not TABULA_AVAILABLE:
        return None

    try:
        print("📊 Intentando extracción con tabula-py...")
        dfs = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)

        if dfs:
            todas_filas = []
            for df in dfs:
                for _, row in df.iterrows():
                    fila = [str(val) if pd.notna(val) else "" for val in row]
                    if any(fila):
                        todas_filas.extend(fila)
            print(f"   ✅ tabula-py encontró {len(todas_filas)} filas")
            return todas_filas
        return None
    except Exception as e:
        print(f"   ⚠️ tabula-py falló: {e}")
        return None

@app.route('/convert', methods=['POST'])
def convert_pdf():
    try:
        if 'file' not in request.files:
            return {'error': 'No file provided'}, 400

        file = request.files['file']
        if file.filename == '':
            return {'error': 'No file selected'}, 400

        if not file.filename.lower().endswith('.pdf'):
            return {'error': 'File must be a PDF'}, 400

        # Guardar archivo temporalmente
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
            file.save(temp_pdf.name)
            pdf_path = temp_pdf.name

        try:
            todas_filas = []

            # Intentar pdfplumber primero
            if PDFPLUMBER_AVAILABLE:
                todas_filas = extraer_con_pdfplumber(pdf_path)

            # Si no funcionó, intentar tabula-py
            if not todas_filas and TABULA_AVAILABLE:
                todas_filas = extraer_con_tabula(pdf_path)

            if not todas_filas:
                return {'error': 'Could not extract data from PDF'}, 500

            # Crear Excel
            wb = Workbook()
            ws = wb.active

            for fila in todas_filas:
                ws.append(fila)

            # Guardar Excel temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_xlsx:
                wb.save(temp_xlsx.name)
                xlsx_path = temp_xlsx.name

            # Enviar archivo
            response = send_file(
                xlsx_path,
                as_attachment=True,
                download_name='converted.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

            # Limpiar archivos temporales después de enviar
            @response.call_on_close
            def cleanup():
                try:
                    os.unlink(pdf_path)
                    os.unlink(xlsx_path)
                except:
                    pass

            return response

        except Exception as e:
            print(f"Error processing PDF: {e}")
            traceback.print_exc()
            return {'error': 'Error processing file'}, 500

    except Exception as e:
        print(f"Unexpected error: {e}")
        traceback.print_exc()
        return {'error': 'Internal server error'}, 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    app.run(host='0.0.0.0', port=port, debug=False)
