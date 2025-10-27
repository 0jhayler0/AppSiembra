import sys
import pdfplumber
from openpyxl import Workbook

def main(pdf_path):
    """
    Lee el PDF usando pdfplumber, extrae las tablas interpretando celdas,
    y convierte las tablas en un archivo Excel.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            wb = Workbook()
            ws = wb.active
            ws.title = "Tabla_Extraida"

            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        # Limpiar celdas: convertir a string y trim
                        cleaned_row = [str(cell).strip() if cell is not None else "" for cell in row]
                        ws.append(cleaned_row)
                    # Agregar fila en blanco entre tablas para separación
                    ws.append([])

        # Generar path de salida reemplazando .pdf con _converted.xlsx
        output_path = pdf_path.replace('.pdf', '_converted.xlsx')
        wb.save(output_path)
        print(output_path)  # Imprimir el path para que server.js lo capture

    except Exception as e:
        print(f"Error procesando el PDF: {e}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python convertidor.py <ruta_al_pdf>")
        sys.exit(1)
    pdf_path = sys.argv[1]
    main(pdf_path)
