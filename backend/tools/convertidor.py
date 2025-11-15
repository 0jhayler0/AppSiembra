#!/usr/bin/env python3
"""
Convertidor robusto de PDF a XLSX con múltiples estrategias de extracción.
Intenta Camelot → PyMuPDF → pdfplumber con fallback automático.
"""
import sys
import os
import re
import time
import pandas as pd
import warnings

# Suprimir advertencias de dependencias opcionales
warnings.filterwarnings('ignore')

# Estrategia 1: Camelot (mejor para tablas)
try:
    import camelot
    CAMELOT_AVAILABLE = True
except ImportError:
    CAMELOT_AVAILABLE = False

# Estrategia 2: PyMuPDF
try:
    import fitz
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

# Estrategia 3: pdfplumber (fallback final)
try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

def normalizar_celda(celda):
    """Normaliza contenido de celda."""
    if celda is None:
        return ""
    celda_str = str(celda).strip()
    # Eliminar saltos de línea múltiples y espacios extra
    celda_str = re.sub(r'\s+', ' ', celda_str)
    return celda_str

def extraer_con_camelot(pdf_path):
    """Intenta extraer usando Camelot (mejor para tablas regulares)."""
    if not CAMELOT_AVAILABLE:
        return None
    
    try:
        print("📊 Intentando extracción con Camelot...")
        
        # Intenta con diferentes "flavors"
        for flavor in ['lattice', 'stream']:
            try:
                tables = camelot.read_pdf(pdf_path, pages='all', flavor=flavor)
                if tables:
                    print(f"   ✅ Camelot encontró {len(tables)} tabla(s) con flavor='{flavor}'")
                    return tables
            except:
                continue
        
        return None
    except Exception as e:
        print(f"   ⚠️ Camelot falló: {e}")
        return None

def extraer_con_pymupdf(pdf_path):
    """Intenta extraer usando PyMuPDF (muy robusto)."""
    if not PYMUPDF_AVAILABLE:
        return None
    
    try:
        print("📊 Intentando extracción con PyMuPDF...")
        doc = fitz.open(pdf_path)
        resultados = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            try:
                tables = page.find_tables()
                
                for tabla in tables:
                    datos = tabla.extract()
                    if datos:
                        resultados.append({
                            'datos': datos,
                            'page': page_num,
                            'texto': page.get_text()
                        })
            except:
                pass
        
        doc.close()
        
        if resultados:
            print(f"   ✅ PyMuPDF encontró {len(resultados)} tabla(s)")
            return resultados
        
        return None
    except Exception as e:
        print(f"   ⚠️ PyMuPDF falló: {e}")
        return None

def extraer_con_pdfplumber(pdf_path):
    """Intenta extraer usando pdfplumber (extracción mejorada)."""
    if not PDFPLUMBER_AVAILABLE:
        return None
    
    try:
        print("📊 Intentando extracción con pdfplumber...")
        
        with pdfplumber.open(pdf_path) as pdf:
            resultados = []
            
            for page_idx, page in enumerate(pdf.pages):
                # Usar extract_tables() con configuración específica
                try:
                    tables = page.extract_tables(
                        table_settings={
                            "vertical_strategy": "lines_strict",
                            "horizontal_strategy": "lines_strict",
                        }
                    )
                    
                    if not tables:
                        # Fallback a estrategia más laxa
                        tables = page.extract_tables()
                    
                    if tables:
                        for tabla in tables:
                            resultados.append({
                                'datos': tabla,
                                'page': page_idx,
                                'texto': page.extract_text()
                            })
                except:
                    pass
            
            if resultados:
                print(f"   ✅ pdfplumber encontró {len(resultados)} tabla(s)")
                return resultados
        
        return None
    except Exception as e:
        print(f"   ⚠️ pdfplumber falló: {e}")
        return None

def procesar_tablas(tablas, fuente="desconocida"):
    """Procesa tablas extraídas en formato estándar."""
    filas_totales = []
    bloques_count = 0
    
    for item in (tablas if isinstance(tablas, list) else []):
        bloques_count += 1
        
        # Obtener datos según la fuente
        if hasattr(item, 'df'):  # Camelot
            datos = item.df.values.tolist()
            titulo = "Tabla extraída con Camelot"
        elif isinstance(item, dict) and 'datos' in item:  # PyMuPDF o pdfplumber
            datos = item['datos']
            texto = item.get('texto', '')
            
            # Intentar extraer título del texto
            patron = re.compile(
                r"Flores de la Victoria S\.A\.S Semana Siembra\s+(\d+)\s+Seccion:\s*(\d+)",
                re.IGNORECASE
            )
            match = patron.search(texto)
            titulo = match.group(0) if match else f"Tabla extraída con {fuente}"
        else:
            datos = item if isinstance(item, list) else []
            titulo = f"Tabla extraída con {fuente}"
        
        if not datos:
            continue
        
        print(f"   📊 Bloque {bloques_count}: {len(datos)} fila(s)")
        
        # Agregar título
        filas_totales.append([titulo] + [""] * 11)
        filas_totales.append([
            "Nave", "Era", "Variedad", "Largo", "Fecha Siembra", "Inicio Corte",
            "Nave", "Era", "Variedad", "Largo", "Fecha Siembra", "Inicio Corte"
        ])
        
        # Agregar datos normalizados
        for fila in datos:
            fila_norm = [normalizar_celda(c) for c in fila]
            while len(fila_norm) < 12:
                fila_norm.append("")
            filas_totales.append(fila_norm[:12])
        
        filas_totales.append([""] * 12)
    
    return filas_totales, bloques_count

def main():
    if len(sys.argv) < 3:
        print("Uso: python convertidor_robusto.py input.pdf output.xlsx")
        sys.exit(1)

    input_pdf = sys.argv[1]
    output_xlsx = sys.argv[2]

    if not os.path.exists(input_pdf):
        print(f"❌ Archivo no encontrado: {input_pdf}")
        print(f"   Ruta: {os.path.abspath(input_pdf)}")
        sys.exit(2)

    print(f"🔍 Procesando: {os.path.abspath(input_pdf)}\n")
    start_time = time.time()
    todas_filas = []
    bloques_procesados = 0
    metodo_usado = ""

    # Mostrar disponibilidad de motores
    print("📦 Motores disponibles:")
    print(f"   - Camelot: {'✅' if CAMELOT_AVAILABLE else '❌'}")
    print(f"   - PyMuPDF: {'✅' if PYMUPDF_AVAILABLE else '❌'}")
    print(f"   - pdfplumber: {'✅' if PDFPLUMBER_AVAILABLE else '❌'}\n")

    # Estrategia 1: Camelot
    if CAMELOT_AVAILABLE and not todas_filas:
        try:
            tablas = extraer_con_camelot(input_pdf)
            if tablas:
                todas_filas, bloques_procesados = procesar_tablas(tablas, "Camelot")
                metodo_usado = "Camelot"
                print(f"✅ Extracción exitosa con Camelot\n")
        except Exception as e:
            print(f"⚠️ Camelot falló: {e}\n")

    # Estrategia 2: PyMuPDF (si Camelot no funcionó)
    if not todas_filas and PYMUPDF_AVAILABLE:
        try:
            tablas = extraer_con_pymupdf(input_pdf)
            if tablas:
                todas_filas, bloques_procesados = procesar_tablas(tablas, "PyMuPDF")
                metodo_usado = "PyMuPDF"
                print(f"✅ Extracción exitosa con PyMuPDF\n")
        except Exception as e:
            print(f"⚠️ PyMuPDF falló: {e}\n")

    # Estrategia 3: pdfplumber
    if not todas_filas and PDFPLUMBER_AVAILABLE:
        try:
            tablas = extraer_con_pdfplumber(input_pdf)
            if tablas:
                todas_filas, bloques_procesados = procesar_tablas(tablas, "pdfplumber")
                metodo_usado = "pdfplumber"
                print(f"✅ Extracción exitosa con pdfplumber\n")
        except Exception as e:
            print(f"⚠️ pdfplumber falló: {e}\n")

    if not todas_filas:
        print("❌ No se pudieron extraer datos con ninguna estrategia.")
        print("   Por favor, instala al menos una de estas dependencias:")
        print("   - pip install camelot-py")
        print("   - pip install PyMuPDF")
        print("   - pip install pdfplumber")
        sys.exit(3)

    # Exportar a Excel
    try:
        df = pd.DataFrame(todas_filas)
        df.to_excel(output_xlsx, index=False, header=False)
        
        end_time = time.time()
        elapsed = end_time - start_time
        
        print(f"🎉 Archivo generado → {output_xlsx}")
        print(f"📊 Total filas en Excel: {len(todas_filas)}")
        print(f"📦 Bloques procesados: {bloques_procesados}")
        print(f"🔧 Método utilizado: {metodo_usado}")
        print(f"⏱️ Tiempo total: {elapsed:.2f} segundos")
        
        return 0

    except Exception as e:
        print(f"❌ Error exportando Excel: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(4)

if __name__ == "__main__":
    sys.exit(main())
