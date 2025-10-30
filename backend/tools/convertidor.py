#!/usr/bin/env python3
import re
import pdfplumber
import pandas as pd
import sys
import os
from tabulate import tabulate  # pip install tabulate

# === Funciones auxiliares ===

def extraer_bloques(texto):
    """
    Separa el PDF en bloques por secciones detectando el encabezado:
    'Flores de la Victoria S.A.S Semana Siembra XXXX Seccion: XX'
    """
    patron = re.compile(
        r"(Flores de la Victoria S\.A\.S Semana Siembra\s+(\d+)\s+Seccion:\s*(\d+))",
        re.IGNORECASE
    )
    bloques = []
    coincidencias = list(patron.finditer(texto))
    for i, match in enumerate(coincidencias):
        inicio = match.end()
        fin = coincidencias[i + 1].start() if i + 1 < len(coincidencias) else len(texto)
        contenido = texto[inicio:fin].strip()
        bloques.append({
            "titulo": match.group(1).strip(),
            "semana": match.group(2),
            "seccion": match.group(3),
            "contenido": contenido
        })
    return bloques

def extraer_filas(texto):
    """
    Extrae filas completas del bloque (lado A y B).
    Si un lado no existe, genera celdas vacías para mantener la estructura.
    Detecta variedades con paréntesis, comas, puntos, guiones, etc.
    """
    filas = []
    lineas = [l.strip() for l in texto.splitlines() if l.strip()]

    patron_fila = re.compile(
        r"(?:(\d{1,2})\s+)?(\d+)\s+(.+?)\s+(\d{1,2}\.\d)\s+(\d{2}-[A-Za-z]{3})\s+(\d{2}-[A-Za-z]{3})",
        re.IGNORECASE | re.DOTALL
    )

    nave_actual = None

    for linea in lineas:
        partes = patron_fila.findall(linea)

        # si no hay nada reconocible, ignoramos la línea
        if not partes:
            continue

        # Caso 1: dos coincidencias → lado A y B en la misma línea
        if len(partes) == 2:
            match_a, match_b = partes

            nave_a = match_a[0] if match_a[0] else nave_actual
            nave_actual = nave_a

            nave_b = match_b[0] if match_b[0] else nave_actual

            filas.append({
                "A": {
                    "Nave": nave_a or "",
                    "Era": match_a[1],
                    "Variedad": match_a[2].strip(),
                    "Largo": match_a[3],
                    "Fecha_Siembra": match_a[4],
                    "Inicio_Corte": match_a[5]
                },
                "B": {
                    "Nave": nave_b or "",
                    "Era": match_b[1],
                    "Variedad": match_b[2].strip(),
                    "Largo": match_b[3],
                    "Fecha_Siembra": match_b[4],
                    "Inicio_Corte": match_b[5]
                }
            })

        # Caso 2: solo un lado → llenamos el otro con vacío
        elif len(partes) == 1:
            match = partes[0]
            nave = match[0] if match[0] else nave_actual
            nave_actual = nave

            lado_a = {
                "Nave": nave or "",
                "Era": match[1],
                "Variedad": match[2].strip(),
                "Largo": match[3],
                "Fecha_Siembra": match[4],
                "Inicio_Corte": match[5]
            }

            # Crea el otro lado vacío (por estructura)
            lado_b = {
                "Nave": "", "Era": "", "Variedad": "", "Largo": "",
                "Fecha_Siembra": "", "Inicio_Corte": ""
            }

            filas.append({"A": lado_a, "B": lado_b})

    return filas


def dividir_lados(filas):
    """
    Divide las filas alternando entre lado A y lado B.
    Si hay más del doble, se empareja secuencialmente.
    """
    lado_a, lado_b = [], []
    for i, fila in enumerate(filas):
        if i % 2 == 0:
            lado_a.append(fila)
        else:
            lado_b.append(fila)
    return lado_a, lado_b

# === Función principal ===

def main():
    if len(sys.argv) < 3:
        print("Uso: python convertidor.py input.pdf output.xlsx")
        sys.exit(1)

    input_pdf = sys.argv[1]
    output_xlsx = sys.argv[2]

    if not os.path.exists(input_pdf):
        print("❌ No se encontró el archivo PDF.")
        sys.exit(2)

    todas_filas = []
    semana_detectada = None

    with pdfplumber.open(input_pdf) as pdf:
        texto_total = ""
        for page in pdf.pages:
            texto_total += page.extract_text() + "\n"

        bloques = extraer_bloques(texto_total)
        if not bloques:
            print("⚠️ No se encontraron bloques de secciones en el PDF.")
            sys.exit(3)

        for b in bloques:
            titulo = b["titulo"]
            semana_detectada = b["semana"]
            filas = extraer_filas(b["contenido"])
            lado_a, lado_b = dividir_lados(filas)
            max_len = max(len(lado_a), len(lado_b))

            print(f"✅ Sección {b['seccion']}: {len(lado_a)} A + {len(lado_b)} B")

            # Fila del título
            todas_filas.append([titulo] + [""] * 12)
            # Encabezados
            todas_filas.append([
                "Nave", "Era", "Variedad", "Largo", "Fecha Siembra", "Inicio Corte",
                "Nave", "Era", "Variedad", "Largo", "Fecha Siembra", "Inicio Corte"
            ])

            
            # Rellenar filas
            for registro in filas:
                a = registro["A"]
                b = registro["B"]

                fila = [
                    a.get("Nave", ""), a.get("Era", ""), a.get("Variedad", ""), a.get("Largo", ""),
                    a.get("Fecha_Siembra", ""), a.get("Inicio_Corte", ""),
                    b.get("Nave", ""), b.get("Era", ""), b.get("Variedad", ""), b.get("Largo", ""),
                    b.get("Fecha_Siembra", ""), b.get("Inicio_Corte", "")
                ]

                todas_filas.append(fila)


            # Espacio entre secciones
            todas_filas.append([""] * 12)

            

    # Crear DataFrame
    df = pd.DataFrame(todas_filas)

    # Mostrar en consola (primeras 40 filas)
    print("\n📋 Vista previa de tabla generada:\n")
    print(tabulate(df, tablefmt="grid", showindex=False))

    # Exportar a Excel
    df.to_excel(output_xlsx, index=False, header=False)
    print(f"\n🎉 Archivo Excel generado correctamente → {output_xlsx}")
    if semana_detectada:
        print(f"🗓 Semana detectada: {semana_detectada}")

if __name__ == "__main__":
    main()
