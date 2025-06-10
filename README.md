import pdfplumber
import re
import os
import pandas as pd

def extraer_con_regex(texto, regex):
    match = re.search(regex, texto, re.DOTALL)
    return match.group(1).strip() if match else "NO ENCONTRADO"

def extraer_datos(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        texto = '\n'.join(page.extract_text() for page in pdf.pages if page.extract_text())

    datos = {'Archivo': os.path.basename(pdf_path)}

    datos['Asegurado'] = extraer_con_regex(texto, r'Asegurado.*?:\s*(.*?)\n')
    datos['DNI'] = extraer_con_regex(texto, r'Nro.Documento.*?:\s*(\d{8,11})')
    datos['Numero de siniestro'] = extraer_con_regex(texto, r'Nro de Siniestro.*?:\s*(\d{5})')
    datos['Numero de poliza'] = extraer_con_regex(texto, r'Nro de P[oó]liza.*?:\s*(\d{7,8})')
    datos['Ramo'] = extraer_con_regex(texto, r'Rama.*?:\s*([A-Z\s]+)\b')
    datos['Causa'] = extraer_con_regex(texto, r'Motivo de Reserva..:\s*(.*?)\s+Monto Reclamado:')
    datos['Fecha de contratación'] = extraer_con_regex(texto, r'Fecha de Emisi[oó]n.*?:\s*(\d{2}/\d{2}/\d{4})')
    datos['Vigencia hasta'] = extraer_con_regex(texto, r'Hasta:\s*(\d{2}/\d{2}/\d{4})')
    datos['Fecha ocurrencia'] = extraer_con_regex(texto, r'Fecha de Ocurrencia.*?:\s*(\d{2}/\d{2}/\d{4})')
    datos['Fecha de denuncia'] = extraer_con_regex(texto, r'Fecha de Denuncia:?\s*(\d{2}/\d{2}/\d{4})')
    from datetime import datetime
    datos['Fecha de asignacion'] = datetime.today().strftime('%d/%m/%Y')
    datos['Domicilio de riesgo'] = extraer_con_regex(texto, r'Riesgo..............:\s*(.*?)\n')
    datos['Localidad'] = extraer_con_regex(texto, r'\d{4}-([A-Z\s]+)\b')
    datos['Cobertura'] = extraer_con_regex(texto, r'Cobertura.*?\n.*?\n.*?\s+([A-Z\s]+)\s+\$')
    datos['Telefono'] = extraer_con_regex(texto, r'Tel[eé]fono.*?:\s*"?(\d+)"?')

    return datos

if __name__ == "__main__":
    carpeta_pdfs = "pdfs"
    archivos = [f for f in os.listdir(carpeta_pdfs) if f.endswith(".pdf")]
    resultados = []

    print(f"📂 Procesando {len(archivos)} archivos en '{carpeta_pdfs}'...\n")

    for archivo in archivos:
        ruta_completa = os.path.join(carpeta_pdfs, archivo)
        datos = extraer_datos(ruta_completa)
        resultados.append(datos)

        print(f"✅ {archivo} procesado")

    # Guardar en Excel
    df = pd.DataFrame(resultados)
    df['Calle'] = df['Domicilio de riesgo'].str.extract(r'^([A-ZÁÉÍÓÚÑ\s]+)').fillna('')
    df['Número'] = df['Domicilio de riesgo'].str.extract(r'(\d+)$').fillna('')
    df.drop(columns=['Domicilio de riesgo'], inplace=True)
    # Orden personalizado de las columnas para el Excel
    columnas_ordenadas = [
    'Archivo',
    'Asegurado',
    'DNI',
    'Numero de siniestro',
    'Numero de poliza',
    'Ramo',
    'Causa',
    'Fecha de contratación',
    'Vigencia hasta',
    'Fecha de denuncia',
    'Fecha ocurrencia',
    'Fecha de asignacion',
    'Calle',
    'Número',
    'Localidad',
    'Telefono',
    'Cobertura',
    ]

    df = df[columnas_ordenadas]
    df.to_excel("resultados.xlsx", index=False)
    print("\n📄 Archivo 'resultados.xlsx' generado con éxito.")
