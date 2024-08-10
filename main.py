import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from docx import Document


# Definición de las dimensiones y categorías
dimensiones = ["Mercado", "Servicio", "Producto", "Datos", "Tecnología", "Aliados", "Capacidades"]
categorias = ["Núcleo integrado", "Servicios digitales", "Marketplace digital", "Ecosistema digital"]

# Inicialización de un diccionario para almacenar las respuestas
matriz = {dim: {cat: "" for cat in categorias} for dim in dimensiones}

# Iteración para pedir al usuario que complete la matriz
for dim in dimensiones:
    for cat in categorias:
        respuesta = input(f"Ingrese el contenido para '{dim}' en la categoría '{cat}': ")
        matriz[dim][cat] = respuesta

# Creación de un DataFrame de pandas
df = pd.DataFrame(matriz).T

# Crear un nuevo archivo de Excel con openpyxl
wb = Workbook()
ws = wb.active
ws.title = "Matriz Completada"

# Añadir las filas del DataFrame al worksheet
for r_idx, row in enumerate(dataframe_to_rows(df, index=True, header=True), 1):
    for c_idx, value in enumerate(row, 1):
        ws.cell(row=r_idx, column=c_idx, value=value)

# Aplicar estilos
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Aplicar estilo a los encabezados de las categorías (primera fila)
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = alignment

# Aplicar estilo a los encabezados de las dimensiones (primera columna)
for cell in ws['A']:
    if cell.row != 1:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = alignment

# Ajuste de ancho de columnas y alto de filas
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column].width = adjusted_width

# Aplicar bordes a todas las celdas
thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))

for row in ws.iter_rows():
    for cell in row:
        cell.border = thin_border

# Guardar el archivo de Excel
output_file = "matriz_completada_formateada.xlsx"
wb.save(output_file)

# Generar el informe en Word
doc = Document()
doc.add_heading('Informe de Matriz Completada', 0)

for dim in dimensiones:
    doc.add_heading(dim, level=1)
    for cat in categorias:
        contenido = matriz[dim][cat]
        if contenido:
            doc.add_heading(f'{cat}:', level=2)
            elementos = [elem.strip() for elem in contenido.split(',') if elem.strip()]  # Filtra elementos vacíos
            for elem in elementos:
                doc.add_paragraph(elem, style='ListBullet')

# Guardar el documento Word
output_word = "informe_matriz.docx"
doc.save(output_word)


print(f"¡Matriz completada y formateada! El archivo Excel '{output_file}' y el informe Word '{output_word}' se han guardado exitosamente.")


