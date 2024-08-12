import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from docx import Document
from docx.shared import RGBColor
from io import BytesIO
from docx.oxml import OxmlElement, ns

# Definición de las dimensiones y categorías
dimensiones = ["Mercado", "Servicio", "Producto", "Datos", "Tecnología", "Aliados", "Capacidades"]
categorias = ["Núcleo integrado\n(dentro de la organización)", "Servicios digitales\n(Un tercero externo)",
              "Marketplace digital\n(Un mercado transaccional)", "Ecosistema digital\n(En la red mediante alianzas colaborativas)"]

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
header_fill_blanco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
dim_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
header_font_negro = Font(color="000000", bold=True)
dim_font = Font(color="FFFFFF", bold=True)
alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Eliminar la segunda fila
ws.delete_rows(2)

# Aplicar estilo a los encabezados de las categorías (primera fila)
for cell in ws[1]:
    cell.fill = header_fill_blanco
    cell.font = header_font_negro
    cell.alignment = alignment

# Aplicar estilo a los encabezados de las dimensiones (primera columna)
for cell in ws['A']:
    if cell.row != 1:
        cell.fill = dim_fill
        cell.font = dim_font
        cell.alignment = alignment
    else:
        cell.value = "Dimensiones"  # Agregar el título "Dimensiones" a la primera celda
        cell.fill = dim_fill
        cell.font = dim_font
        cell.alignment = alignment

# Ajuste del alto de la primera fila (encabezados)
ws.row_dimensions[1].height = 60  # Aumentar el alto para acomodar el texto

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
    adjusted_width = min((max_length + 2), 20)
    ws.column_dimensions[column].width = adjusted_width

# Ajuste del alto de las filas
for row in ws.iter_rows(min_row=2):
    ws.row_dimensions[row[0].row].height = 30  # Ajusta este valor según sea necesario

# Aplicar bordes a todas las celdas
medium_border = Border(left=Side(style='medium', color="000000"),
                       right=Side(style='medium', color="000000"),
                       top=Side(style='medium', color="000000"),
                       bottom=Side(style='medium', color="000000"))

for row in ws.iter_rows():
    for cell in row:
        cell.border = medium_border

# Guardar el archivo de Excel
output_file = "matriz_completada_formateada.xlsx"
wb.save(output_file)

# Crear o abrir el documento Word
output_word = "informe_matriz_con_tabla.docx"
try:
    doc = Document(output_word)  # Intentar abrir el archivo existente
except Exception:
    doc = Document()  # Crear un nuevo documento si no existe

# Agregar el título principal
doc.add_heading('Informe de Matriz Completada', level=0)

# Agregar un espacio de dos líneas antes de la tabla
doc.add_paragraph()  # Primer espacio
doc.add_paragraph()  # Segundo espacio

# Agregar el título de la tabla
doc.add_heading('Formato final', level=1)

# Leer el archivo de Excel y agregar la tabla al documento Word
excel_stream = BytesIO()
wb.save(excel_stream)
excel_stream.seek(0)

# Agregar la tabla de Excel al documento Word
df_table = pd.read_excel(excel_stream, sheet_name='Matriz Completada')

# Convertir el DataFrame en una tabla para el documento Word
table = doc.add_table(rows=1, cols=len(df_table.columns))

# Estilo de la tabla
table.style = 'Table Grid'  # Aplica un estilo de tabla con bordes

# Encabezados de la tabla
hdr_cells = table.rows[0].cells
for i, col_name in enumerate(df_table.columns):
    hdr_cells[i].text = col_name

# Filas de la tabla
for index, row in df_table.iterrows():
    row_cells = table.add_row().cells
    for i, value in enumerate(row):
        row_cells[i].text = str(value)

# Aplicar color azul a la primera columna y texto blanco y negrita
for row in table.rows:
    cell = row.cells[0]
    cell_fill = OxmlElement('w:shd')
    cell_fill.set(ns.qn('w:fill'), '4472C4')  # Azul
    cell._element.get_or_add_tcPr().append(cell_fill)
    
    cell_text = cell.paragraphs[0]
    for run in cell_text.runs:
        run.font.color.rgb = RGBColor(255, 255, 255)  # Blanco
        run.font.bold = True  # Negrita

# Aplicar bordes a la tabla
for row in table.rows:
    for cell in row.cells:
        cell_borders = OxmlElement('w:tcBorders')
        for border_name in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{border_name}')
            border.set(ns.qn('w:val'), 'single')
            border.set(ns.qn('w:space'), '0')
            border.set(ns.qn('w:sz'), '4')
            cell_borders.append(border)
        cell._element.get_or_add_tcPr().append(cell_borders)

# Agregar un espacio de dos líneas después de la tabla
doc.add_paragraph()  # Primer espacio
doc.add_paragraph()  # Segundo espacio

# Agregar el contenido adicional (si existe) después de la tabla
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
doc.save(output_word)

print(f"¡Informe actualizado con la tabla de Excel! El archivo Word '{output_word}' se ha guardado exitosamente.")
