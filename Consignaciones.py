import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import os
import zipfile
import re
import math  # <-- Para usar math.ceil en el ajuste dinámico

st.title("Liquidaciones de Consignaciones")

# 1) Subir el archivo Excel maestro
uploaded_file = st.file_uploader("Sube el archivo Excel maestro (.xlsx)", type=["xlsx"])

def load_default_logo():
    """
    Carga el logo por defecto (logo.png) si existe en la misma carpeta.
    Retorna un BytesIO con los datos de la imagen o None si no se encuentra.
    """
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(current_dir, "logo.png")
        if os.path.exists(logo_path):
            with open(logo_path, "rb") as f:
                return BytesIO(f.read())
    except Exception:
        pass
    return None

def create_export_excel(df, editorial, logo_content=None):
    """
    Genera el archivo Excel para la editorial dada, con el formato solicitado.
    
    Parámetros:
    - df: DataFrame con columnas [Unidades a liquidar, Producto, ISBN].
    - editorial: string con el nombre de la editorial en MAYÚSCULAS.
    - logo_content: contenido en bytes del logo (no un BytesIO abierto).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Liquidación"
    ws.sheet_view.showGridLines = False

    # Estilos
    title_font = Font(name="Arial", size=16, bold=True)
    header_font = Font(name="Arial", size=11, bold=True)
    normal_font = Font(name="Arial", size=10)
    bold_font = Font(name="Arial", size=10, bold=True)
    thin_border = Border(
        left=Side(style="thin"), 
        right=Side(style="thin"),
        top=Side(style="thin"), 
        bottom=Side(style="thin")
    )
    
    # Ajustar altura de la primera fila para dar espacio al logo
    ws.row_dimensions[1].height = 45

    # Insertar logo (si se proporcionó)
    if logo_content is not None:
        try:
            logo_bytes = BytesIO(logo_content)
            img = OpenpyxlImage(logo_bytes)
            img.width = 80
            img.height = 50
            ws.add_image(img, "A1")
        except Exception:
            pass

    # Título principal, en mayúsculas
    ws.merge_cells("B1:D2")
    cell_title = ws["B1"]
    title_text = f"LIQUIDACION CONSIGNACIONES {editorial}"
    cell_title.value = title_text
    cell_title.font = title_font
    cell_title.alignment = Alignment(horizontal="center", vertical="center")

    # Información de cliente
    ws.merge_cells("B3:D6")
    cell_cliente = ws["B3"]
    cell_cliente.value = (
        "CLIENTE: Librería Virtual y Distribuidora El Ático Ltda.\n"
        "Venta y Distribución de Libros\n"
        "General Bari 234, Providencia - Santiago, Teléfono: (56)2 21452308\n"
        "Rut: 70.082.998-0"
    )
    cell_cliente.font = normal_font
    cell_cliente.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center")

    # Tabla de proveedor/contacto
    fields = ["PROVEEDOR:", "CONTACTO:", "FONO / MAIL:", "DESCUENTO:", "PAGO:", "FECHA:"]
    row_start = 8
    for i, field in enumerate(fields):
        row_i = row_start + i
        ws.cell(row=row_i, column=2, value=field).font = header_font
        # Fusionar celdas en columnas C y D para el valor
        ws.merge_cells(start_row=row_i, start_column=3, end_row=row_i, end_column=4)
        merged_cell = ws.cell(row=row_i, column=3)
        merged_cell.value = ""
        merged_cell.font = normal_font
        # Aplicar bordes
        for c in range(2, 5):
            ws.cell(row=row_i, column=c).border = thin_border

    # Encabezados de la tabla principal (a partir de fila 16, columna B)
    start_row = 16
    start_col = 2
    for offset, header in enumerate(df.columns):
        col_index = start_col + offset
        cell = ws.cell(row=start_row, column=col_index, value=header)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        # Asignar un ancho fijo "base" para cada columna de la tabla
        ws.column_dimensions[get_column_letter(col_index)].width = 20

    # Datos de la tabla principal
    row_counter = start_row
    for row_data in df.itertuples(index=False):
        row_counter += 1
        for offset, value in enumerate(row_data):
            col_index = start_col + offset
            cell = ws.cell(row=row_counter, column=col_index, value=value)
            cell.font = normal_font
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = thin_border

    # Observaciones
    obs_row = row_counter + 2
    ws.merge_cells(start_row=obs_row, start_column=2, end_row=obs_row+3, end_column=4)
    obs_cell = ws.cell(row=obs_row, column=2)
    obs_cell.value = (
        "OBSERVACIONES:\n"
        "1.- DESPACHAR A GENERAL BARI 234, PROVIDENCIA, SANTIAGO.\n"
        "2.- HORARIO RECEPCION DE PEDIDOS: LUNES A VIERNES 09:30 A 13:00 Y 16:00 A 18:30"
    )
    obs_cell.font = bold_font
    obs_cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Bordes en el bloque de observaciones
    for r in range(obs_row, obs_row+4):
        for c in range(2, 5):
            ws.cell(row=r, column=c).border = thin_border

    # Ocultar columnas F en adelante
    max_col = ws.max_column
    if max_col > 5:
        ws.delete_cols(6, max_col - 5)
    for col in range(6, 16385):
        ws.column_dimensions[get_column_letter(col)].hidden = True

    # --- Ajuste dinámico aproximado para el título ---
    # Medimos la longitud del texto completo y calculamos un ancho total.
    # Luego, repartimos ese ancho entre las columnas B, C y D.
    title_length = len(title_text)
    approx_total_width = 2 * title_length  # factor 2 para un ajuste "cómodo"
    width_per_col = math.ceil(approx_total_width / 3)

    for col in [2, 3, 4]:  # Columnas B, C y D
        current_width = ws.column_dimensions[get_column_letter(col)].width
        # Asignamos el máximo entre el ancho actual y el nuevo calculado
        ws.column_dimensions[get_column_letter(col)].width = max(current_width, width_per_col)

    # Guardar en BytesIO
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

def process_master_file(file_bytes, logo_content=None):
    """
    Lee el archivo maestro y genera un diccionario con {nombre_excel: contenido_en_bytes}
    para cada editorial detectada, excluyendo stock negativo en BODEGA GENERAL BARI
    y calculando Unidades a liquidar > 0.
    """
    # Leer la hoja principal, fila 6 como encabezado
    df = pd.read_excel(file_bytes, sheet_name=0, header=5)
    df.columns = df.columns.str.strip()
    
    # Renombrar "Código" a "Codigo" si existe
    if "Código" in df.columns:
        df.rename(columns={"Código": "Codigo"}, inplace=True)

    # Detectar columnas que contengan "consignacion" o "consignaciones" (ignorar mayúsculas)
    all_cols = df.columns.tolist()
    consign_cols = [col for col in all_cols if re.search(r'consignacion', col, re.IGNORECASE)]
    
    output_files = {}
    for col in consign_cols:
        # 1) Remover "CONSIGNACION" o "CONSIGNACIONES"
        editorial_name = re.sub(r'(?i)consignacion(es)?', '', col)
        # 2) Reemplazar múltiples espacios por uno solo
        editorial_name = re.sub(r'\s+', ' ', editorial_name)
        # 3) Quitar : y otros caracteres sobrantes
        editorial_name = re.sub(r'[:]+', '', editorial_name)
        # 4) Eliminar secuencias de dígitos (y guiones) para no incluir RUT ni similares
        editorial_name = re.sub(r'[0-9-]+', '', editorial_name)
        # 5) Pasar a mayúsculas y hacer strip final
        editorial_name = editorial_name.strip().upper()
        # Si no queda nada, poner "SIN EDITORIAL"
        if not editorial_name:
            editorial_name = "SIN EDITORIAL"

        required_cols = ["Producto", "Codigo", "BODEGA GENERAL BARI"]
        if not all(x in df.columns for x in required_cols):
            st.error("Faltan columnas requeridas en el archivo maestro.")
            return {}
        
        # Filtrar las columnas que necesitamos
        temp_df = df[["Producto", "Codigo", "BODEGA GENERAL BARI", col]].copy()
        temp_df.rename(columns={col: "Consignaciones"}, inplace=True)
        
        # Excluir stock negativo en BODEGA GENERAL BARI
        temp_df = temp_df[temp_df["BODEGA GENERAL BARI"] >= 0]

        # Calcular Unidades a liquidar
        temp_df["Unidades a liquidar"] = temp_df["Consignaciones"] - temp_df["BODEGA GENERAL BARI"]
        temp_df = temp_df[temp_df["Unidades a liquidar"] > 0]
        temp_df = temp_df.sort_values(by="Producto")

        if temp_df.empty:
            continue
        
        # Prepara DataFrame de exportación
        export_df = temp_df[["Unidades a liquidar", "Producto", "Codigo"]].copy()
        export_df.rename(columns={"Codigo": "ISBN"}, inplace=True)
        # Limitar ISBN a 13 caracteres antes de la barra
        export_df["ISBN"] = export_df["ISBN"].astype(str).apply(lambda x: x.split("/")[0][:13])
        
        # Generar Excel con el formato
        excel_bytes = create_export_excel(export_df, editorial_name, logo_content)
        filename = f"Liquidacion_Consignaciones_{editorial_name}.xlsx"
        output_files[filename] = excel_bytes

    return output_files

if uploaded_file is not None:
    try:
        # Leer el Excel en memoria
        file_bytes = BytesIO(uploaded_file.read())
        
        # Cargar el logo por defecto (si existe)
        default_logo = load_default_logo()
        logo_data = default_logo.read() if default_logo else None

        results = process_master_file(file_bytes, logo_data)

        if results:
            st.success("Liquidaciones generadas para las siguientes editoriales:")
            for name in results.keys():
                st.write("📄", name)
            # Empaquetar en ZIP
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for filename, content in results.items():
                    zip_file.writestr(filename, content)
            zip_buffer.seek(0)
            st.download_button(
                label="📦 Descargar ZIP con todas las liquidaciones",
                data=zip_buffer,
                file_name="Liquidaciones.zip",
                mime="application/zip"
            )
        else:
            st.error("No se generaron liquidaciones (posiblemente no hay registros con unidades a liquidar > 0).")
    except Exception as e:
        st.error(f"Error al procesar el archivo maestro: {e}")
