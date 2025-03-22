import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import os

st.title("Control de Stock y Consignaciones")

# Subir el archivo Excel con el stock
uploaded_file = st.file_uploader("Sube el archivo Excel (.xlsx)", type=["xlsx"])

# Opción para subir un logo diferente (opcional)
logo_file = st.file_uploader("Sube un logo de la empresa (opcional)", type=["png", "jpg", "jpeg"])

def load_default_logo():
    """
    Carga el logo por defecto desde 'logo.png', ubicado en la misma carpeta que este script.
    Ajustamos la ruta absoluta para evitar problemas de localización.
    """
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(current_dir, "logo.png")
        if os.path.exists(logo_path):
            with open(logo_path, "rb") as f:
                return BytesIO(f.read())
        else:
            st.warning("Logo por defecto no encontrado en la carpeta de la aplicación.")
    except Exception as e:
        st.warning(f"No se pudo cargar el logo por defecto: {e}")
    return None

if uploaded_file is not None:
    try:
        # 1. LECTURA Y FILTRADO DE DATOS
        df = pd.read_excel(uploaded_file, sheet_name="Stock actual", header=5)
        df.columns = df.columns.str.strip()
        
        # Renombrar columnas para homogeneizar
        df.rename(columns={
            "Código": "Codigo",
            "CONSIGNACIONES SM": "Consignaciones"
        }, inplace=True)
        
        # Filtrar columnas necesarias
        columnas_necesarias = ["Producto", "Codigo", "BODEGA GENERAL BARI", "Consignaciones"]
        df_filtrado = df[columnas_necesarias].copy()
        
        # Calcular "Unidades a liquidar"
        df_filtrado["Unidades a liquidar"] = df_filtrado["Consignaciones"] - df_filtrado["BODEGA GENERAL BARI"]
        
        # Filtrar filas con "Unidades a liquidar" > 0 (estrictamente mayor que 0)
        df_filtrado = df_filtrado[df_filtrado["Unidades a liquidar"] > 0]
        
        # Ordenar alfabéticamente por "Producto"
        df_filtrado = df_filtrado.sort_values(by="Producto")
        
        # Mostrar datos filtrados en la app
        st.subheader("Datos filtrados")
        st.dataframe(df_filtrado)
        
        # Preparar datos para exportar
        df_export = df_filtrado[["Unidades a liquidar", "Producto", "Codigo"]].copy()
        df_export.rename(columns={"Codigo": "ISBN"}, inplace=True)
        
        # Ajustar el ISBN para que sean solo los primeros 13 caracteres antes de la barra
        df_export["ISBN"] = df_export["ISBN"].astype(str).apply(lambda x: x.split("/")[0][:13])
        
        # 2. FUNCIÓN PARA CREAR EL EXCEL CON FORMATO
        def create_export_excel(df, logo_bytes=None):
            """
            Crea un archivo Excel con el siguiente formato:
              - Oculta las líneas de cuadrícula (showGridLines=False)
              - Título principal en celdas B1:D2
              - Debajo del título, info fija de "CLIENTE" (centrada) en B3:D6
              - Tabla de proveedor/contacto (B8:D13), con celdas C-D fusionadas y bordes
              - Tabla principal (Unidades a liquidar, Producto, ISBN) desde la fila 16, columna B
              - Bloque de Observaciones (texto fijo en negrita) al final
              - Elimina/oculta las columnas F->XFD para que solo aparezcan A->E
            """
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Liquidación"
            
            # Ocultar cuadrículas
            ws.sheet_view.showGridLines = False
            
            # ----------------------------
            # ESTILOS Y CONFIGURACIONES
            # ----------------------------
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
            
            # Ajustar la altura de la primera fila para dar espacio al logo
            ws.row_dimensions[1].height = 45
            
            # ----------------------------
            # LOGO (OPCIONAL)
            # ----------------------------
            if logo_bytes is not None:
                try:
                    img = OpenpyxlImage(logo_bytes)
                    img.width = 80
                    img.height = 50
                    ws.add_image(img, "A1")
                except Exception as e:
                    st.warning(f"No se pudo insertar el logo: {e}")
            
            # ----------------------------
            # TÍTULO PRINCIPAL (B1:D2)
            # ----------------------------
            ws.merge_cells("B1:D2")
            cell_title = ws["B1"]
            cell_title.value = "LIQUIDACION CONSIGNACIONES SM"
            cell_title.font = title_font
            cell_title.alignment = Alignment(horizontal="center", vertical="center")
            
            # ----------------------------
            # INFORMACIÓN DE CLIENTE (B3:D6), centrado
            # ----------------------------
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
            
            # ----------------------------
            # TABLA PROVEEDOR/CONTACTO (B8:D13)
            # ----------------------------
            fields = ["PROVEEDOR:", "CONTACTO:", "FONO / MAIL:", "DESCUENTO:", "PAGO:", "FECHA:"]
            row_start = 8
            for i, field in enumerate(fields):
                row_i = row_start + i
                # Etiqueta en columna B
                cell_label = ws.cell(row=row_i, column=2, value=field)
                cell_label.font = header_font
                cell_label.alignment = Alignment(horizontal="left", vertical="center")
                
                # Fusionar celdas en columnas C y D para el valor
                ws.merge_cells(start_row=row_i, start_column=3, end_row=row_i, end_column=4)
                merged_cell = ws.cell(row=row_i, column=3)
                merged_cell.value = ""
                merged_cell.font = normal_font
                merged_cell.alignment = Alignment(horizontal="left", vertical="center")
            
            # Agregar bordes para cerrar el rectángulo B8:D13
            for r in range(row_start, row_start + len(fields)):
                for c in range(2, 5):  # columnas B(2) a D(4)
                    ws.cell(row=r, column=c).border = thin_border
            
            # ----------------------------
            # TABLA PRINCIPAL (a partir de fila 16, columna B)
            # ----------------------------
            start_row = 16
            start_col = 2
            headers = list(df.columns)  # ["Unidades a liquidar", "Producto", "ISBN"]
            
            # Encabezados
            for offset, header in enumerate(headers):
                col_index = start_col + offset
                cell = ws.cell(row=start_row, column=col_index, value=header)
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
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
            
            # ----------------------------
            # BLOQUE DE OBSERVACIONES (en negrita)
            # ----------------------------
            obs_row = row_counter + 2
            ws.merge_cells(start_row=obs_row, start_column=2, end_row=obs_row+3, end_column=4)
            obs_cell = ws.cell(row=obs_row, column=2)
            observaciones_text = (
                "OBSERVACIONES:\n"
                "1.- DESPACHAR A GENERAL BARI 234, PROVIDENCIA, SANTIAGO.\n"
                "2.- HORARIO RECEPCION DE PEDIDOS: LUNES A VIERNES 09:30 A 13:00 Y 16:00 A 18:30"
            )
            obs_cell.value = observaciones_text
            obs_cell.font = bold_font
            obs_cell.alignment = Alignment(wrap_text=True, vertical="top")
            
            for r in range(obs_row, obs_row+4):
                for c in range(2, 5):
                    ws.cell(row=r, column=c).border = thin_border
            
            # ----------------------------
            # ELIMINAR / OCULTAR COLUMNAS F->XFD
            # ----------------------------
            # 1) Eliminar columnas que openpyxl considera "existentes" (si max_column > 5)
            max_col = ws.max_column
            if max_col > 5:
                ws.delete_cols(6, max_col - 5)
            
            # 2) Ocultar columnas 6->16384 (F->XFD), para que no se vean al abrir
            for col in range(6, 16385):
                ws.column_dimensions[get_column_letter(col)].hidden = True
            
            # Guardar en un objeto BytesIO y retornar el contenido
            output = BytesIO()
            wb.save(output)
            return output.getvalue()
        
        # 3. BOTÓN PARA EXPORTAR
        if st.button("Exportar a Excel con Formato"):
            if logo_file is not None:
                logo_bytes = BytesIO(logo_file.read())
                logo_file.seek(0)
            else:
                logo_bytes = load_default_logo()
            
            excel_data = create_export_excel(df_export, logo_bytes)
            st.download_button(
                label="Descargar Excel",
                data=excel_data,
                file_name="Liquidacion_Consignaciones.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error("Error al procesar el archivo: " + str(e))
