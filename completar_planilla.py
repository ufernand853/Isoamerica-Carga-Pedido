import pandas as pd
from openpyxl import load_workbook

# === CONFIGURACIÓN ===
PEDIDO_FILE = "Planilla pedido 10.12.2025 Destino.xlsx"
LISTADO_FILE = "Listado general para PLANILLAS TRADU BRs.xlsx"
OUTPUT_FILE = "Planilla pedido 10.12.2025 Destino_COMPLETADA.xlsx"

# Nombre de hoja (None = primera)
PEDIDO_SHEET = None
LISTADO_SHEET = None

# Nombre de la columna clave en la planilla de pedido (columna E en tu ejemplo)
PEDIDO_KEY_COL = "Codigo Principal"

# Nombre de la columna clave en el listado general
LISTADO_KEY_COL = "Codigo"

# Mapeo desde columnas del LISTADO a columnas destino en la planilla de pedido
COLUMN_MAPPING = {
    # destino_en_pedido : origen_en_listado
    "EAN - Cod Barras": "EAN",                      # Columna D
    "Descrição": "Descripcion",                     # Columna F - traducción
    "Marca": "Fabricante",                          # Columna G (a falta de columna Marca, usamos Fabricante)
    "Pais de Origem": "Pais",                       # Columna H
    "NCM": "NCM",                                   # Columna I
    "Peso Neto Unitario": "Peso",                   # Columna R
    "Nome do Fabricante - Razão Social": "Fabricante",   # Columna W
    "Endereço do Fabricante - Rua - Numero - Cidade - Estado - CEP": "Ubicacion",  # Columna X
}

def cargar_listado(path, sheet_name=None):
    # Si no se especifica hoja, usamos la primera
    if sheet_name is None:
        listado = pd.read_excel(path, sheet_name=0)
    else:
        listado = pd.read_excel(path, sheet_name=sheet_name)

    # En tu archivo actual el listado tiene 9 columnas sin encabezados lógicos:
    # asumimos el orden: EAN, Codigo, Descripcion, Pais, NCM, Peso, Fabricante, Ubicacion, Extra
    if listado.shape[1] == 9:
        listado.columns = [
            "EAN",        # col 0
            "Codigo",     # col 1
            "Descripcion",# col 2
            "Pais",       # col 3
            "NCM",        # col 4
            "Peso",       # col 5
            "Fabricante", # col 6
            "Ubicacion",  # col 7
            "Extra",      # col 8 (no usada)
        ]
    return listado

def construir_diccionario(listado, key_col):
    """
    Construye un índice: clave -> fila (Series)
    para búsqueda rápida por código.
    """
    listado_key = listado.copy()
    listado_key[key_col] = listado_key[key_col].astype(str).str.strip()
    listado_key = listado_key.drop_duplicates(subset=key_col, keep="first")
    return listado_key.set_index(key_col)

def completar_planilla_pedido():
    # --- Cargar listado general ---
    listado = cargar_listado(LISTADO_FILE, LISTADO_SHEET)
    indexado = construir_diccionario(listado, LISTADO_KEY_COL)

    # --- Cargar planilla de pedido como DataFrame para identificar columnas ---
    # header=4 porque los encabezados reales están en la fila 5 de Excel
    if PEDIDO_SHEET is None:
        pedido_raw = pd.read_excel(PEDIDO_FILE, sheet_name=0, header=4)
    else:
        pedido_raw = pd.read_excel(PEDIDO_FILE, sheet_name=PEDIDO_SHEET, header=4)

    # La primera fila de pedido_raw contiene los nombres de las columnas
    header_row = pedido_raw.iloc[0]
    pedido_data = pedido_raw[1:].copy()
    pedido_data.columns = header_row

    # Normalizar la columna clave de pedido
    pedido_data[PEDIDO_KEY_COL] = pedido_data[PEDIDO_KEY_COL].astype(str).str.strip()

    # --- Abrir el Excel original con openpyxl para mantener formato ---
    wb = load_workbook(PEDIDO_FILE)
    ws = wb[wb.sheetnames[0]] if PEDIDO_SHEET is None else wb[PEDIDO_SHEET]

    # La primera fila de datos (Item 1) es la fila 6 de Excel.
    # En el DataFrame, la primera fila de datos tiene índice 1, por lo que:
    # fila_excel = índice_df + 5
    total_rows = 0
    matched_rows = 0

    for idx, row in pedido_data.iterrows():
        codigo = str(row.get(PEDIDO_KEY_COL, "")).strip()
        if not codigo or codigo.lower() == "nan":
            continue

        total_rows += 1

        if codigo in indexado.index:
            matched_rows += 1
            fuente = indexado.loc[codigo]
            excel_row = idx + 5  # ver comentario arriba

            for destino_col, origen_col in COLUMN_MAPPING.items():
                if origen_col not in fuente.index:
                    continue

                valor = fuente[origen_col]

                if destino_col not in pedido_data.columns:
                    continue

                col_idx = list(pedido_data.columns).index(destino_col)
                excel_col = col_idx + 1  # A=1, B=2, etc.

                ws.cell(row=excel_row, column=excel_col, value=valor)

    wb.save(OUTPUT_FILE)

    print(f"Filas procesadas (con código en la columna E): {total_rows}")
    print(f"Filas con coincidencia en el listado: {matched_rows}")
    print(f"Archivo generado: {OUTPUT_FILE}")

if __name__ == "__main__":
    completar_planilla_pedido()
