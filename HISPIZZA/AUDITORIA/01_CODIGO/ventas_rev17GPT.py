import pandas as pd
import openpyxl
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime

# Global variables for DataFrames
balanza_soft_df = None
diccionario_df = None
ordenes_ped_ya_df = None
ordenes_pwr_df = None

# Function to load each file
def cargar_archivo(tipo):
    global balanza_soft_df, diccionario_df, ordenes_ped_ya_df, ordenes_pwr_df

    filetypes = [("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]

    def load_excel(path, skip_rows=0):
        return pd.read_excel(path, engine='openpyxl', skiprows=skip_rows)

    file_type_titles = {
        'balanza_soft': "Selecciona Balanza de Softland",
        'diccionario': "Selecciona Registro de Sucursales",
        'ordenes_ped_ya': "Selecciona Ordenes de Pedidos Ya",
        'ordenes_pwr': "Selecciona Ordenes de PWR"
    }
    skip_rows = {'balanza_soft': 8, 'ordenes_ped_ya': 1} # ¿?

    if tipo in file_type_titles:
        path = filedialog.askopenfilename(title=file_type_titles[tipo], filetypes=filetypes)
        if path:
            try:
                skip = skip_rows.get(tipo, 0)
                df = load_excel(path, skip)
                if df.empty:
                    raise ValueError(f"El archivo de {file_type_titles[tipo]} está vacío")
                if tipo == 'balanza_soft':
                    balanza_soft_df = df
                elif tipo == 'diccionario':
                    diccionario_df = df
                elif tipo == 'ordenes_ped_ya':
                    ordenes_ped_ya_df = df
                elif tipo == 'ordenes_pwr':
                    ordenes_pwr_df = df
                messagebox.showinfo("Éxito", f"Archivo de {file_type_titles[tipo]} cargado correctamente")
            except Exception as e:
                messagebox.showerror("Error", f"Error al cargar archivo de {file_type_titles[tipo]}: {e}")

# Function to check if all files are loaded
def verificar_archivos():
    if all([balanza_soft_df is not None, diccionario_df is not None, ordenes_ped_ya_df is not None, ordenes_pwr_df is not None]):
        messagebox.showinfo("Éxito", "Todos los archivos se han cargado correctamente")
    else:
        messagebox.showwarning("Advertencia", "Debe cargar todos los archivos")

# Function to process the data and save the file
def generar_archivo():
    global balanza_soft_df, diccionario_df, ordenes_ped_ya_df, ordenes_pwr_df

    if not all([balanza_soft_df is not None, diccionario_df is not None, ordenes_ped_ya_df is not None, ordenes_pwr_df is not None]):
        messagebox.showwarning("Advertencia", "Debe cargar todos los archivos antes de procesar")
        return

    # Ensure columns for merging are string type
    def ensure_str_columns(df, columns):
        for column in columns:
            df[column] = df[column].astype(str)
        return df

    diccionario_df = ensure_str_columns(diccionario_df, ['NUMERO'])
    ordenes_ped_ya_df = ensure_str_columns(ordenes_ped_ya_df, ['Restaurant name'])
    ordenes_pwr_df = ensure_str_columns(ordenes_pwr_df, ['Store'])

    # Filter rows starting with "PedidosYa" in 'Referencia' column
    n_balance_df = balanza_soft_df[balanza_soft_df['Referencia'].str.startswith('PedidosYa', na=False)].copy()

    # Split 'Referencia' column
    split_df = n_balance_df['Referencia'].str.split(' ', expand=True)
    split_df.columns = ['REFERENCIA', 'TIENDA']

    # Concatenate and drop unnecessary columns
    n_balance_df = pd.concat([n_balance_df, split_df], axis=1)
    n_balance_df.drop(columns=['Centro Costo', 'Referencia', 'Fuente', 'Tipo Asiento', 'Origen', 'Unnamed: 8', 'Unnamed: 9'], inplace=True)

    # Rename columns
    n_balance_df.rename(columns={
        'Fecha': 'FECHA_CUADRE',
        'Asiento': 'ASIENTO',
        'Montos': 'MONTOS_CUADRE'
    }, inplace=True)

    # Perform merge and filter data
    diccionario_df['NUMERO'] = diccionario_df['NUMERO'].astype(str)
    n_ord_ped_df = ordenes_ped_ya_df.merge(diccionario_df[['NOMBRE', 'NUMERO']], how='left', left_on='Restaurant name', right_on='NOMBRE')

    # Identify rows without a match in the merge
    sin_coincidencia_df = n_ord_ped_df[n_ord_ped_df['NUMERO'].isnull()]

    if not sin_coincidencia_df.empty:
        no_match_restaurants_ped_ya = sin_coincidencia_df['Restaurant name'].unique()
        no_match_restaurants_diccionario = diccionario_df['NOMBRE'].unique()
        no_match_message = (
            "Las siguientes órdenes en ordenes_ped_ya_df no tienen una coincidencia en diccionario_df:\n"
            + "\n".join(no_match_restaurants_ped_ya)
            + "\n\nNombres de restaurantes en diccionario_df:\n"
            + "\n".join(no_match_restaurants_diccionario)
        )
        messagebox.showwarning("Advertencia", no_match_message)

    n_ord_ped_df.dropna(subset=['NUMERO'], inplace=True)
    n_ord_ped_df['NUMERO'] = n_ord_ped_df['NUMERO'].astype(str)

    # Create a new DataFrame with specified columns and drop rows with NaN in 'Cancellation reason'
    ordenes_canceladas_df = n_ord_ped_df[['Subtotal', 'Restaurant name', 'NUMERO', 'Order ID', 'Order received at', 'Cancellation reason', 'Cancellation owner']].copy()
    ordenes_canceladas_df.dropna(subset=['Cancellation reason'], inplace=True)
    ordenes_canceladas_df.reset_index(drop=True, inplace=True)

    # Filter out canceled orders from n_ord_ped_df
    n_ord_ped_df = n_ord_ped_df[n_ord_ped_df['Order status'] != 'Cancelled']

    # Create cuadre DataFrame
    cuadre_df = n_balance_df[['FECHA_CUADRE', 'TIENDA', 'MONTOS_CUADRE']].copy()
    cuadre_df['FECHA_CUADRE'] = pd.to_datetime(cuadre_df['FECHA_CUADRE'], errors='coerce').dt.date
    cuadre_df['TIENDA'] = cuadre_df['TIENDA'].astype(str)

    # Add Subtotal by NUMERO and FECHA_PEDIDOSYA
    aggregated_df = n_ord_ped_df.groupby(['NUMERO', pd.to_datetime(n_ord_ped_df['Order received at']).dt.date])['Subtotal'].sum().reset_index()
    aggregated_df.rename(columns={
        'Order received at': 'FECHA_PEDIDOSYA',
        'Subtotal': 'MONTOS_PEDIDOSYA'
    }, inplace=True)
    aggregated_df['FECHA_PEDIDOSYA'] = pd.to_datetime(aggregated_df['FECHA_PEDIDOSYA'], errors='coerce').dt.date
    aggregated_df['NUMERO'] = aggregated_df['NUMERO'].astype(str)

    # Merge DataFrames
    merged_df = pd.merge(cuadre_df, aggregated_df, how='left', left_on=['FECHA_CUADRE', 'TIENDA'], right_on=['FECHA_PEDIDOSYA', 'NUMERO'])

    # Rename columns for clarity
    merged_df.rename(columns={
        'TIENDA': 'TIENDA_CUADRE',
        'NUMERO': 'TIENDA_PEDIDOSYA',
        'FECHA_PEDIDOSYA': 'FECHA_PEDIDOSYA',
        'MONTOS_PEDIDOSYA': 'MONTOS_PEDIDOSYA'
    }, inplace=True)

    # Reorder columns
    merged_df = merged_df[['FECHA_CUADRE', 'TIENDA_CUADRE', 'MONTOS_CUADRE', 'FECHA_PEDIDOSYA', 'TIENDA_PEDIDOSYA', 'MONTOS_PEDIDOSYA']]

    # Process PWR orders
    ordenes_pwr_df = ordenes_pwr_df.rename(columns=lambda x: x.strip())
    ordenes_pwr_df_h = ordenes_pwr_df[ordenes_pwr_df['Type'] == 'H']
    n_ordenes_pwr_df = ordenes_pwr_df_h[ordenes_pwr_df_h['Order Status'] == 'Completed'].copy()
    n_ordenes_pwr_df.loc[:, 'Order Date / Time'] = pd.to_datetime(n_ordenes_pwr_df['Order Date / Time'])
    n_ordenes_pwr_df.loc[:, 'Fecha'] = n_ordenes_pwr_df['Order Date / Time'].dt.date
    n_ordenes_pwr_df.loc[:, 'Hora'] = n_ordenes_pwr_df['Order Date / Time'].dt.time
    n_ordenes_pwr_df.reset_index(drop=True, inplace=True)

    # Convert to string to avoid issues
    n_ordenes_pwr_df['Store'] = n_ordenes_pwr_df['Store'].astype(str)
    n_ord_ped_df['NUMERO'] = n_ord_ped_df['NUMERO'].astype(str)
    n_ord_ped_df['Order received at'] = pd.to_datetime(n_ord_ped_df['Order received at'])
    n_ord_ped_df['Fecha'] = n_ord_ped_df['Order received at'].dt.date
    n_ord_ped_df['Hora'] = n_ord_ped_df['Order received at'].dt.time

    # PWR analysis
    analisis_pwr_df = n_ordenes_pwr_df[['Store', 'Order Number', 'Order Date / Time', 'Type', 'Order Status', 'Customer Name', 'Order Phone No.', 'Final Order Price', 'Comments', 'Fecha', 'Hora']].copy()
    analisis_pwr_df.reset_index(drop=True, inplace=True)
    analisis_pwr_df.sort_values(by=['Store', 'Fecha', 'Hora'], inplace=True)

    # PedidosYa analysis
    analisis_pedidos_ya_df = n_ord_ped_df[['Subtotal', 'Restaurant name', 'NUMERO', 'Order ID', 'Order received at', 'Fecha', 'Hora']].copy()
    analisis_pedidos_ya_df.sort_values(by=['NUMERO', 'Fecha', 'Hora'], inplace=True)
    analisis_pedidos_ya_df.reset_index(drop=True, inplace=True)

    # Rename and reorder columns
    merged_df = merged_df.rename(columns={'FECHA_CUADRE': 'FECHA_SOFT', 'TIENDA_CUADRE': 'TIENDA_SOFT', 'MONTOS_CUADRE': 'MONTOS_SOFT'})
    ordenes_canceladas_df = ordenes_canceladas_df.rename(columns={'Subtotal': 'SUBTOTAL', 'Restaurant name': 'NOMBRE_RESTAURANTE', 'NUMERO': 'NUMERO_RESTAURANTE', 'Order ID': 'NUMERO_ID', 'Order received at': 'FECHA_HORA_RESTAURANTE', 'Cancellation reason': 'RAZON_CANCELACION', 'Cancellation owner': 'PERSONAL_CANCELACION'})
    ordenes_canceladas_df = ordenes_canceladas_df.reindex(columns=['NUMERO_RESTAURANTE', 'NOMBRE_RESTAURANTE', 'FECHA_HORA_RESTAURANTE', 'NUMERO_ID', 'SUBTOTAL', 'RAZON_CANCELACION', 'PERSONAL_CANCELACION'])
    analisis_pwr_df = analisis_pwr_df.rename(columns={'Store': 'TIENDA_PWR', 'Order Number': 'NUMERO_ORDEN_PWR', 'Order Date / Time': 'FECHA_HORA_PWR', 'Type': 'TIPO_ORDEN_PWR', 'Order Status': 'STATUS_ORDEN_PWR', 'Customer Name': 'NOMBRE_CLT_PWR', 'Order Phone No.': 'TELEFONO_PWR', 'Final Order Price': 'PRECIO_PWR', 'Comments': 'COMENTARIOS', 'Fecha': 'FECHA_PWR', 'Hora': 'HORA_PWR'})
    analisis_pwr_df = analisis_pwr_df.reindex(columns=['TIENDA_PWR', 'FECHA_HORA_PWR', 'FECHA_PWR', 'HORA_PWR', 'TIPO_ORDEN_PWR', 'STATUS_ORDEN_PWR', 'NUMERO_ORDEN_PWR', 'NOMBRE_CLT_PWR', 'TELEFONO_PWR', 'PRECIO_PWR', 'COMENTARIOS'])
    analisis_pedidos_ya_df = analisis_pedidos_ya_df.rename(columns={'Subtotal': 'PRECIO_PED_YA', 'Restaurant name': 'NOMBRE_TIENDA_PED_YA', 'NUMERO': 'TIENDA_PED_YA', 'Order ID': 'NUMERO_ORDEN_PED_YA', 'Order received at': 'FECHA_HORA_PED_YA', 'Fecha': 'FECHA_PED_YA', 'Hora': 'HORA_PED_YA'})
    analisis_pedidos_ya_df = analisis_pedidos_ya_df.reindex(columns=['TIENDA_PED_YA', 'NOMBRE_TIENDA_PED_YA', 'FECHA_HORA_PED_YA', 'FECHA_PED_YA', 'HORA_PED_YA', 'NUMERO_ORDEN_PED_YA', 'PRECIO_PED_YA'])

    # Reset indexes of DataFrames
    merged_df.reset_index(drop=True, inplace=True)
    ordenes_canceladas_df.reset_index(drop=True, inplace=True)

    # Function to add counter with custom columns
    def agregar_contador(df, columna_tienda, columna_fecha, columna_orden, nombre_columna_contador):
        df = df.sort_values(by=[columna_tienda, columna_fecha, columna_orden])
        df[nombre_columna_contador] = df.groupby([columna_tienda, columna_fecha]).cumcount() + 1
        return df

    # Function to join columns of STORE, DATE and COUNTER
    def unir_columnas(df, columna_tienda, columna_fecha, columna_contador, nueva_columna):
        df[nueva_columna] = df[columna_tienda].astype(str) + '-' + df[columna_fecha].astype(str) + '-' + df[columna_contador].astype(str)
        return df

    analisis_pedidos_ya_df = agregar_contador(analisis_pedidos_ya_df, 'TIENDA_PED_YA', 'FECHA_PED_YA', 'NUMERO_ORDEN_PED_YA', 'CONTADOR_PED_YA')
    analisis_pwr_df = agregar_contador(analisis_pwr_df, 'TIENDA_PWR', 'FECHA_PWR', 'NUMERO_ORDEN_PWR', 'CONTADOR_PWR')

    # Fill NaN with 0 before converting to integer
    analisis_pedidos_ya_df['CONTADOR_PED_YA'] = analisis_pedidos_ya_df['CONTADOR_PED_YA'].fillna(0).astype(int)
    analisis_pwr_df['CONTADOR_PWR'] = analisis_pwr_df['CONTADOR_PWR'].fillna(0).astype(int)

    analisis_pedidos_ya_df = unir_columnas(analisis_pedidos_ya_df, 'TIENDA_PED_YA', 'FECHA_PED_YA', 'CONTADOR_PED_YA', 'COD_INTERNO_PED_YA')
    analisis_pwr_df = unir_columnas(analisis_pwr_df, 'TIENDA_PWR', 'FECHA_PWR', 'CONTADOR_PWR', 'COD_INTERNO_PWR')

    analisis_cuadre_df = pd.merge(analisis_pedidos_ya_df, analisis_pwr_df, left_on='COD_INTERNO_PED_YA', right_on='COD_INTERNO_PWR', how='outer')

    # Create the DIFERENCIA_DESCUENTO column considering NaN as 0
    analisis_cuadre_df['DIFERENCIA_DESCUENTO'] = (analisis_cuadre_df['PRECIO_PWR'].fillna(0) - analisis_cuadre_df['PRECIO_PED_YA'].fillna(0)).infer_objects(copy=False)

    # Create the columns TIENDA_UNICA and FECHA_UNICA
    analisis_cuadre_df['TIENDA_UNICA'] = analisis_cuadre_df['TIENDA_PWR'].combine_first(analisis_cuadre_df['TIENDA_PED_YA'])
    analisis_cuadre_df['FECHA_UNICA'] = analisis_cuadre_df['FECHA_PWR'].combine_first(analisis_cuadre_df['FECHA_PED_YA'])

    # Sum DIFERENCIA_DESCUENTO by TIENDA_UNICA and FECHA_UNICA
    diferencia_descuentos_totales = analisis_cuadre_df.groupby(['TIENDA_UNICA', 'FECHA_UNICA'])['DIFERENCIA_DESCUENTO'].sum().reset_index()

    # Rename the columns to merge with merged_df
    diferencia_descuentos_totales.rename(columns={
        'TIENDA_UNICA': 'TIENDA_SOFT',
        'FECHA_UNICA': 'FECHA_SOFT',
        'DIFERENCIA_DESCUENTO': 'DIFERENCIA_DESCUENTOS_TOTAL'
    }, inplace=True)

    # Convert dates to datetime type to ensure they match
    merged_df['FECHA_SOFT'] = pd.to_datetime(merged_df['FECHA_SOFT'], errors='coerce').dt.date
    diferencia_descuentos_totales['FECHA_SOFT'] = pd.to_datetime(diferencia_descuentos_totales['FECHA_SOFT'], errors='coerce').dt.date

    # Merge to add DIFERENCIA_DESCUENTOS_TOTAL to merged_df
    merged_df = pd.merge(merged_df, diferencia_descuentos_totales, on=['TIENDA_SOFT', 'FECHA_SOFT'], how='left')

    # Create the DIFERENCIA_REAL column
    merged_df['DIFERENCIA_REAL'] = (merged_df['MONTOS_SOFT'] - merged_df['MONTOS_PEDIDOSYA'] - merged_df['DIFERENCIA_DESCUENTOS_TOTAL']).round(2)

    # Create a new CONTADOR_UNICO column for correct ordering
    analisis_cuadre_df['CONTADOR_UNICO'] = analisis_cuadre_df['CONTADOR_PWR'].combine_first(analisis_cuadre_df['CONTADOR_PED_YA']).fillna(0).astype(int)

    # Calculate the difference in minutes between FECHA_HORA_PED_YA and FECHA_HORA_PWR
    analisis_cuadre_df['MINS'] = analisis_cuadre_df.apply(
        lambda row: (pd.to_datetime(row['FECHA_HORA_PWR']) - pd.to_datetime(row['FECHA_HORA_PED_YA'])).total_seconds() / 60 if pd.notnull(row['FECHA_HORA_PWR']) and pd.notnull(row['FECHA_HORA_PED_YA']) else 0,
        axis=1
    ).astype(int)

    # Extract information from COMENTARIOS up to the first ";" and add it as a new column
    analisis_cuadre_df['MATCH_PED_YA'] = analisis_cuadre_df['COMENTARIOS'].str.extract(r'^(.*?);', expand=False)

    # Create the VALIDACION_MATCH column comparing MATCH_PED_YA and NUMERO_ORDEN_PED_YA
    analisis_cuadre_df['VALIDACION_MATCH'] = analisis_cuadre_df.apply(
        lambda row: row['MATCH_PED_YA'] == row['NUMERO_ORDEN_PED_YA'], axis=1
    )

    # Create the VALIDACION_CANCELADAS column comparing VALIDACION_MATCH with NUMERO_ID from ordenes_canceladas_df
    canceladas_ids = set(ordenes_canceladas_df['NUMERO_ID'].astype(str))
    analisis_cuadre_df['VALIDACION_CANCELADAS'] = analisis_cuadre_df['NUMERO_ORDEN_PED_YA'].astype(str).apply(
        lambda x: x in canceladas_ids)

    # Reorder columns in analisis_cuadre_df to place TIENDA_UNICA and FECHA_UNICA at the beginning
    analisis_cuadre_df = analisis_cuadre_df[
        ['TIENDA_UNICA', 'FECHA_UNICA', 'CONTADOR_UNICO', 'MINS', 'MATCH_PED_YA', 'VALIDACION_MATCH', 'VALIDACION_CANCELADAS',
         'TIENDA_PWR', 'FECHA_HORA_PWR', 'FECHA_PWR', 'HORA_PWR', 'TIPO_ORDEN_PWR', 'STATUS_ORDEN_PWR', 'NUMERO_ORDEN_PWR',
         'NOMBRE_CLT_PWR', 'TELEFONO_PWR', 'PRECIO_PWR', 'COMENTARIOS', 'CONTADOR_PWR', 'COD_INTERNO_PWR',
         'TIENDA_PED_YA', 'NOMBRE_TIENDA_PED_YA', 'FECHA_HORA_PED_YA', 'FECHA_PED_YA', 'HORA_PED_YA', 'NUMERO_ORDEN_PED_YA',
         'PRECIO_PED_YA', 'CONTADOR_PED_YA', 'COD_INTERNO_PED_YA', 'DIFERENCIA_DESCUENTO']]

    # Sort by TIENDA_UNICA, FECHA_UNICA, and CONTADOR_UNICO to ensure correct order
    analisis_cuadre_df = analisis_cuadre_df.sort_values(by=['TIENDA_UNICA', 'FECHA_UNICA', 'CONTADOR_UNICO']).reset_index(drop=True)

    # ESPACIO DE IAN SORTEANDO Y FILTRANDO ANÁLISIS CUADRE DF
    columnas_posiciones = [0, 1, 13, 11, 12, 14, 15, 16, 29, 26, 23, 21, 20, 25, 17]
    cuadre_limpio_df = analisis_cuadre_df.iloc[:, columnas_posiciones]
    nuevos_nombres = ['Store', 'Order Date', 'Order Number', 'Type', 'Order Status', 'Customer Name', 'Order Phone No.', 'Final Order Price', 'DIFERENCIA DESCUENTO', 'Subtotal', 'FECHA', 'Restaurant name', 'TIENDA', 'Order ID', 'Comentarios']
    cuadre_limpio_df.columns = nuevos_nombres

    # Create cuadre_montado_df with match based on date and monto
    cuadre_montado_df = pd.merge(merged_df, analisis_cuadre_df, left_on=['FECHA_SOFT', 'TIENDA_SOFT'], right_on=['FECHA_UNICA', 'TIENDA_UNICA'], how='outer')

    # Choose location to save the file
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        try:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                merged_df.to_excel(writer, sheet_name='CUADRE', index=False)
                cuadre_montado_df.to_excel(writer, sheet_name='ANALISIS_TIENDA_MONTOS', index=False)
                ordenes_canceladas_df.to_excel(writer, sheet_name='CANCELADAS', index=False)
            messagebox.showinfo("Éxito", f"Archivo guardado exitosamente en {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo: {e}")
    else:
        messagebox.showwarning("Advertencia", "Debe seleccionar una ubicación para guardar el archivo")

# GUI configuration
app = ctk.CTk()
app.geometry("500x500")
app.title("Carga de Archivos y Generación de Reportes")

# Create labels and buttons to load each file
ctk.CTkLabel(app, text="Seleccione el archivo Balanza de Softland:").pack(pady=5)
cargar_balanza_soft_button = ctk.CTkButton(app, text="EXPLORAR", command=lambda: cargar_archivo('balanza_soft'))
cargar_balanza_soft_button.pack(pady=5)

ctk.CTkLabel(app, text="Seleccione el archivo Registro de Sucursales:").pack(pady=5)
cargar_diccionario_button = ctk.CTkButton(app, text="EXPLORAR", command=lambda: cargar_archivo('diccionario'))
cargar_diccionario_button.pack(pady=5)

ctk.CTkLabel(app, text="Seleccione el archivo ordenes de Pedidos Ya:").pack(pady=5)
cargar_ordenes_ped_ya_button = ctk.CTkButton(app, text="EXPLORAR", command=lambda: cargar_archivo('ordenes_ped_ya'))
cargar_ordenes_ped_ya_button.pack(pady=5)

ctk.CTkLabel(app, text="Seleccione el archivo ordenes de PWR:").pack(pady=5)
cargar_ordenes_pwr_button = ctk.CTkButton(app, text="EXPLORAR", command=lambda: cargar_archivo('ordenes_pwr'))
cargar_ordenes_pwr_button.pack(pady=5)

# Button to check the file upload
verificar_button = ctk.CTkButton(app, text="Confirmar Subida de Archivos", command=verificar_archivos)
verificar_button.pack(pady=20)

# Button to process and generate the file
generar_button = ctk.CTkButton(app, text="Generar Archivo", command=generar_archivo)
generar_button.pack(pady=20)

# Run the application
app.mainloop()
