import os
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Variables globales para las rutas de los archivos
directorio = None
archivo_adicional_path = None

# Función para cargar el directorio
def cargar_directorio():
    global directorio
    instrucciones = (
        "Instrucciones para cargar la carpeta de sucursales:\n"
        "1. Debe contener los archivos de las sucursales de GreenStar.\n"
        "2. Los nombres de los archivos deben llamarse como están registrados en Pedidos Ya.\n"
        "3. Deben ser archivos CSV."
    )
    messagebox.showinfo("Instrucción", instrucciones)
    directorio = filedialog.askdirectory(title="Selecciona el Directorio Principal")
    if directorio:
        directorio_label.configure(text=f"Directorio cargado: {directorio}")
        messagebox.showinfo("Éxito", "Directorio cargado correctamente")

# Función para cargar el archivo adicional
def cargar_archivo_adicional():
    global archivo_adicional_path
    instrucciones = (
        "Instrucciones para cargar el archivo de Ordenes de Pedidos Ya:\n"
        "1. Debe tener el filtro únicamente de Starbucks.\n"
        "2. Debe ser un archivo CSV."
    )
    messagebox.showinfo("Instrucción", instrucciones)
    archivo_adicional_path = filedialog.askopenfilename(title="Selecciona el Archivo Adicional", filetypes=[("CSV files", "*.csv")])
    if archivo_adicional_path:
        archivo_adicional_label.configure(text=f"Archivo adicional cargado: {archivo_adicional_path}")
        messagebox.showinfo("Éxito", "Archivo adicional cargado correctamente")

## IGCF FUNCIONES PARA LA UNIÓN
def unir_csv(): ## IGCF HACER LA FUNCIÓN PARA UNIR CSVs
    global directorio2
    instrucciones = (
        "Esta función es necesaria en caso de haberse descargado varios archivos de la misma sucursal:\n"
        "1. Debe seleccionar la carpeta con los archivos CSV para unir.\n"
        "2. Se le va a solicitar que guarde el archivo resultante.\n"        
    )
    messagebox.showinfo("Instrucción", instrucciones)
    directorio2 = filedialog.askdirectory(title="Selecciona el directorio con los archivos")
    if directorio2:
        unir_csv_label.configure(text=f"Directorio cargado: {directorio2}")
        messagebox.showinfo("Éxito", "Directorio cargado correctamente\nFavor guardar el archivo resultante.")

    try: # ARREGLAR FUNCIÓN PARA QUE FUSIONE TO DO
        # Listar los archivos en el directorio que terminan con .csv
        archivos = [f for f in os.listdir(directorio2) if f.endswith('.csv')]

        # Crear una lista para almacenar los DataFrames
        dataframes = []

        # Cargar y modificar los archivos de Starbucks
        for archivo in archivos:
            # path_completo = os.path.join(directorio2, archivo)
            # df = pd.read_csv(path_completo)
            # # Eliminar la última fila
            # df = df.iloc[:-1]
            """Concatenates all CSV files in a folder and saves the result to a new file."""
            
            path_completo = os.path.join(directorio2, archivo)
            df = load_and_trim_csv(path_completo)
            dataframes.append(df)
        #
        if dataframes: #IGCF CHEQUEAR AQUÍ
            result_df = pd.concat(dataframes, ignore_index=True)
            save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV (Comma delimited)", "*.csv")])
            print(save_path)

            if save_path:
                    # Save the DataFrames as separate CSV files
                result_df.to_csv(save_path, index=False)
                
                #result_df.to_csv(result_df)
                messagebox.showinfo("Éxito", f"Archivo guardado exitosamente en {save_path}")
            else:
                print("No CSV files found in the folder.")

            # Añadir columna con el nombre del archivo (sin extensión)
            # nombre_local = os.path.splitext(os.path.basename(archivo))[0]
            # print(nombre_local)
            # df['Nombre del local'] = nombre_local
            # dataframes.append(df)


    except Exception as e:
        messagebox.showerror("Error", f"Error durante el procesamiento: {e}")

## IGCF FUNCIONES PARA EL CORTE
def load_and_trim_csv(file_path):
    """Loads a CSV file and trims the last line."""
    df = pd.read_csv(file_path)
    df = df.iloc[:-1]  # Remove the last row
    return df

# Función para verificar si ambos archivos están cargados correctamente
def verificar_carga():
    if directorio and archivo_adicional_path:
        messagebox.showinfo("Éxito", "Ambos archivos se han cargado correctamente. Puede proceder a generar el archivo.")
        generar_button.configure(state="normal")
    else:
        messagebox.showwarning("Advertencia", "Debe cargar tanto el directorio con archivos CSV de las sucursales como el archivo de Ordenes de Pedidos Ya antes de proceder.")

# Función para verificar coincidencias y mostrar advertencias
def verificar_coincidencias(cuadre_st_df, ordenes_ped_ya_df):
    # Obtener las tiendas y fechas únicas de cada DataFrame
    tiendas_cuadre = cuadre_st_df['TIENDA_SIMPHONY'].unique()
    fechas_cuadre = cuadre_st_df['FECHA_SIMPHONY'].unique()
    tiendas_pedidosya = ordenes_ped_ya_df['TIENDA_PED_YA'].unique()
    fechas_pedidosya = pd.to_datetime(ordenes_ped_ya_df['FECHA_PED_YA'], errors='coerce').dt.date.unique()

    # Tiendas y fechas que están en cuadre_st_df pero no en ordenes_ped_ya_df
    tiendas_sin_coincidencia = set(tiendas_cuadre) - set(tiendas_pedidosya)
    fechas_sin_coincidencia = set(fechas_cuadre) - set(fechas_pedidosya)

    mensaje = ""
    if tiendas_sin_coincidencia:
        mensaje += "Las siguientes tiendas en los archivos de sucursales no tienen coincidencia en el archivo de Ordenes de Pedidos Ya:\n" + "\n".join(tiendas_sin_coincidencia) + "\n\n"
    if fechas_sin_coincidencia:
        mensaje += "Las siguientes fechas en los archivos de sucursales no tienen coincidencia en el archivo de Ordenes de Pedidos Ya:\n" + "\n".join(map(str, fechas_sin_coincidencia)) + "\n\n"

    if mensaje:
        messagebox.showwarning("Advertencia", mensaje)

# Función para ejecutar el procesamiento
def ejecutar_procesamiento():
    if not directorio or not archivo_adicional_path:
        messagebox.showwarning("Advertencia", "Debe cargar el directorio y el archivo adicional antes de procesar")
        return

    try:
        # Listar los archivos en el directorio que terminan con .csv
        archivos = [f for f in os.listdir(directorio) if f.endswith('.csv')]

        # Crear una lista para almacenar los DataFrames
        dataframes = []

        # Cargar y modificar los archivos de Starbucks
        for archivo in archivos:
            path_completo = os.path.join(directorio, archivo)
            df = pd.read_csv(path_completo)
            # Eliminar la última fila
            df = df.iloc[:-1]
            # Añadir columna con el nombre del archivo (sin extensión)
            nombre_local = os.path.splitext(os.path.basename(archivo))[0]
            print(nombre_local)
            df['Nombre del local'] = nombre_local
            dataframes.append(df)

        # Concatenar todos los DataFrames en uno solo
        tienda_st_df = pd.concat(dataframes, ignore_index=True)

        # Filtrar filas donde "Tipo de medio de pago" es "PedidosYa"
        tienda_st_df = tienda_st_df[tienda_st_df['Tipo de medio de pago'] == 'PedidosYa']

        # Ajustar el formato de 'Importe de pago' y convertirlo a numérico
        tienda_st_df['Importe de pago'] = tienda_st_df['Importe de pago'].astype(float)

        # Extraer solo la fecha y la hora de "Fecha y hora de cierre de ticket"
        tienda_st_df['Fecha'] = pd.to_datetime(tienda_st_df['Fecha y hora de cierre de ticket'], errors='coerce').dt.date
        #print(tienda_st_df['Fecha y hora de cierre de ticket'])
        tienda_st_df['Hora'] = pd.to_datetime(tienda_st_df['Fecha y hora de cierre de ticket'], errors='coerce').dt.time

        # Seleccionar las columnas necesarias y crear una copia
        cuadre_df = tienda_st_df[['Nombre del local', 'Fecha', 'Importe de pago']].copy()

        # Sumar "Importe de pago" por "Nombre del local" y "Fecha"
        cuadre_df = cuadre_df.groupby(['Nombre del local', 'Fecha'], as_index=False)['Importe de pago'].sum()

        # Renombrar las columnas
        cuadre_df.rename(columns={
            'Nombre del local': 'TIENDA_SIMPHONY',
            'Fecha': 'FECHA_SIMPHONY',
            'Importe de pago': 'MONTO_SIMPHONY'
        }, inplace=True)

        # Cargar el archivo adicional proporcionado
        ordenes_ped_ya_df = pd.read_csv(archivo_adicional_path)

        # Ajustar el formato de 'Total del pedido' y convertirlo a numérico
        ordenes_ped_ya_df['Total del pedido'] = ordenes_ped_ya_df['Total del pedido'].str.replace('.', '').str.replace(',', '.').astype(float)

        # Crear el DataFrame ordenes_canceladas_df con las columnas especificadas
        ordenes_canceladas_df = ordenes_ped_ya_df.loc[ordenes_ped_ya_df['Estado del pedido'] == 'Cancelado', [
            'Total del pedido', 'Nombre del local', 'ID del restaurante', 'ID de pedido', 'Cancelado en', 'Motivo de rechazo', 'Tipo de rechazo'
        ]].copy()

        # Crear el DataFrame con las columnas especificadas
        new_ordenes_ped_ya_df = ordenes_ped_ya_df.loc[ordenes_ped_ya_df['Estado del pedido'] != 'Cancelado', [
            'Nombre del local', 'ID de pedido', 'Fecha del pedido', 'Total del pedido', 'Estado del pedido'
        ]].copy()

        # Ajustar los nombres de las columnas
        new_ordenes_ped_ya_df.rename(columns={
            'Nombre del local': 'TIENDA_PED_YA',
            'ID de pedido': 'NUMERO_PED_YA',
            'Fecha del pedido': 'FECHA_HORA_PED_YA',
            'Total del pedido': 'TOTAL_PED_YA',
            'Estado del pedido': 'ESTADO_PED_YA'
        }, inplace=True)

        # Extraer solo la fecha y la hora de "Fecha del pedido"
        new_ordenes_ped_ya_df['FECHA_PED_YA'] = pd.to_datetime(new_ordenes_ped_ya_df['FECHA_HORA_PED_YA'], errors='coerce').dt.date
        new_ordenes_ped_ya_df['HORA_PED_YA'] = pd.to_datetime(new_ordenes_ped_ya_df['FECHA_HORA_PED_YA'], errors='coerce').dt.time

        # Seleccionar solo las columnas especificadas y en el orden correcto
        analisis_pedidos_df = new_ordenes_ped_ya_df[['TIENDA_PED_YA', 'NUMERO_PED_YA', 'FECHA_HORA_PED_YA', 'TOTAL_PED_YA', 'ESTADO_PED_YA', 'FECHA_PED_YA', 'HORA_PED_YA']].copy()

        # Ordenar y agregar el contador correctamente # IGCF cambio para organizar por fecha y monto en vez de fecha y hora
        analisis_pedidos_df = analisis_pedidos_df.sort_values(by=['TIENDA_PED_YA', 'FECHA_PED_YA', 'TOTAL_PED_YA'], ascending=[True, True, False]).copy()
        analisis_pedidos_df['CONTADOR_UNICO'] = analisis_pedidos_df.groupby(['TIENDA_PED_YA', 'FECHA_PED_YA']).cumcount() + 1
        analisis_pedidos_df['TIENDA_UNICO'] = analisis_pedidos_df['TIENDA_PED_YA']
        analisis_pedidos_df['FECHA_UNICO'] = analisis_pedidos_df['FECHA_PED_YA']
        analisis_pedidos_df['CONTADOR_PED_YA'] = analisis_pedidos_df['CONTADOR_UNICO']
        analisis_pedidos_df['MATCH_PED_YA'] = analisis_pedidos_df.apply(lambda row: f"{row['TIENDA_PED_YA']}-{row['FECHA_PED_YA']}-{row['CONTADOR_PED_YA']}", axis=1)

        # Reiniciar los índices
        analisis_pedidos_df.reset_index(drop=True, inplace=True)

        # Seleccionar las columnas necesarias en tienda_st_df
        new_tienda_st_df = tienda_st_df[['Nombre del local', 'Empleado de transacción', 'Fecha y hora de cierre de ticket', 'Fecha', 'Hora', 'Número de cuenta', 'Importe de pago']].copy()
        new_tienda_st_df.rename(columns={
            'Nombre del local': 'TIENDA_ST',
            'Empleado de transacción': 'EMPLEADO_ST',
            'Fecha y hora de cierre de ticket': 'FECHA_HORA_ST',
            'Fecha': 'FECHA_ST',
            'Hora': 'HORA_ST',
            'Número de cuenta': 'NUMERO_ST',
            'Importe de pago': 'TOTAL_ST'
        }, inplace=True)

        # Ordenar y agregar el contador correctamente #IGCF, modifiqué de hora_st a total_st
        analisis_tienda_df = new_tienda_st_df.sort_values(by=['TIENDA_ST', 'FECHA_ST', 'TOTAL_ST'], ascending=[True, True, False]).copy()
        analisis_tienda_df['CONTADOR_UNICO'] = analisis_tienda_df.groupby(['TIENDA_ST', 'FECHA_ST']).cumcount() + 1
        analisis_tienda_df['TIENDA_UNICO'] = analisis_tienda_df['TIENDA_ST']
        analisis_tienda_df['FECHA_UNICO'] = analisis_tienda_df['FECHA_ST']
        analisis_tienda_df['CONTADOR_ST'] = analisis_tienda_df['CONTADOR_UNICO']
        analisis_tienda_df['MATCH_ST'] = analisis_tienda_df.apply(lambda row: f"{row['TIENDA_ST']}-{row['FECHA_ST']}-{row['CONTADOR_ST']}", axis=1)

        # Reiniciar los índices
        analisis_tienda_df.reset_index(drop=True, inplace=True)

        # Sumar "Total del pedido" por "Nombre del local" y "Fecha"
        resumen_pedidos_df = new_ordenes_ped_ya_df.groupby(['TIENDA_PED_YA', 'FECHA_PED_YA'], as_index=False)['TOTAL_PED_YA'].sum()

        # Renombrar las columnas
        resumen_pedidos_df.rename(columns={
            'TIENDA_PED_YA': 'TIENDA_PEDIDOSYA',
            'FECHA_PED_YA': 'FECHA_PEDIDOSYA',
            'TOTAL_PED_YA': 'MONTO_PEDIDOSYA'
        }, inplace=True)

        # Merge con cuadre_df por "Nombre del local" y "Fecha"
        cuadre_st_df = pd.merge(cuadre_df, resumen_pedidos_df, left_on=['TIENDA_SIMPHONY', 'FECHA_SIMPHONY'], right_on=['TIENDA_PEDIDOSYA', 'FECHA_PEDIDOSYA'], how='left')

        # Agregar la columna DIFERENCIA_DESCUENTO vacía
        cuadre_st_df['DIFERENCIA_DESCUENTO'] = pd.NA

        # Unificar los DataFrames utilizando las columnas únicas
        unificado_df = pd.merge(analisis_tienda_df, analisis_pedidos_df, on=['TIENDA_UNICO', 'FECHA_UNICO', 'CONTADOR_UNICO'], how='outer', suffixes=('_ST', '_PED_YA'))

        # Rellenar valores NA con 0 en las columnas de totales
        unificado_df['TOTAL_ST'] = unificado_df['TOTAL_ST'].fillna(0)
        unificado_df['TOTAL_PED_YA'] = unificado_df['TOTAL_PED_YA'].fillna(0)

        # Calcular la diferencia entre TOTAL_ST y TOTAL_PED_YA
        unificado_df['DIFERENCIA'] = unificado_df['TOTAL_ST'] - unificado_df['TOTAL_PED_YA']

        # Añadir la columna de verificación
        unificado_df['VERIFICADOR'] = unificado_df['DIFERENCIA'].apply(lambda x: 'SOBRANTE' if x < -0.2 else ('FALTANTE' if x > 1 else 'NA'))

        # Asegurarse de que NUMERO_PED_YA y NUMERO_ST sean tratadas como texto
        unificado_df['NUMERO_PED_YA'] = unificado_df['NUMERO_PED_YA'].astype(str).str.replace('.0', '')
        unificado_df['NUMERO_ST'] = unificado_df['NUMERO_ST'].astype(str).str.replace('.0', '')

        # Ordenar por las columnas únicas
        unificado_df = unificado_df.sort_values(by=['TIENDA_UNICO', 'FECHA_UNICO', 'CONTADOR_UNICO'])

        # Reorganizar las columnas en el orden especificado
        columnas_ordenadas = [
            'TIENDA_UNICO', 'FECHA_UNICO', 'CONTADOR_UNICO',
            'CONTADOR_ST', 'MATCH_ST',
            'TIENDA_ST', 'EMPLEADO_ST', 'FECHA_HORA_ST', 'FECHA_ST', 'HORA_ST', 'NUMERO_ST', 'TOTAL_ST',
            'CONTADOR_PED_YA', 'MATCH_PED_YA',
            'TIENDA_PED_YA', 'ESTADO_PED_YA', 'FECHA_HORA_PED_YA', 'FECHA_PED_YA', 'HORA_PED_YA', 'NUMERO_PED_YA', 'TOTAL_PED_YA',
            'DIFERENCIA', 'VERIFICADOR'
        ]

        unificado_df = unificado_df[columnas_ordenadas]

        # Agrupar DIFERENCIA por tienda y fecha
        diferencia_agrupada = unificado_df.groupby(['TIENDA_UNICO', 'FECHA_UNICO'])['DIFERENCIA'].sum().reset_index()
        diferencia_agrupada.rename(columns={'TIENDA_UNICO': 'TIENDA_SIMPHONY', 'FECHA_UNICO': 'FECHA_SIMPHONY', 'DIFERENCIA': 'DIFERENCIA_ANALISIS'}, inplace=True)

        # Merge para añadir DIFERENCIA_ANALISIS a cuadre_st_df
        cuadre_st_df = pd.merge(cuadre_st_df, diferencia_agrupada, on=['TIENDA_SIMPHONY', 'FECHA_SIMPHONY'], how='left')

        # Calcular DIFERENCIA_REAL
        cuadre_st_df['DIFERENCIA_REAL'] = cuadre_st_df['MONTO_SIMPHONY'] - cuadre_st_df['MONTO_PEDIDOSYA'] - cuadre_st_df['DIFERENCIA_ANALISIS']

        # Renombrar la columna DIFERENCIA DESCUENTO a DIFERENCIA_DESCUENTO
        cuadre_st_df.rename(columns={'DIFERENCIA DESCUENTO': 'DIFERENCIA_DESCUENTO'}, inplace=True)

        # Redondear las columnas especificadas a dos decimales
        cuadre_st_df['MONTO_SIMPHONY'] = cuadre_st_df['MONTO_SIMPHONY'].round(2)
        cuadre_st_df['MONTO_PEDIDOSYA'] = cuadre_st_df['MONTO_PEDIDOSYA'].round(2)
        cuadre_st_df['DIFERENCIA_ANALISIS'] = cuadre_st_df['DIFERENCIA_ANALISIS'].round(2)
        cuadre_st_df['DIFERENCIA_REAL'] = cuadre_st_df['DIFERENCIA_REAL'].round(2)

        unificado_df['TOTAL_ST'] = unificado_df['TOTAL_ST'].round(2)
        unificado_df['TOTAL_PED_YA'] = unificado_df['TOTAL_PED_YA'].round(2)
        unificado_df['DIFERENCIA'] = unificado_df['DIFERENCIA'].round(2)

        ordenes_canceladas_df['Total del pedido'] = ordenes_canceladas_df['Total del pedido'].round(2)

        # Verificar coincidencias y mostrar advertencias si es necesario
        verificar_coincidencias(cuadre_st_df, new_ordenes_ped_ya_df)
        # print(unificado_df['TIENDA_ST'].unique())
        # print(unificado_df['TIENDA_UNICO'].unique())
        # print(unificado_df['TIENDA_PED_YA'].unique())

        # Guardar los DataFrames en un único archivo Excel con diferentes hojas
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

        columnas_posiciones = [1,6,10,5,11,21,14,19,17,20,9,18]
        cuadre_limpio_df = unificado_df.iloc[:, columnas_posiciones]
        nuevos_nombres = ['Order Date', 'Empleado', 'Numero', 'Store', 'Monto', 'DIFERENCIA DESCUENTO', 'Store Pedidos Ya', 'Order ID', 'Fecha', 'Subtotal', 'Hora ST', 'Hora Pedidos Ya']
        cuadre_limpio_df.columns = nuevos_nombres


        #cuadre_limpio_df = df.assign(Comentarios='')


        if save_path:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                cuadre_st_df.to_excel(writer, sheet_name='CUADRE', index=False)
                #unificado_df.to_excel(writer, sheet_name='ANALISIS tienda feo', index=False)
                cuadre_limpio_df.to_excel(writer, sheet_name='ANALISIS TIENDA', index=False)
                ordenes_canceladas_df.to_excel(writer, sheet_name='ORDENES CANCELADAS', index=False)
            messagebox.showinfo("Éxito", f"Archivo guardado exitosamente en {save_path}")
        else:
            messagebox.showwarning("Advertencia", "Debe seleccionar una ubicación para guardar el archivo")

    except Exception as e:
        messagebox.showerror("Error", f"Error durante el procesamiento: {e}")

# Configuración de la interfaz gráfica
app = ctk.CTk()
app.geometry("600x400")
app.title("Carga de Archivos y Generación de Reportes")

# Crear descripciones y botones para cargar el directorio y archivo adicional
directorio_label = ctk.CTkLabel(app, text="Seleccionar carpeta de Starbucks (CSV):")
directorio_label.pack(pady=5)
cargar_directorio_button = ctk.CTkButton(app, text="EXPLORAR", command=cargar_directorio)
cargar_directorio_button.pack(pady=5)

archivo_adicional_label = ctk.CTkLabel(app, text="Seleccionar Ordenes de Pedidos Ya (CSV):")
archivo_adicional_label.pack(pady=5)
cargar_archivo_adicional_button = ctk.CTkButton(app, text="EXPLORAR", command=cargar_archivo_adicional)
cargar_archivo_adicional_button.pack(pady=5)

unir_csv_label = ctk.CTkLabel(app, text="En caso de tener varios CSV.")
unir_csv_label.place(relx=1.0, rely=1.0, x=-10, y=-50, anchor='se')
cargar_unir_csv_label = ctk.CTkButton(
    app,
    text="Unir archivos", 
    command=unir_csv,
    fg_color="#74C158",  # Color de fondo
    hover_color="#99FF74",  # Color al pasar el cursor
    text_color="white",  # Color del texto
    font=("Arial", 12),
    width=75,  # Ancho del botón
    height=40  # Alto del botón
    )
#cargar_unir_csv_label.pack(pady=25, padx=2)
# Posicionar el botón en la esquina inferior derecha usando place
cargar_unir_csv_label.place(relx=1.0, rely=1.0, x=-10, y=-10, anchor='se')

# Botón para verificar la carga de archivos
verificar_button = ctk.CTkButton(app, text="Confirmar Carga de Archivos", command=verificar_carga)
verificar_button.pack(pady=20)

# Botón para procesar y generar el archivo, inicialmente deshabilitado
generar_button = ctk.CTkButton(app, text="Generar Archivo", command=ejecutar_procesamiento, state="disabled")
generar_button.pack(pady=20)

# Ejecutar la aplicación
app.mainloop()
