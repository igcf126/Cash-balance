import os
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

def limpiarResto(): ## IGCF HACER LA FUNCIÓN PARA UNIR CSVs
    global directorio2
    instrucciones = (
        "Limpia todo lo que dice Resto del archivo de pedidos ya"        
    )
    messagebox.showinfo("Instrucción", instrucciones)
    archivoPedYa = filedialog.askopenfilename(title="Selecciona el archivo de Pedidos Ya", filetypes=[("CSV files", "*.csv")])
    # if archivoPedYa:
    #     archivoPedYa.configure(text=f"Archivo de Pedidos Ya cargado: {archivoPedYa}")
    #     messagebox.showinfo("Éxito", "Archivo de Pedidos ya cargado correctamente")

    try: 
    # Specify the string to search and remove
        target_string = " Resto"

        # Perform the operation
        PedYaLimpio = remove_string_from_cells(archivoPedYa, target_string)

        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV (Comma delimited)", "*.csv")])
        print(save_path)

        if save_path:
            # Save the DataFrames as separate CSV files
            PedYaLimpio.to_csv(save_path, index=False)
            
            #result_df.to_csv(result_df)
            messagebox.showinfo("Éxito", f"Archivo guardado exitosamente en {save_path}")
        else:
            print("No CSV files found in the folder.")


    except Exception as e:
        messagebox.showerror("Error", f"Error durante el procesamiento: {e}")

def remove_string_from_cells(csv_file, target_string):
    # Load the CSV file into a DataFrame
    df = pd.read_csv(csv_file)

    # Replace the target string with an empty string in all cells
    output_file = df.applymap(lambda x: str(x).replace(target_string, '') if isinstance(x, str) else x)

    return output_file


# Configuración de la interfaz gráfica
app = ctk.CTk()
app.geometry("600x400")
app.title("Carga de Archivos y Generación de Reportes")


limpiar_csv_label = ctk.CTkLabel(app, text="Limpia resto.")
limpiar_csv_label.place(relx=1.0, rely=1.0, x=-10, y=-50, anchor='se')
cargar_limpiar_csv_label = ctk.CTkButton(
    app,
    text="Unir archivos", 
    command=limpiarResto,
    fg_color="#74C158",  # Color de fondo
    hover_color="#99FF74",  # Color al pasar el cursor
    text_color="white",  # Color del texto
    font=("Arial", 12),
    width=75,  # Ancho del botón
    height=40  # Alto del botón
    )

cargar_limpiar_csv_label.place(relx=1.0, rely=1.0, x=-10, y=-10, anchor='se')

# Ejecutar la aplicación
app.mainloop()
