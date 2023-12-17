#pip install xlsxwriter
#pip install pandas

from numpy import row_stack
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import tkinter as tk
from tkinter import scrolledtext
from tkinter import filedialog
import datetime

input_cols = [7, 15, 18, 21, 23, 30, 31, 89]
column_names = ["Fecha", "Marca", "Medio","Programa", "Inversion", "Version", "Total Insercion", "Multimedia" ]

# Definir el diccionario de valores por medio
Deflactor = {
    "MEGA": 0.452,
    "CHILEVISION": 0.34,
    "CANAL 13": 0.736,
    "TVN": 0.329,
    "TV+": 0.549,
    "LA RED": 0.665,
    "UCV TV": 0.119,
    "TELEVISION NAC": 0.329,
    "MEGA 2": 0.452
}


def cargar_archivo_campanas():
    global df_origen_cruz_verde
    global df_origen_otras_marcas
    global df_origen
    
    ruta_archivo_campanas = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    label_archivo_origen.config(text=ruta_archivo_campanas)
    
    if ruta_archivo_campanas:
        try:
            # Carga los datos de Cruz Verde desde la hoja correspondiente
            with pd.ExcelFile(ruta_archivo_campanas) as xls:
                df_origen_cruz_verde = pd.read_excel(xls, sheet_name="Cruz Verde")
                df_origen_otras_marcas = pd.read_excel(xls, sheet_name="Otras Campañas")
            
            
            # Combina los DataFrames en uno solo
            df_origen = pd.concat([df_origen_cruz_verde, df_origen_otras_marcas], ignore_index=True)
            print(df_origen_cruz_verde)
            print(df_origen_otras_marcas)
            print("Archivo de Campañas cargado exitosamente.")
        except Exception as e:
            print(f"Error al cargar el archivo de origen: {str(e)}")

def cargar_archivo_origen():
    global df_origen
    ruta_archivo_origen = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    label_archivo_origen.config(text=ruta_archivo_origen)
    
    if ruta_archivo_origen:
        try:
            df_origen = pd.read_excel(ruta_archivo_origen, usecols=input_cols, names=column_names)
            print("Archivo de Origen cargado exitosamente.")
        except Exception as e:
            print(f"Error al cargar el archivo de origen: {str(e)}")


def copiar_datos():
    global df_origen
    global df_origen_cruz_verde
    global df_origen_otras_marcas
    if df_origen is not None:

        scrolled_text.delete("1.0", tk.END)
        
        # Elimina las primeras 16 filas de la hoja de origen
        df_origen = df_origen.iloc[16:]
        df_origen["Deflactor"] = df_origen["Medio"].map(Deflactor)
        df_origen["Neto"] = df_origen["Deflactor"] * df_origen["Inversion"]
        # Formatear la columna "Neto" como moneda chilena
        #df_origen["Neto"] = df_origen["Neto"].apply(lambda x: f"${x:,.0f}")

        # Asignar campañas
        for index, row in df_origen.iterrows():
            if row['Marca'] == "FARMACIAS CRUZ VERDE":
                version = row['Version']
                marca = row['Marca']
                campaña = df_origen_cruz_verde[df_origen_cruz_verde['Version'] == version]['Campaña'].values
                if campaña.any():
                    df_origen.at[index, 'Campaña'] = campaña[0]
                # Verificar si la versión existe en df_origen_cruz_verde
                if version not in df_origen_cruz_verde['Version'].values:
                    new_row = {'Version': version}
                    df_origen_cruz_verde = pd.concat([df_origen_cruz_verde, pd.DataFrame([new_row])],
                                                    ignore_index=True)
                    versiones_sin_campaña = df_origen_cruz_verde[df_origen_cruz_verde['Campaña'].isnull()][
                        'Version']
                    for version in versiones_sin_campaña:
                        resultado = f"Versión sin campaña de la marca {marca}: {version}"
                        print(resultado)
                        scrolled_text.insert(tk.END, resultado + '\n')

            elif row['Marca'] in ["MAICAO", "SALCOBRAND", "AHUMADA", "PREUNIC","FARMACIAS AHUMADA"]:
                version = row['Version']
                marca = row['Marca']
                campaña = df_origen_otras_marcas[df_origen_otras_marcas['Version'] == version]['Campaña'].values
                if campaña.any():
                    df_origen.at[index, 'Campaña'] = campaña[0]
                if version not in df_origen_otras_marcas['Version'].values:
                    #new_row = {'Version': version}
                    new_row = {'Version': version, 'Campaña': None, 'Marca': marca}
                    df_origen_otras_marcas = pd.concat([df_origen_otras_marcas, pd.DataFrame([new_row])],
                                                       ignore_index=True)
                    versiones_sin_campaña = df_origen_otras_marcas[df_origen_otras_marcas['Campaña'].isnull()][
                        'Version']
                    for version in versiones_sin_campaña:
                        resultado = f"Versión sin campaña de la marca {marca}: {version}"
                        print(resultado)
                        scrolled_text.insert(tk.END, resultado + '\n')

        resultado_final = f"\n{df_origen}\n"
        scrolled_text.insert(tk.END, resultado_final)

def asignar_campañas_a_versiones_sin_campaña():
    global df_origen_cruz_verde
    global df_origen_otras_marcas
    versiones_sin_campaña_cruz_verde = df_origen_cruz_verde[df_origen_cruz_verde['Campaña'].isnull()]
    versiones_sin_campaña_otras_marcas = df_origen_otras_marcas[df_origen_otras_marcas['Campaña'].isnull()]
    for index, row in versiones_sin_campaña_cruz_verde.iterrows():
        version = row['Version']
        nueva_campaña = tk.simpledialog.askstring("Nueva Campaña", f"Ingrese el nombre de la campaña para la versión {version}:")
        
        if nueva_campaña is not None:
            df_origen_cruz_verde['Campaña'] = df_origen_cruz_verde['Campaña'].astype(str)
            
            # Actualizar la columna "Campaña" con el nombre de la campaña ingresada
            df_origen_cruz_verde.loc[df_origen_cruz_verde['Version'] == version, 'Campaña'] = nueva_campaña
            print(f"Versión {version} actualizada con la campaña {nueva_campaña}")
    for index, row in versiones_sin_campaña_otras_marcas.iterrows():
        version = row['Version']
        nueva_campaña = tk.simpledialog.askstring("Nueva Campaña", f"Ingrese el nombre de la campaña para la versión {version}:")
        
        if nueva_campaña is not None:
            df_origen_otras_marcas['Campaña'] = df_origen_otras_marcas['Campaña'].astype(str)
            
            # Actualizar la columna "Campaña" con el nombre de la campaña ingresada
            df_origen_otras_marcas.loc[df_origen_otras_marcas['Version'] == version, 'Campaña'] = nueva_campaña
            print(f"Versión {version} actualizada con la campaña {nueva_campaña}")
def guardar_campañas():
    global df_origen_cruz_verde
    global df_origen_otras_marcas

    ruta_archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")])
    if ruta_archivo:
        try:
            with pd.ExcelWriter(ruta_archivo, engine='xlsxwriter') as writer:
                # Guardar df_origen_cruz_verde en la hoja "Cruz Verde"
                df_origen_cruz_verde.to_excel(writer, sheet_name='Cruz Verde', index=False)
                # Guardar df_origen_otras_marcas en la hoja "Otras Campañas"
                df_origen_otras_marcas.to_excel(writer, sheet_name='Otras Campañas', index=False)
            print(f"Archivo actualizado y guardado en: {ruta_archivo}")
        except Exception as e:
            print(f"Error al guardar el archivo: {str(e)}")
def guardar_base():
    global df_origen
    fecha_actual = datetime.datetime.now().strftime("%d-%m-%Y")
    nombre_archivo = f"BASE_{fecha_actual}.xlsx"
    ruta_archivo_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=nombre_archivo)
    if ruta_archivo_salida:
        df_origen.to_excel(ruta_archivo_salida, index=False)
        print(f"Datos exportados exitosamente a '{ruta_archivo_salida}'.")


# Crear una ventana
root = tk.Tk()
root.title("Copiar Datos de Excel")


scrolled_text = scrolledtext.ScrolledText(root, width=100, height=20, wrap=tk.WORD)
scrolled_text.pack(pady=10)

# Variables globales para almacenar los DataFrames
df_origen_cruz_verde = None
df_origen_otras_marcas = None

# Etiquetas
label_archivo_campaña = tk.Label(root, text="Archivo de Campaña:")
label_archivo_campaña.pack()

# Botones para cargar los archivos de campaña
boton_cargar_campaña = tk.Button(root, text="Cargar Archivo de Campaña", command=cargar_archivo_campanas)
boton_cargar_campaña.pack()

label_archivo_campaña = tk.Label(root, text="Archivo de Origen:")
label_archivo_campaña.pack()
# Botones para cargar los archivos de origen
boton_cargar_origen = tk.Button(root, text="Cargar Archivo de Origen", command=cargar_archivo_origen)
boton_cargar_origen.pack()

label_archivo_origen = tk.Label(root, text="")
label_archivo_origen.pack()

# Botones para realizar acciones
boton_copiar = tk.Button(root, text="Copiar Datos", command=copiar_datos)
boton_copiar.pack()

boton_campanas = tk.Button(root, text="Definir Campañas", command=asignar_campañas_a_versiones_sin_campaña)
boton_campanas.pack()

boton_guardar_excel = tk.Button(root, text="Actualizar Campañas", command=guardar_campañas)
boton_guardar_excel.pack()

boton_guardar_base = tk.Button(root, text="Guardar Base", command=guardar_base)
boton_guardar_base.pack()

# Función para cerrar la aplicación
def cerrar_aplicacion():
    root.destroy()

# Botón para cerrar la aplicación
boton_cerrar = tk.Button(root, text="Cerrar", command=cerrar_aplicacion)
boton_cerrar.pack()

root.mainloop()