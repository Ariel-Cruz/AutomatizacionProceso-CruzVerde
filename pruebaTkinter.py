import tkinter.messagebox
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import customtkinter
from tkinter import scrolledtext
from numpy import row_stack

input_cols = [7, 15, 18, 21, 23, 30, 31, 89]
column_names = ["Fecha", "Marca", "Medio", "Programa", "Inversion", "Version", "Total Insercion", "Multimedia"]
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

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.df_origen_cruz_verde = None
        self.df_origen_otras_marcas = None
        self.df_origen = None

        # configure window
        self.title("Automatizacion Cruz Verde.py")
        self.geometry(f"{1100}x{580}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Automatizacion Cruz Verde",
                                                 font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        # Botones para cargar los archivos de campaña
        self.boton_cargar_campaña = customtkinter.CTkButton(self.sidebar_frame, text="Cargar Archivo de Campaña",
                                                            command=self.cargar_archivo_campanas)
        self.boton_cargar_campaña.grid(row=1, column=0, padx=20, pady=10)

        # Botones para cargar los archivos de origen
        self.boton_cargar_origen = customtkinter.CTkButton(self.sidebar_frame, text="Cargar Archivo de Origen",
                                                           command=self.cargar_archivo_origen)
        self.boton_cargar_origen.grid(row=2, column=0, padx=20, pady=10)

        self.boton_copiar = customtkinter.CTkButton(self.sidebar_frame, text="Copiar Datos", command=self.copiar_datos)
        self.boton_copiar.grid(row=3, column=0, padx=20, pady=10)

        # Crear el Treeview
        self.tree = ttk.Treeview(self, columns=column_names, show="headings")
        for col in column_names:
            self.tree.heading(col, text=col)
        self.tree.grid(row=1, column=1, pady=10, padx=20, sticky='nsew')

        for col in column_names:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='center')  # Adjust the width as needed

        # Asignar el número de filas como altura inicial


        self.tree.grid(row=0, column=1, pady=10, padx=20, sticky='nsew')

        self.tree_scroll_y = ttk.Scrollbar(self, orient='vertical', command=self.tree.yview)
        self.tree_scroll_y.grid(row=0, column=2, pady=10, sticky='ns')
        self.tree.configure(yscroll=self.tree_scroll_y.set)

        # # Crear el Segundo Treeview
        # #self.tree = ttk.Treeview(self, columns=self.versiones_sin_campaña)
        # for col in column_names:
        #     self.tree.heading(col, text=col)
        # self.tree.grid(row=1, column=1, pady=10, padx=20, sticky='nsew')

        # for col in column_names:
        #     self.tree.heading(col, text=col)
        #     self.tree.column(col, width=100, anchor='center')  # Adjust the width as needed



        # create main entry and button
        self.entry = customtkinter.CTkEntry(self, placeholder_text="CTkEntry")
        self.entry.grid(row=3, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")

        self.main_button_1 = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"))
        self.main_button_1.grid(row=3, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")

    def sidebar_button_event(self):
        print("boton Seleccionado")

    def cargar_archivo_campanas(self):
        global df_origen
        global df_origen_cruz_verde
        global df_origen_otras_marcas
        ruta_archivo_campanas = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
        self.boton_cargar_campaña.configure(text=ruta_archivo_campanas)

        if ruta_archivo_campanas:
            try:
                with pd.ExcelFile(ruta_archivo_campanas) as xls:
                    self.df_origen_cruz_verde = pd.read_excel(xls, sheet_name="Cruz Verde")
                    self.df_origen_otras_marcas = pd.read_excel(xls, sheet_name="Otras Campañas")

                self.df_origen_cruz_verde['Campaña'] = self.df_origen_cruz_verde['Campaña'].astype(str)
                self.df_origen_otras_marcas['Campaña'] = self.df_origen_otras_marcas['Campaña'].astype(str)
                self.df_origen = pd.concat([self.df_origen_cruz_verde, self.df_origen_otras_marcas], ignore_index=True)

                print(self.df_origen_cruz_verde)
                print(self.df_origen_otras_marcas)
                print("Archivo de Campañas cargado exitosamente.")
            except Exception as e:
                print(f"Error al cargar el archivo de origen: {str(e)}")

    def cargar_archivo_origen(self):
        global df_origen
        ruta_archivo_origen = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
        self.boton_cargar_origen.configure(text=ruta_archivo_origen)

        if ruta_archivo_origen:
            try:
                self.df_origen = pd.read_excel(ruta_archivo_origen, usecols=input_cols, names=column_names)
                print("Archivo de Origen cargado exitosamente.")
            except Exception as e:
                print(f"Error al cargar el archivo de origen: {str(e)}")

    def copiar_datos(self):
        if self.df_origen is not None:

            # Elimina las primeras 16 filas de la hoja de origen
            self.df_origen = self.df_origen.iloc[16:]
            self.df_origen["Deflactor"] = self.df_origen["Medio"].map(Deflactor)
            self.df_origen["Neto"] = self.df_origen["Deflactor"] * self.df_origen["Inversion"]

            # Asignar campañas
            for index, row in self.df_origen.iterrows():
                if row['Marca'] == "FARMACIAS CRUZ VERDE":
                    version = row['Version']
                    marca = row['Marca']
                    campaña = self.df_origen_cruz_verde[self.df_origen_cruz_verde['Version'] == version]['Campaña'].values
                    if campaña.any():
                        self.df_origen.at[index, 'Campaña'] = campaña[0]
                    # Verificar si la versión existe en self.df_origen_cruz_verde
                    if version not in self.df_origen_cruz_verde['Version'].values:
                        new_row = {'Version': version}
                        self.df_origen_cruz_verde = pd.concat([self.df_origen_cruz_verde, pd.DataFrame([new_row])],
                                                            ignore_index=True)
                        versiones_sin_campaña = self.df_origen_cruz_verde[
                            self.df_origen_cruz_verde['Campaña'].isnull()][
                            'Version']
                        for version in versiones_sin_campaña:
                            resultado = f"Versión sin campaña de la marca {marca}: {version}"
                            print(resultado)

                elif row['Marca'] in ["MAICAO", "SALCOBRAND", "AHUMADA", "PREUNIC", "FARMACIAS AHUMADA"]:
                    version = row['Version']
                    marca = row['Marca']
                    campaña = self.df_origen_otras_marcas[self.df_origen_otras_marcas['Version'] == version]['Campaña'].values
                    if campaña.any():
                        self.df_origen.at[index, 'Campaña'] = campaña[0]
                    if version not in self.df_origen_otras_marcas['Version'].values:
                        new_row = {'Version': version, 'Campaña': None, 'Marca': marca}
                        self.df_origen_otras_marcas = pd.concat([self.df_origen_otras_marcas, pd.DataFrame([new_row])],
                                                                ignore_index=True)
                        self.versiones_sin_campaña = self.df_origen_otras_marcas[
                            self.df_origen_otras_marcas['Campaña'].isnull()][
                            'Version']
                        for version in self.versiones_sin_campaña:
                            resultado = f"Versión sin campaña de la marca {marca}: {version}"
                            print(resultado)

            # Limpiar el Treeview antes de insertar datos
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Asignar el número de filas actualizado como altura
            self.tree.config(height=len(self.df_origen) + 1)

            # Insertar todas las filas al final del bucle
            resultado_final = self.df_origen.to_records(index=False)
            for row in resultado_final:
                self.tree.insert("", "end", values=list(row))

            # Ajustar la altura del ScrolledText
            self.scrolled_text.config(height=len(resultado_final) + 2)  # +2 para incluir dos líneas adicionales


  
    def asignar_campañas_a_versiones_sin_campaña(self):
        self.versiones_sin_campaña_cruz_verde = self.df_origen_cruz_verde[self.df_origen_cruz_verde['Campaña'].isnull()]
        self.versiones_sin_campaña_otras_marcas = self.df_origen_otras_marcas[self.df_origen_otras_marcas['Campaña'].isnull()]

        for index, row in self.versiones_sin_campaña_cruz_verde.iterrows():
            version = row['Version']
            nueva_campaña = ttk.Treeview_simpledialog.askstring("Nueva Campaña", f"Ingrese el nombre de la campaña para la versión {version}:")

            if nueva_campaña is not None:
                self.df_origen_cruz_verde['Campaña'] = self.df_origen_cruz_verde['Campaña'].astype(str)

                # Actualizar la columna "Campaña" con el nombre de la campaña ingresada
                self.df_origen_cruz_verde.loc[self.df_origen_cruz_verde['Version'] == version, 'Campaña'] = nueva_campaña
                print(f"Versión {version} actualizada con la campaña {nueva_campaña}")

        for index, row in self.versiones_sin_campaña_otras_marcas.iterrows():
            version = row['Version']
            nueva_campaña = ttk.Treeview_simpledialog.askstring("Nueva Campaña", f"Ingrese el nombre de la campaña para la versión {version}:")

            if nueva_campaña is not None:
                self.df_origen_otras_marcas['Campaña'] = self.df_origen_otras_marcas['Campaña'].astype(str)

                # Actualizar la columna "Campaña" con el nombre de la campaña ingresada
                self.df_origen_otras_marcas.loc[self.df_origen_otras_marcas['Version'] == version, 'Campaña'] = nueva_campaña
                print(f"Versión {version} actualizada con la campaña {nueva_campaña}")




if __name__ == "__main__":
    app = App()
    app.mainloop()
