import tkinter as tk
from tkinter import filedialog
import pandas as pd

libro = None
resultado = None

def seleccionar_archivo():
    global libro
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if archivo:
        # Mostrar animación de carga
        etiqueta_animacion.config(text="Cargando archivo...")
        ventana.update()

        # Procesar el archivo seleccionado
        # Aquí puedes realizar cualquier operación adicional con el archivo
        # como leer datos de las hojas, realizar cálculos, etc.
        # Por ahora, solo imprimimos el nombre del archivo seleccionado
        archivo_excel = pd.read_excel(archivo, sheet_name= None)
        archivo_excel.keys()
        final = pd.concat(archivo_excel, ignore_index=True)
        final['FVC'] = pd.to_datetime(final['FVC'])
        final['Número a Portar'] = final['Número a Portar'].astype(str)
        final['Número a Portar'] = final['Número a Portar'].str.replace('.0', '')
        
        #print("Archivo seleccionado:", archivo)
        #print(final['Operador'].value_counts())

        etiqueta_animacion["text"] = "Archivo seleccionado: " + archivo
        ventana.update()

        libro = final


def buscar_datos():
    global libro
    global resultado
    global didas

    texto_informacion.delete("1.0", tk.END)

    marca = entry_marca.get()
    fecha = entry_fecha.get()
    
    print("Marca:", marca)
    print("Fecha:", fecha)
    
    resultado = libro.query(f'Marca == "{marca}" & FVC > "{fecha}" & FVC_Solicitada == "SI"')
    print(resultado[['Nombre', 'Número a Portar', 'FVC', 'Operador']], f"\ntotal de portabilidades de {marca} a partir de {fecha}: ",
    len(resultado.index))
    didas = resultado['Operador'].value_counts()
    print(f'Operadores que se han venido con {marca} a partir de {fecha}.')
    print(didas)

    texto_informacion.insert(tk.END, " ".join(str(resultado[['Nombre', 'Número a Portar', 'FVC', 'Operador']])))
    texto_informacion.insert(tk.END, f"\ntotal de portabilidades de {marca} a partir de {fecha}: ",
    (len(resultado.index)))
    texto_informacion.insert(tk.END, " ".join(str(didas)))

def descargar():
    global resultado
    resultado.to_excel('resultado.xlsx', index=False)
    
#-----------------------------------------------------------------------------------------#

# Crear la ventana principal
ventana = tk.Tk()

ventana.geometry("1400x700")

# Obtener el ancho y alto de la pantalla
ancho_pantalla = ventana.winfo_screenwidth()
alto_pantalla = ventana.winfo_screenheight()

# Calcular las coordenadas x e y para centrar la ventana
pos_x = (ancho_pantalla // 2) - (1400 // 2)
pos_y = (alto_pantalla // 2) - (700 // 2)

# Posicionar la ventana en el centro de la pantalla
ventana.geometry(f"1400x700+{pos_x}+{pos_y}")

#-----------------------------------------------------------------------------------------#
# Crear el botón para seleccionar el archivo
boton_seleccionar = tk.Button(ventana, text="Seleccionar archivo", command=seleccionar_archivo)
boton_seleccionar.pack()

# Crear la etiqueta para mostrar la animación de carga
etiqueta_animacion = tk.Label(ventana, text='Ningun archivo seleccionado ')
etiqueta_animacion.pack()

# Crear etiqueta y entrada para el nombre de la marca
label_marca = tk.Label(ventana, text="Nombre de la marca:")
label_marca.pack()

entry_marca = tk.Entry(ventana)
entry_marca.pack()

# Crear etiqueta y entrada para la fecha
label_fecha = tk.Label(ventana, text="Fecha (2023-12-30):")
label_fecha.pack()

entry_fecha = tk.Entry(ventana)
entry_fecha.pack()

# Crear el botón "Buscar"
boton_buscar = tk.Button(ventana, text="Buscar", command=buscar_datos)
boton_buscar.pack(pady=10)

texto_informacion = tk.Text(ventana, height=30, width=160)
texto_informacion.pack()

boton_descargar = tk.Button(ventana, text="Descargar", command=descargar)
boton_descargar.pack(pady=10)

#-----------------------------------------------------------------------------------------#

# Ejecutar el bucle de eventos de la ventana
ventana.mainloop()
