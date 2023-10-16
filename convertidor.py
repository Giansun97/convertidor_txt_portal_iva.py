import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Formateamos el DataFrame
pd.options.display.max_columns = None
pd.options.display.width = None


def leer_archivos_txt(nombre_archivo_txt: str, column_width: list[int], column_names: list[str]):
    """
    Esta funcion lee los archivos txt, los corta y devuelve un dataframe como resultado

    :param nombre_archivo_txt:
    :param column_width: lista de numeros enteros por los cuales se va a cortar el txt
    :param column_names: lista de nombres que van a tener el nuevo DataFrame:return: con los datos del archivo txt
    """

    df = pd.read_fwf(nombre_archivo_txt,
                     header=None,
                     widths=column_width,
                     names=column_names,
                     encoding='latin1',
                     decimal=',')

    return df


def agregar_ceros(df_name: str, column_name: str, cantidad_ceros: int):
    df_name[column_name] = df_name[column_name].astype(str).str.zfill(cantidad_ceros)


def agrupar_alicuotas(compras_alicuotas):
    """
    Agrupa y resume datos de compras con diferentes alícuotas impositivas.

    Esta función toma un DataFrame de compras con el detalle de las alícuotas impositivas y realiza
    una agrupación basada en el ID del comprobante, proporcionando un resumen de la información
    relevante para cada grupo.

    :param compras_alicuotas:
    :return:
    """

    agg_function = {
        'Tipo de comprobante': 'first',
        'Número de identificación del vendedor': 'first',
        'Importe neto gravado': 'sum',
        'Impuesto liquidado': 'sum'
    }

    compras_alicuotas_agrupado = compras_alicuotas.groupby('ID Cbte').agg(agg_function).reset_index()

    return compras_alicuotas_agrupado


def dividir_y_convertir(df, column_name, divisor, tipo_dato):
    df[column_name] = df[column_name] / divisor
    df[column_name] = df[column_name].astype(tipo_dato)


def cargar_datos_desde_txt_portal_iva(excel_formato_txt, txt_comprobantes_compras, txt_alicuotas_compras):
    """
    Esta funcion lee el archivo que contiene el formato de los txt con el objetivo de extraer las listas necesarias.
    1) Lee los archivos txt.
    2) Agrega ceros a las columnas de Punto de Venta y Número de comprobante.
    3) Crea la columna ID_Cbte, concaténate el punto de venta y el número de comprobante.

    Por último devuelve dos DataFrames compras_cbte y compras_alicuotas

    :param excel_formato_txt:Ruta del archivo de Excel que contiene el formato de las columnas para los
    archivos de texto.
    :param txt_comprobantes_compras: Ruta del archivo de texto que contiene los comprobantes de compras.
    :param txt_alicuotas_compras: Ruta del archivo de texto que contiene las alicuotas de compras.
    :return: DataFrames generados a partir de los txt de comprobantes de compras y alicuotas de compras
    """

    formato_txt_compras = pd.read_excel(
        excel_formato_txt,
        sheet_name='Comprobante_C',
        engine='openpyxl'
    )

    formato_txt_alicuotas = pd.read_excel(
        excel_formato_txt,
        sheet_name='Alicuota_C',
        engine='openpyxl'
    )

    column_names_txt_compras = formato_txt_compras['Descripcion']
    column_names_txt_alicuotas = formato_txt_alicuotas['Descripcion'].tolist()
    column_width_txt_compras = formato_txt_compras['Ancho']
    column_width_txt_alicuotas = formato_txt_alicuotas['Ancho'].tolist()

    compras_cbte = leer_archivos_txt(txt_comprobantes_compras,
                                     column_width_txt_compras,
                                     column_names_txt_compras)

    compras_alicuotas = leer_archivos_txt(txt_alicuotas_compras,
                                          column_width_txt_alicuotas,
                                          column_names_txt_alicuotas)

    # Agregamos ceros a las columnas de punto de venta y número de comprobante
    agregar_ceros(compras_cbte, 'Punto de venta', 4)
    agregar_ceros(compras_cbte, 'Número de comprobante', 10)

    agregar_ceros(compras_alicuotas, 'Punto de venta', 4)
    agregar_ceros(compras_alicuotas, 'Número de comprobante', 10)

    # Creamos la columna ID
    compras_cbte['ID Cbte'] = compras_cbte['Punto de venta'] + '-' + compras_cbte['Número de comprobante']
    compras_alicuotas['ID Cbte'] = compras_alicuotas['Punto de venta'] + '-' + compras_alicuotas[
        'Número de comprobante']

    return compras_cbte, compras_alicuotas


def limpiar_datos(compras_definitivo):
    dividir_y_convertir(compras_definitivo, 'Importe total de la operación', 100, float)
    dividir_y_convertir(compras_definitivo, 'Importe total de conceptos que no integran el precio neto gravado',
                        100, float)
    dividir_y_convertir(compras_definitivo, 'Importe de operaciones exentas',
                        100, float)
    dividir_y_convertir(compras_definitivo, 'Importe de percepciones o pagos a cuenta del Impuesto al Valor Agregado',
                        100, float)
    dividir_y_convertir(compras_definitivo, 'Importe de percepciones o pagos a cuenta de otros impuestos nacionales',
                        100, float)
    dividir_y_convertir(compras_definitivo, 'Importe de percepciones de Ingresos Brutos',
                        100, float)
    dividir_y_convertir(compras_definitivo, 'Importe de Impuestos Internos',
                        100, float)
    dividir_y_convertir(compras_definitivo, 'Importe de percepciones de Impuestos Municipales',
                        100, float)
    dividir_y_convertir(compras_definitivo, 'Tipo de cambio', 1000000, float)

    compras_definitivo['Importe total de la operación'] = (compras_definitivo['Importe total de la operación'] *
                                                           compras_definitivo['Tipo de cambio'])

    dividir_y_convertir(compras_definitivo, 'Importe neto gravado', 100, float)
    dividir_y_convertir(compras_definitivo, 'Impuesto liquidado', 100, float)

    compras_definitivo['Impuesto liquidado'] = (compras_definitivo['Impuesto liquidado'] *
                                                compras_definitivo['Tipo de cambio'])

    compras_definitivo['Importe neto gravado'] = (compras_definitivo['Importe neto gravado'] *
                                                  compras_definitivo['Tipo de cambio'])

    # Ponemos negativos a las notas de credito
    mask = compras_definitivo['Tipo de comprobante'] == 3
    compras_definitivo.loc[mask, 'Impuesto liquidado'] *= -1
    compras_definitivo.loc[mask, 'Importe neto gravado'] *= -1
    compras_definitivo.loc[mask, 'Importe total de la operación'] *= -1
    compras_definitivo.loc[mask, 'Importe total de conceptos que no integran el precio neto gravado'] *= -1
    compras_definitivo.loc[mask, 'Importe de operaciones exentas'] *= -1
    compras_definitivo.loc[mask, 'Importe de percepciones o pagos a cuenta del Impuesto al Valor Agregado'] *= -1
    compras_definitivo.loc[mask, 'Importe de percepciones o pagos a cuenta de otros impuestos nacionales'] *= -1
    compras_definitivo.loc[mask, 'Importe de percepciones de Ingresos Brutos'] *= -1
    compras_definitivo.loc[mask, 'Importe de Impuestos Internos'] *= -1
    compras_definitivo.loc[mask, 'Importe de percepciones de Impuestos Municipales'] *= -1

    # En las Facturas C ponemos el mismo importe en no gravado que el que figura en el total
    mask_facturas_c = compras_definitivo['Tipo de comprobante'] == 11
    compras_definitivo.loc[mask_facturas_c, 'Importe total de conceptos que no integran el precio neto gravado'] = \
        compras_definitivo['Importe total de la operación']


def unir_dataframes(compras_cbte, compras_alicuotas_agrupado):
    cols_to_merge_alic = compras_alicuotas_agrupado[
        ['Importe neto gravado',
         'Impuesto liquidado',
         'ID Cbte']
    ]

    compras_definitivo = pd.merge(
        compras_cbte,
        cols_to_merge_alic,
        how='left',
        on='ID Cbte'
    )

    id_comprobante = compras_definitivo['ID Cbte']
    compras_definitivo = compras_definitivo.drop('ID Cbte', axis=1)
    compras_definitivo = compras_definitivo.assign(ID_Cbte=id_comprobante)
    return compras_definitivo


def seleccionar_archivo(entry):
    archivo = filedialog.askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(0, archivo)


def main():
    """ Esta funcion es la que ejecuta el programa principal."""
    nombre_formato_excel = r'C:/Users/WNS/PycharmProjects/convertidor_txt_portal_iva/data/Formato.xlsx'
    txt_compras_alicuotas = compras_alicuotas_entry.get()
    txt_compras_cbte = compras_cbte_entry.get()

    print("Formato Excel:", nombre_formato_excel)
    print("Compras Alicuotas:", txt_compras_alicuotas)
    print("Compras CBTE:", txt_compras_cbte)

    # Cargamos los datos de los TXT del Portal IVA que nos envia el cliente
    compras_cbte, compras_alicuotas = cargar_datos_desde_txt_portal_iva(
        nombre_formato_excel,
        txt_compras_cbte,
        txt_compras_alicuotas
    )

    compras_alicuotas_agrupado = agrupar_alicuotas(compras_alicuotas)
    compras_definitivo = unir_dataframes(compras_cbte, compras_alicuotas_agrupado)
    limpiar_datos(compras_definitivo)

    # Exportar DataFrame a excel
    directory = os.path.dirname(txt_compras_alicuotas)  # Get the directory of the selected file
    output_file = os.path.join(directory, "txt_convertido.xlsx")
    compras_definitivo.to_excel(output_file, sheet_name="Convertido", index=False)

    messagebox.showinfo("Proceso Completado", "Proceso Completado.\n"
                                              f"Excel guardado en:\n{output_file}")


# --- MAIN APP UI ---
# crear el estilo
app = tk.Tk()
app.title("Convertidor TXT a Excel Portal IVA")

app.geometry("600x300")

# Texto explicativo
explicacion_text = "Seleccione los txt de comprobantes y alicuotas\n"

# Create a label for the explanatory text
tk.Label(app, text=explicacion_text, font=("Arial", 11), justify="left").grid(row=0, column=1, pady=20)

# crear los widgets con el estilo deseado
tk.Label(app, text="Compras CBTE:").grid(row=1, column=0, pady=20)
compras_cbte_entry = tk.Entry(app, width=50)
compras_cbte_entry.grid(row=1, column=1)
tk.Button(app, text="Seleccionar", command=lambda: seleccionar_archivo(compras_cbte_entry)).grid(row=1, column=2)

tk.Label(app, text="Compras Alicuotas:").grid(row=2, column=0, padx=5)
compras_alicuotas_entry = tk.Entry(app, width=50)
compras_alicuotas_entry.grid(row=2, column=1)
tk.Button(app, text="Seleccionar", command=lambda: seleccionar_archivo(compras_alicuotas_entry)).grid(row=2, column=2, padx=5)

tk.Button(app, text="Ejecutar", command=main).grid(row=3, column=1, pady=20)

app.mainloop()