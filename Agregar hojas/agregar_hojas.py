import pandas as pd
import glob
import os

def agregarHoja():
    print("La hoja que quieres agregar tiene que estar en un documento de Excel")
    
    nombreDocumento = input("Ingresa el nombre de tu documento (sin extensión): ")
    nombreHoja = input("Ingresa el nombre del hoja: ")
    directorio = 'C:/Users/jpoot/Documents/Practicas/Agregar_hojas'
    
    patron = '*.xlsx'
    archivos = glob.glob(os.path.join(directorio, patron))
    
    if not archivos:
        print("No se encontraron archivos Excel en el directorio especificado.")
        return

    try:
        nuevahoja = pd.read_excel(f"{nombreDocumento}.xlsx", sheet_name=nombreHoja)
    except Exception as e:
        print(f"Error al leer la hoja {nombreHoja} del documento {nombreDocumento}.xlsx: {e}")
        return
    
    for archivo in archivos:
        try:
            with pd.ExcelWriter(archivo, mode='a', engine='openpyxl') as writer:
                nuevahoja.to_excel(writer, sheet_name=nombreHoja, index=False)
            print(f"Hoja '{nombreHoja}' agregada a {archivo} exitosamente.")
        except Exception as e:
            print(f"Error al agregar la hoja a {archivo}: {e}")

if __name__ == "__main__":
    agregarHoja()
