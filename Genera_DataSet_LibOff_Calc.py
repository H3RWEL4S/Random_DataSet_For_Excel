#--!Es necesario instalar librerías complementarias para que el Script funcione
#--pip install libreoffice-connect
import libreoffice
import random
from libreoffice.connect import LibreOfficeConnect

# Establecer la ruta al ejecutable de LibreOffice
libreoffice_path = "/path/to/libreoffice"  # Ajusta la ruta según tu sistema

# Crear una instancia de LibreOfficeConnect
conn = LibreOfficeConnect(libreoffice_path)

# Abrir un nuevo documento de Calc
doc = conn.open_document("NuevoDocumento.ods")  # Ajusta el nombre del archivo

# Obtener la hoja activa
sheet = doc.get_sheet("Hoja1")  # Ajusta el nombre de la hoja si es necesario

# Solicitar al usuario los encabezados de las columnas
columnas = []
for i in range(1, 6):
  encabezado = input(f"Ingrese el encabezado de la columna {i}: ")
  columnas.append(encabezado)

# Escribir los encabezados en la fila 1
for i, encabezado in enumerate(columnas):
  sheet.get_cell(row=1, column=i+1).value = encabezado

# Generar 100 filas de datos aleatorios
for fila in range(2, 102):
  for i, columna in enumerate(columnas):
    # Generar datos aleatorios según el tipo de columna
    if columna == "Número":
      dato = random.randint(1, 1000)
    elif columna == "Fecha":
      dato = random.date(2020, 1, 1, 2024, 12, 31)
    elif columna == "Cadena":
      dato = chr(random.randint(65, 90)) + chr(random.randint(97, 122)) + str(random.randint(100, 999))
    else:
      dato = "Dato no válido"

    sheet.get_cell(row=fila, column=i+1).value = dato

# Guardar el archivo y cerrar el documento
doc.save("DatosAleatorios.ods")  # Ajusta el nombre del archivo si es necesario
doc.close()

print(f"Datos aleatorios generados en el archivo: DatosAleatorios.ods")
