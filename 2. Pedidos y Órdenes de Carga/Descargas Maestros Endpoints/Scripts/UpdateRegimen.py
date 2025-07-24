import pandas as pd
from openpyxl import load_workbook

# Ruta del archivo
archivo = r'C:\Users\Mateo Zapata\OneDrive\Escritorio\Automatizaciones\2. Pedidos y Órdenes de Carga\Descargas Maestros Endpoints\listado_terceros.xlsx'

# Cargar archivo
df = pd.read_excel(archivo)

# Paso 1: Crear columna 'Nombre completo'
df['Nombre completo'] = df[['NOMBRE', '1ER_APELLIDO', '2DO_APELLIDO']].fillna('').agg(' '.join, axis=1)

# Paso 2: Reemplazar 'Ã‘' por 'Ñ'
df['Nombre completo'] = df['Nombre completo'].str.replace('Ã‘', 'Ñ')

# Paso 3: Crear 'Nombre Completo 2' con tipo NomPropio
df['Nombre Completo 2'] = df['Nombre completo'].str.title()

# Paso 4: Crear columna 'COD_REGIMEN'
mapa_regimen = {
    ' RÃ©gimen Simple de TributaciÃ³n': 7,
    'Gran Contribuyente': 3,
    'No responsable del IVA': 2,
    'Responsable del IVA': 1,
    'Tributario Especial': 5
}
df['COD_REGIMEN'] = df['REGIMEN'].map(mapa_regimen)

# Guardar el resultado en un nuevo archivo
df.to_excel(archivo, index=False)

print('Se actualiza el archivo de Terceros')
