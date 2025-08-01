import pandas as pd
import numpy as np

# Leer archivo
df = pd.read_excel("conjuntos-python.xlsx")

# Asegurar que los nombres de columnas estén bien escritos
df.columns = df.columns.str.strip()

# Revisar nombres exactos de columnas
col_atras = 'Ø OPGW ATRAS (mm)'
col_adelante = 'Ø OPGW ADELANTE (mm)'
col_conjunto = 'CONJUNTO'

# Diccionario de textos base por figura
conjunto_textos = {
    'Fig.1': 'CONJUNTO DE SUSPENSIÓN NORMAL CON CABALLETE PARA CABLES OPGW',
    'Fig.2': 'CONJUNTO DE SUSPENSIÓN EN BAJADA PARA CABLES OPGW',
    'Fig.7': 'CONJUNTO DE ANCLAJE TERMINAL EN BAJADA PARA CABLE OPGW',
    'Fig.8': 'CONJUNTO DE ANCLAJE NORMAL PASANTE PARA CABLE OPGW',
    'Fig.9': 'CONJUNTO DE ANCLAJE CON TENSOR PASANTE PARA CABLE OPGW',
    'Fig.10': 'CONJUNTO DE BAJADA EN ANCLAJE PARA CABLES OPGW',
}

# Limpiar columnas
df[col_atras] = df[col_atras].astype(str).str.strip()
df[col_adelante] = df[col_adelante].astype(str).str.strip()

# Reemplazar guiones por NaN
df[col_atras] = df[col_atras].replace(['-', '–', ''], pd.NA)
df[col_adelante] = df[col_adelante].replace(['-', '–', ''], pd.NA)

# Rellenar NaN donde ambos son vacíos usando el último Ø ADELANTE válido
last_valid = None
for i, row in df.iterrows():
    if pd.isna(row[col_atras]) and pd.isna(row[col_adelante]):
        if last_valid:
            df.at[i, col_atras] = last_valid
            df.at[i, col_adelante] = last_valid
    else:
        if not pd.isna(row[col_adelante]):
            last_valid = row[col_adelante]

# Crear columna CONJUNTO DETALLADO
detallados = []
for i, row in df.iterrows():
    figura = row[col_conjunto].strip()
    base = conjunto_textos.get(figura, "CONJUNTO DESCONOCIDO")
    atras = row[col_atras]
    adelante = row[col_adelante]

    # Filtrar valores válidos
    diametros_validos = set()
    for val in [atras, adelante]:
        if isinstance(val, str):
            val_clean = val.replace(',', '.').replace(' ', '')
            try:
                float(val_clean)
                diametros_validos.add(val.strip())
            except ValueError:
                continue

    # Ordenar por valor numérico
    diametros_ordenados = sorted(diametros_validos, key=lambda x: float(x.replace(',', '.')))

    # Crear texto
    if diametros_ordenados:
        diametros_texto = ', '.join([f"Ø {d} mm" for d in diametros_ordenados])
        detalle = f"{base} ({diametros_texto}) - {figura}"
    else:
        detalle = f"{base} - {figura}"

    detallados.append(detalle)

# Agregar columna final
df['CONJUNTO DETALLADO'] = detallados

# Crear hoja resumen sin duplicados y con conteo
resumen = df['CONJUNTO DETALLADO'].value_counts().reset_index()
resumen.columns = ['CONJUNTO DETALLADO', 'CANTIDAD']

# Extraer número de figura y ordenarlo
import re
resumen['FIGURA'] = resumen['CONJUNTO DETALLADO'].apply(
    lambda x: int(re.search(r'Fig\.(\d+)', x).group(1)) if re.search(r'Fig\.(\d+)', x) else 999
)

# Ordenar por número de figura
resumen = resumen.sort_values(by='FIGURA').drop(columns='FIGURA')

# Calcular columna CARRETE (m)
df['CARRETE (m)'] = pd.NA

# Encontrar índices donde hay CAJA EMPALME = 'SÍ'
empalme_indices = df.index[df['CAJA EMPALME'].astype(str).str.upper().str.strip() == 'SÍ'].tolist()

for i in range(len(empalme_indices) - 1):
    start_idx = empalme_indices[i]
    end_idx = empalme_indices[i + 1]

    # Cable length for the batch (from start to end)
    cable_length = df.at[start_idx, 'LONGITUD DE CABLE (m)']

    # Altura estructura y H para inicio y fin
    altura_ini = df.at[start_idx, 'ALTURA ESTRUCTURA (m)']
    h_ini = df.at[start_idx, 'H (m)']
    altura_fin = df.at[end_idx, 'ALTURA ESTRUCTURA (m)']
    h_fin = df.at[end_idx, 'H (m)']

    # Calcular longitud de carrete inicial
    carrete = cable_length + (altura_ini - h_ini) + (altura_fin - h_fin) + 100

    # Sumar VANO (m) y LONGITUD DE CABLE (m) para el batch
    batch_vano_sum = df.loc[start_idx:end_idx-1, 'VANO (m)'].sum()
    batch_longitud_sum = df.loc[start_idx:end_idx-1, 'LONGITUD DE CABLE (m)'].sum()

    # Si la suma de VANO (m) es mayor que el carrete calculado, usar la suma de LONGITUD DE CABLE (m)
    if batch_vano_sum > carrete:
        carrete = batch_longitud_sum + (altura_ini - h_ini) + (altura_fin - h_fin) + 100

    # Asignar a todas las filas del batch (incluye end_idx)
    df.loc[start_idx:end_idx, 'CARRETE (m)'] = carrete

    # Solo asignar el valor comercial al primer elemento del batch
    comercial = int(np.ceil(float(carrete) / 500) * 500) if pd.notna(carrete) else pd.NA
    df.at[start_idx, 'COMERCIAL (m)'] = comercial

# Si tienes lógica para el último batch especial (ML/MARCO DE LÍNEA), haz lo mismo:
if empalme_indices:
    last_idx = empalme_indices[-1]
    if (
        str(df.at[last_idx, 'CAJA EMPALME']).strip().upper() == 'SÍ' and (
            str(df.at[last_idx, 'TORRE']).strip().upper() == 'ML' or
            str(df.at[last_idx, 'TIPO']).strip().upper() == 'MARCO DE LÍNEA'
        )
    ):
        # ...cálculo de carrete como ya tienes...
        df.loc[last_idx:, 'CARRETE (m)'] = carrete
        comercial = int(np.ceil(float(carrete) / 500) * 500) if pd.notna(carrete) else pd.NA
        df.at[last_idx, 'COMERCIAL (m)'] = comercial

# Crear hoja de resumen de COMERCIAL (m) sin duplicados y con conteo
comercial_summary = df['COMERCIAL (m)'].value_counts(dropna=True).reset_index()
comercial_summary.columns = ['COMERCIAL (m)', 'CANTIDAD']
comercial_summary = comercial_summary.sort_values(by='COMERCIAL (m)', ascending=True)

# Guardar ambas hojas en el mismo archivo Excel
with pd.ExcelWriter("conjuntos-python-detallado.xlsx", engine='openpyxl', mode='w') as writer:
    df.to_excel(writer, index=False, sheet_name='Detalle')
    resumen.to_excel(writer, index=False, sheet_name='Resumen')
    comercial_summary.to_excel(writer, index=False, sheet_name='Comercial')

print("✅ Archivo actualizado, hoja 'Resumen' ordenada por número de figura.")

# Custom warning message
print("\nEnglish: Review your result, especially in spans with Line Portals and where there is a Communications Splice Box but no cable stringing.")
print("Español: Revise su resultado, especialmente en vanos con Marcos de Línea y donde hay Caja de Empalme de Comunicaciones pero no hay tendido de cable")