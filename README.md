# OPGW Cable Quantities / Cantidades de Cable OPGW

Calculates and summarizes OPGW cable sets, lengths, and commercial reel requirements for power line spans from Excel data, generating detailed and summary sheets for engineering and procurement in a single output file.

Calcula y resume los conjuntos de cable OPGW, las longitudes y los requerimientos de carretes comerciales para vanos de líneas de transmisión a partir de datos en Excel, generando hojas detalladas y resumidas para ingeniería y compras en un solo archivo de salida.

---

### 📊 Required Excel Column Order / Orden Requerido de Columnas en el Excel

The input Excel file **must contain the following columns in this exact order**:

El archivo Excel de entrada **debe contener las siguientes columnas en este orden exacto**:

| Column Name / Nombre de la Columna | Description / Descripción |
|------------------------------------|----------------------------|
| TORRE                              | Tower identifier / Identificador de la torre |
| TIPO                               | Type of tower / Tipo de torre |
| Ø OPGW ATRAS (mm)                  | OPGW diameter behind the tower (in mm) / Diámetro del OPGW por detrás de la torre (en mm) |
| Ø OPGW ADELANTE (mm)               | OPGW diameter ahead of the tower (in mm) / Diámetro del OPGW por delante de la torre (en mm) |
| CONJ BAJA                          | Downward assembly / Conjunto de bajada |
| CONJ NO BAJA                       | Non-downward assembly / Conjunto sin bajada |
| CONJUNTO                           | Assembly group / Grupo de conjuntos |
| ALTURA ESTRUCTURA (m)              | Structure height (in meters) / Altura de la estructura (en metros) |
| H (m)                              | Distance from floor to the communications splice box (in meters) / Altura desde el suelo hasta la caja de empalme de comunicaciones (en m) |
| VANO (m)                           | Span length (in meters) / Longitud del vano entre estructuras (en metros) |
| CAJA EMPALME                       | Splice box indicator — must say **SÍ** where there is a communications splice box / Indicador de caja de empalme de comunicaciones — debe decir **SÍ** donde haya una |
| LONGITUD DE CABLE (m)              | Cable length as reported by PLS-CADD (in meters) / Longitud del cable reportada por PLS-CADD (en metros) |

> ⚠️ **Important / Importante:**  
> Column names must match exactly, including special characters like accents and parentheses.  
> Los nombres de las columnas deben coincidir exactamente, incluyendo caracteres especiales como tildes y paréntesis.

