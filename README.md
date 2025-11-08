# Herramienta para crear órdenes – Seller Addi (Streamlit)

Aplicación web desarrollada en Streamlit para procesar archivos Excel, mapear columnas entre un archivo origen y un template, y generar archivos Excel listos para importar en lotes configurables.

---

## 📋 Descripción General

Esta aplicación permite:

- **Cargar archivo Excel origen** con datos de órdenes
- **Cargar archivo Excel template** con formato requerido
- **Mapear columnas** entre origen y destino de forma interactiva
- **Aplicar transformaciones automáticas**:
  - Limpieza y validación de datos (teléfonos, correos)
  - Asignación automática de bodega según ciudad/departamento
  - Consolidación de registros por criterios configurables
- **Generar archivos en lotes** (configurable, default 100 registros por archivo)
- **Descargar resultado** como archivo ZIP

---

## 🏗️ Arquitectura y Funcionalidades

### 1. Sistema de Autenticación

**Ubicación:** Líneas 15-47

**Propósito:** Control de acceso mediante contraseña.

**Funcionamiento:**
- Prioriza contraseña desde `st.secrets["APP_PASSWORD"]`
- Si no existe, busca en variable de entorno `APP_PASSWORD`
- Si no existe, usa valor hardcoded `REQUIRED_PASSWORD`
- Guarda estado de autenticación en `st.session_state.auth_ok`
- Muestra formulario de contraseña si no está autenticado
- Recarga la aplicación al autenticarse correctamente

**Componentes clave:**
- `APP_TITLE`: Título de la aplicación
- `APP_SUBTITLE`: Descripción breve
- `REQUIRED_PASSWORD`: Contraseña por defecto

---

### 2. Sistema de Progreso

**Ubicación:** Líneas 49-66

**Propósito:** Mostrar progreso visual del procesamiento.

**Clase `ProgressTracker`:**
- Inicializa con total de filas a procesar
- Actualiza porcentaje y contador `(actual/total)`
- Maneja casos donde el progreso excede el total (capping)
- Permite personalizar etiquetas de texto

**Métodos:**
- `__init__(total_rows, label)`: Inicializa barra de progreso
- `add(n, label)`: Incrementa progreso en `n` unidades
- `finish(label)`: Marca como completado (100%)

---

### 3. Configuración de Bodegas

**Ubicación:** Líneas 68-103

**Propósito:** Definir bodegas disponibles y funciones de normalización.

#### Lista de Bodegas (`WAREHOUSES`)

Estructura:
```python
WAREHOUSES = [
    {"label": "Nombre completo de la bodega", "city": "Ciudad normalizada"},
]
```

- `label`: Nombre que se escribirá en el Excel generado
- `city`: Ciudad normalizada (sin acentos, minúscula) que se usa para mapeo

#### Funciones de Normalización

**`_norm(s: str) -> str`**
- Convierte a minúsculas
- Elimina espacios al inicio/fin
- Normaliza Unicode (NFKD) y elimina diacríticos (acentos)
- Ejemplo: `"Medellín"` → `"medellin"`

**`_norm_hard(s: str) -> str`**
- Aplica `_norm()` y colapsa espacios múltiples a uno solo
- Ejemplo: `"San   Gil"` → `"san gil"`

**`_slugify_no_spaces(s: str) -> str`**
- Normaliza y elimina todo excepto a-z0-9
- Elimina espacios completamente
- Ejemplo: `"Brand Name"` → `"brandname"`

**`make_external_order_slug(brand: str, empresa: str) -> str`**
- Crea slug combinado: `brand-empresa`
- Usa `_slugify_no_spaces()` en ambos parámetros
- Ejemplo: `"Brand A"` + `"Empresa B"` → `"branda-empresab"`

**`_get_wh_label_for_city(hub_city_norm: str) -> str`**
- Busca bodega por ciudad normalizada en `WAREHOUSES`
- Retorna el `label` de la bodega correspondiente
- Si no encuentra, retorna la primera bodega (fallback)

---

### 4. Mapeo de Ciudades y Departamentos a Hubs

**Ubicación:** Líneas 105-247

**Propósito:** Asignar automáticamente bodega según ciudad o departamento de destino.

#### CITY_TO_HUB (Líneas 105-203)

Diccionario que mapea ciudades normalizadas a hubs de distribución.

**Estructura:**
```python
CITY_TO_HUB = {
    "ciudad_normalizada": "hub_normalizado",
    "medellin": "medellin",
    "medellín": "medellin",  # Variante con acento
    # ... más ciudades
}
```

**Características:**
- Incluye variantes con y sin acentos
- Los valores deben coincidir con ciudades en `WAREHOUSES` (normalizadas)
- Tiene prioridad sobre `DEPT_TO_HUB`

#### DEPT_TO_HUB (Líneas 205-232)

Diccionario que mapea departamentos normalizados a hubs (fallback).

**Estructura:**
```python
DEPT_TO_HUB = {
    "departamento_normalizado": "hub_normalizado",
    "antioquia": "medellin",
    "cundinamarca": "bogota",
    # ... más departamentos
}
```

**Características:**
- Se usa solo si la ciudad no está en `CITY_TO_HUB`
- Incluye variantes con y sin acentos

#### KEYWORDS (Líneas 235-236)

Listas de palabras clave para asignación cuando ni ciudad ni departamento están mapeados.

```python
KEYWORDS_MEDELLIN = ["medellin", "sabaneta", ...]
KEYWORDS_BOGOTA = ["bogota", "cundinamarca", ...]
```

#### Función `assign_bodega_by_city(row: pd.Series) -> str`

**Algoritmo de asignación (en orden de prioridad):**
1. Normaliza ciudad y busca en `CITY_TO_HUB`
2. Si no encuentra, normaliza departamento y busca en `DEPT_TO_HUB`
3. Si no encuentra, busca keywords en ciudad o departamento
4. Si nada funciona, retorna bodega por defecto (primera en `WAREHOUSES`)

**Retorna:** Label de la bodega (ej: `"Bogotá #2 - Montevideo"`)

---

### 5. Interfaz de Usuario - Sidebar

**Ubicación:** Líneas 249-260

**Propósito:** Configuración de parámetros del procesamiento.

**Parámetros configurables:**
- **`chunk_size`**: Tamaño máximo de registros por archivo (default: 100)
- **`header_row`**: Fila donde están los encabezados del template (default: 1)
- **`start_row`**: Fila inicial donde escribir datos (default: 3, es decir A3)
- **`default_prefix`**: Prefijo para nombres de archivos generados (default: "template_part")

---

### 6. Carga de Archivo Origen

**Ubicación:** Líneas 264-329

**Propósito:** Cargar y procesar archivo Excel con datos origen.

**Flujo:**
1. Usuario carga archivo `.xlsx` mediante `st.file_uploader`
2. Aplicación lee nombres de hojas con `pd.ExcelFile`
3. Usuario selecciona hoja a procesar
4. Aplicación parsea hoja con `pd.read_excel(dtype=object)` para preservar tipos
5. Limpia nombres de columnas (elimina espacios)
6. Muestra resumen: número de filas y columnas
7. Aplica transformaciones automáticas (ver sección 7)

**Características:**
- Soporta archivos `.xlsx`
- Preserva tipos de datos originales
- Muestra vista previa de primeras 20 filas

---

### 7. Limpiezas y Formateo Automático

**Ubicación:** Líneas 280-321

**Propósito:** Aplicar transformaciones automáticas a los datos del origen.

#### 7.1 Autocompletado de Teléfonos (Líneas 290-298)

**Funcionalidad:**
- Detecta teléfonos vacíos en columna "Celular"
- Genera números aleatorios válidos para Colombia:
  - 10 dígitos
  - Inician en "3"
  - Formato: `3XXXXXXXXX`

**Implementación:**
```python
def _random_phone():
    return "3" + "".join(str(random.randint(0, 9)) for _ in range(9))
```

#### 7.2 Limpieza de Correos Electrónicos (Líneas 300-308)

**Funcionalidad:**
- Convierte todos los correos a minúscula
- Solo mantiene correos de dominios permitidos:
  - `@gmail.com`
  - `@hotmail.com`
- Todos los demás correos se convierten a cadena vacía

**Implementación:**
```python
mask_valid = (
    src_df["Correo electrónico"].str.endswith("@gmail.com")
    | src_df["Correo electrónico"].str.endswith("@hotmail.com")
)
src_df.loc[~mask_valid, "Correo electrónico"] = ""
```

#### 7.3 Generación de "Número de orden externo" (Líneas 310-315)

**Funcionalidad:**
- Genera campo "Número de orden externo" combinando Brand Slug y Nombre de la empresa
- Formato: `brand-empresa` (slug sin espacios, sin acentos, minúscula)
- Usa función `make_external_order_slug()`

**Ejemplo:** `"Brand A"` + `"Empresa B"` → `"branda-empresab"`

---

### 8. Carga de Archivo Template

**Ubicación:** Líneas 331-365

**Propósito:** Cargar template Excel y extraer estructura de columnas.

**Flujo:**
1. Usuario carga archivo `.xlsx` template
2. Aplicación lee bytes del archivo
3. Abre con `openpyxl.load_workbook(data_only=True)` para valores calculados
4. Usuario selecciona hoja a usar
5. Aplicación lee encabezados de la fila `header_row`
6. Crea estructuras de mapeo:
   - `headers`: Lista de nombres de columnas
   - `header_index`: `{nombre_columna: índice_columna}` (1-indexed)
   - `header_positions`: `{nombre_columna: [índices...]}` para columnas duplicadas

**Características:**
- Maneja columnas con nombres duplicados (ej: múltiples "Indicativo")
- Usa `openpyxl` para preservar formato del Excel
- Muestra lista de encabezados encontrados

---

### 9. Mapeo de Columnas

**Ubicación:** Líneas 367-464

**Propósito:** Definir relación entre columnas del template y origen/constantes.

#### 9.1 Mapeo Predefinido (`preset_mapping`)

**Ubicación:** Líneas 372-390

Define mapeos por defecto entre columnas del template y origen/constantes.

**Modos disponibles:**
- `"source"`: Toma valor de columna del origen
- `"const"`: Valor constante fijo
- `"template_name"`: Nombre del archivo template
- `"source_filename"`: Nombre del archivo origen
- `"(no escribir)"`: No escribir nada en esa columna

**Ejemplo:**
```python
preset_mapping = {
    "Nombre completo del comprador": {
        "mode": "source",
        "source_col": "Nombre completo"
    },
    "País": {
        "mode": "const",
        "const_value": "Colombia"
    },
    # ...
}
```

#### 9.2 Interfaz de Mapeo (`draw_mapping_ui`)

**Ubicación:** Líneas 395-461

**Funcionalidad:**
- Crea UI interactiva para mapear cada columna del template
- Permite seleccionar modo y origen/valor para cada columna
- Guarda estado en `st.session_state.mapping_state` para persistencia
- Bloquea edición de columnas "Bodega" y "CEDIS de origen" (se calculan automáticamente)

**Características:**
- Muestra todas las columnas del template
- Dropdown para seleccionar columna origen (si modo es "source")
- Input de texto para valor constante (si modo es "const")
- Mantiene selecciones anteriores al recargar

---

### 10. Funciones Auxiliares de Procesamiento

**Ubicación:** Líneas 466-560

#### 10.1 `resolve_value(spec, row, template_name, source_name)`

**Propósito:** Resuelve el valor a escribir en una celda según la especificación del mapeo.

**Parámetros:**
- `spec`: Diccionario con `mode` y datos adicionales
- `row`: Serie de pandas con datos de la fila origen
- `template_name`: Nombre del template
- `source_name`: Nombre del archivo origen

**Lógica:**
- Si `mode == "source"`: Extrae valor de `row[source_col]`
- Si `mode == "const"`: Retorna valor constante (intenta convertir a número si es posible)
- Si `mode == "template_name"`: Retorna `template_name`
- Si `mode == "source_filename"`: Retorna `source_name`
- Si `mode == "(no escribir)"`: Retorna `None`

#### 10.2 `fill_one_chunk(...)`

**Propósito:** Llena un chunk (lote) de datos en el template Excel.

**Parámetros:**
- `tmpl_bytes`: Bytes del template
- `target_sheet`: Nombre de la hoja destino
- `header_index`: Diccionario de índices de columnas
- `header_positions`: Diccionario de posiciones múltiples
- `start_row`: Fila inicial para escribir
- `chunk_df`: DataFrame con datos a escribir
- `mapping`: Diccionario de mapeo de columnas
- `template_name`: Nombre del template
- `source_name`: Nombre del origen
- `prog`: Instancia de ProgressTracker

**Proceso:**
1. Carga el workbook del template
2. Obtiene la hoja destino
3. Detecta columna destino para bodega ("Bodega" o "CEDIS de origen")
4. Para cada fila del chunk:
   - Escribe valores según mapeo normal
   - Asigna bodega automáticamente con `assign_bodega_by_city()`
   - Maneja columna "Indicativo": solo columna C (índice 3) con valor 57, otras vacías
5. Guarda el workbook en BytesIO
6. Retorna bytes y estadísticas

**Regla especial - Indicativo:**
- Si hay múltiples columnas "Indicativo", solo se llena la columna C (índice 3)
- Las demás columnas "Indicativo" se dejan vacías
- Si no existe columna C, se usa la primera encontrada

**Estadísticas retornadas:**
- `rows`: Número de filas procesadas
- `nw_written`: Filas donde se escribió bodega
- `no_dest_bodega`: Filas donde no se pudo escribir bodega

---

### 11. Consolidación de Datos

**Ubicación:** Líneas 562-628

**Propósito:** Consolidar registros agrupando por criterios y aplicando tope máximo.

#### Función `consolidate_by_brand_company(df: pd.DataFrame) -> pd.DataFrame`

**Algoritmo:**
1. Valida que existan columnas necesarias: "Brand Slug", "Nombre de la empresa", "Número de tiendas"
2. Genera "Número de orden externo" para todo el DataFrame
3. Normaliza Brand Slug y Nombre de la empresa para agrupar
4. Agrupa por (`__b__`, `__e__`) normalizados
5. Para cada grupo:
   - Suma "Número de tiendas" (convierte a numérico)
   - Aplica tope de 4 unidades (CAP)
   - Toma primera fila del grupo como representante
   - Actualiza "Número de tiendas" con valor con tope
   - Recalcula "Número de orden externo"
6. Retorna DataFrame consolidado

**Características:**
- Agrupa por combinación única de (Brand Slug, Nombre de la empresa)
- Suma unidades y aplica tope máximo de 4 por grupo
- Mantiene primera fila de cada grupo como representante
- Genera métricas: grupos creados, filas eliminadas, cantidad total después del tope

**Métricas guardadas:**
- `brand_company_groups`: Número de grupos creados
- `brand_company_removed`: Filas eliminadas por consolidación
- `cap_per_group`: Tope aplicado (4)
- `total_qty_after_cap`: Cantidad total después del tope

---

### 12. Empaquetado Inteligente por Registros y Unidades (Opcional)

**Ubicación:** Nueva funcionalidad a implementar después de consolidación y antes de división en archivos

**⚠️ IMPORTANTE:** Esta funcionalidad optimiza cómo se distribuyen los registros en archivos de 100 registros, agrupando por combinación (producto+tienda). Tiene DOS limitantes:
1. **REGISTROS:** Máximo 100 registros por archivo
2. **UNIDADES TOTALES:** Máximo X unidades totales por combinación en el archivo (opcional, si se especifica)

Solo aplica si usas consolidación por producto+tienda.

**Propósito:** Optimizar la distribución de registros en archivos agrupando por combinación y respetando límites de registros y unidades.

**Qué hace:**
- Agrupa registros por combinación (tienda + producto)
- Controla que cada archivo tenga máximo 100 registros
- **OPCIONAL:** Controla que las unidades totales por combinación en el archivo no superen un límite especificado
- Optimiza el empaquetado: si un archivo tiene 80 registros de una combinación, busca otras combinaciones que quepan en los 20 registros restantes
- Maneja combinaciones grandes (>100 registros) según la estrategia configurada

**Reglas de empaquetado (con límite de registros):**
1. **Si combinación < 100 registros:** Dejarlos todos en un archivo (si también cumple límite de unidades si está configurado)
2. **Si combinación = 100 registros:** Dejarlos solo en un archivo (si también cumple límite de unidades si está configurado)
3. **Si archivo termina con X registros (< 100):** Buscar otra combinación que:
   - Tenga máximo (100 - X) registros para que quepa
   - Si hay límite de unidades: que las unidades totales de esa combinación no excedan el límite
   - Si no hay, dejarlo con X registros y continuar
4. **Si combinación > 100 registros (ej: 150):** Según estrategia configurada:
   - **Estrategia A (Dividir):** Archivo 1 con 100 registros de esa combinación, Archivo 2 con 50 registros restantes + otras combinaciones
   - **Estrategia B (Archivos completos):** Crear archivos completos solo con esa combinación (sin dividir)
   - **Estrategia C (Un archivo por combinación):** Un archivo por combinación completa, sin importar cuántos registros tenga

**Reglas adicionales (con límite de unidades):**
- Si se especifica límite de unidades por combinación, cada combinación en un archivo no puede superar ese límite
- Ejemplo: Si límite es 100 unidades y Combinación A tiene 80 unidades, puede agregarse al archivo
- Si Combinación B tiene 50 unidades y el archivo ya tiene Combinación A (80 unidades), puede agregarse siempre y cuando no supere la cantidad de registros
- El límite de unidades se aplica POR COMBINACIÓN en el archivo, no al total del archivo
- **IMPORTANTE:** El límite de unidades se aplica por cada combinación individualmente. Diferentes combinaciones pueden coexistir en el mismo archivo siempre que cada una cumpla su límite de unidades.

**Estrategias disponibles:**
- **Estrategia A (Dividir):** Combinaciones grandes se dividen, permitiendo mezclar combinaciones en el mismo archivo
- **Estrategia B (Archivos completos):** Cada archivo contiene solo una combinación (puede dividirse si > 100 registros)
- **Estrategia C (Un archivo por combinación):** Cada combinación va en su propio archivo completo, sin importar tamaño

**Ejemplo detallado con límite de unidades:**

**Escenario después de consolidación:**
- Combinación A (Tienda X + Producto Y): 80 registros, 80 unidades totales
- Combinación B (Tienda Z + Producto W): 20 registros, 50 unidades totales
- Combinación C (Tienda Y + Producto Z): 150 registros, 150 unidades totales
- Combinación D (Tienda X + Producto W): 15 registros, 15 unidades totales

**Límites configurados:**
- Máximo 100 registros por archivo (obligatorio)
- Máximo 100 unidades totales por combinación en el archivo (opcional, si se especifica)

**Resultado esperado con límite de unidades:**
- Archivo 1: Combinación A (80 registros, 80 unidades) + Combinación D (15 registros, 15 unidades) = 95 registros
  - Combinación A: 80 unidades ≤ 100 ✓
  - Combinación D: 15 unidades ≤ 100 ✓
  - No se puede agregar Combinación B porque excedería el límite de registros (95 + 20 = 115 > 100)
- Archivo 2: Combinación B (20 registros, 50 unidades) - archivo completo
  - Combinación B: 50 unidades ≤ 100 ✓
- Archivo 3: Combinación C (100 registros, 100 unidades) - primera parte
  - Combinación C parte 1: 100 unidades ≤ 100 ✓ (justo en el límite)
- Archivo 4: Combinación C (50 registros restantes, 50 unidades) - segunda parte
  - Combinación C parte 2: 50 unidades ≤ 100 ✓

**Nota importante:** El límite de unidades se aplica por cada combinación individualmente. Diferentes combinaciones pueden coexistir en el mismo archivo siempre que cada una cumpla su límite de unidades. El límite NO es la suma total de todas las combinaciones en el archivo, sino el máximo permitido para cada combinación individual.

**Nota:** Esta funcionalidad es opcional y puede omitirse. Si se omite, los archivos se dividen secuencialmente por número de registros (100 por archivo), sin considerar agrupación por combinaciones ni límites de unidades.

---

### 13. Generación de Archivos ZIP

**Ubicación:** Líneas 630-713

**Propósito:** Procesar datos consolidados y generar archivos Excel en lotes dentro de un ZIP.

**Flujo completo:**
1. Valida que haya datos para procesar
2. **Consolida datos** con `consolidate_by_brand_company()` (si está habilitado)
3. **Aplica empaquetado inteligente** (si está habilitado) para optimizar distribución
4. Calcula número de partes a generar según `chunk_size` o resultado del empaquetado inteligente
5. Crea `ProgressTracker` para mostrar progreso
6. Crea archivo ZIP en memoria (`BytesIO`)
7. Para cada parte:
   - Extrae chunk del DataFrame (consolidado y empaquetado)
   - Llama a `fill_one_chunk()` para generar Excel
   - Agrega archivo al ZIP con nombre `{prefix}{número}.xlsx`
8. Finaliza barra de progreso
9. Muestra botón de descarga del ZIP
10. Muestra resumen con métricas

**Características:**
- Divide datos en chunks del tamaño especificado (o según empaquetado inteligente)
- Cada chunk se escribe en un archivo Excel separado
- Todos los archivos se comprimen en un ZIP
- Muestra progreso visual durante el procesamiento
- Incluye métricas de limpieza y consolidación en el resumen

**Nombre de archivos generados:**
- Formato: `{default_prefix}{número}.xlsx`
- Ejemplo: `template_part01.xlsx`, `template_part02.xlsx`, ...
- ZIP: `{default_prefix}_lotes.zip`

---

## 📦 Dependencias

**Archivo:** `requirements.txt`

```
streamlit>=1.36
pandas>=2.1
openpyxl>=3.1
```

**Descripción:**
- **streamlit**: Framework web para la interfaz de usuario
- **pandas**: Manipulación y procesamiento de datos
- **openpyxl**: Lectura y escritura de archivos Excel (preserva formato)

---

## 🚀 Instalación y Uso

### Instalación Local

```bash
# Instalar dependencias
pip install -r requirements.txt

# Ejecutar aplicación
streamlit run app_streamlit_addi_v2.py
```

### Despliegue en Streamlit Cloud

1. Crear repositorio con:
   - `app_streamlit_addi_v2.py`
   - `requirements.txt`
   - `README.md` (opcional)

2. En Streamlit Cloud:
   - Seleccionar repositorio
   - Especificar archivo principal: `app_streamlit_addi_v2.py`

3. Configurar secretos (opcional pero recomendado):
   - Key: `APP_PASSWORD`
   - Value: Contraseña deseada (ej: `addi2025*`)
   - Si no se configura, la app usará `addi2025*` por defecto

---

## 📝 Columnas Requeridas en Archivo Origen

El archivo Excel origen debe contener las siguientes columnas (nombres exactos):

**Obligatorias:**
- `Brand Slug`
- `Nombre de la empresa`
- `Ciudad`
- `Departamento`
- `Dirección`
- `Nombre completo`
- `Referencia`
- `Número de tiendas`

**Opcionales (se procesan automáticamente si existen):**
- `Celular` (se autocompleta si está vacío)
- `Correo electrónico` (se limpia según reglas)

**Generadas automáticamente:**
- `Número de orden externo` (se genera desde Brand Slug + Empresa)

---

## 🔧 Configuración y Personalización

### Parámetros Configurables en Sidebar

- **Tamaño por archivo**: Número máximo de registros por archivo generado (default: 100)
- **Fila de encabezados**: Fila donde están los encabezados del template (default: 1)
- **Fila inicial de escritura**: Fila donde comenzar a escribir datos (default: 3)
- **Prefijo del nombre**: Prefijo para nombres de archivos generados (default: "template_part")

### Personalización del Código

Para personalizar la aplicación, consultar:
- **`Modificacion.md`**: Guía detallada de cómo solicitar modificaciones
- **`reglas.md`**: Reglas que el agente debe seguir al modificar código

**Secciones modificables:**
- Título y autenticación
- Bodegas y mapeo de ciudades/departamentos
- Reglas de limpieza y formateo
- Lógica de consolidación
- Empaquetado inteligente por registros y unidades (opcional)
- Notas de la aplicación

---

## 📊 Reglas de Procesamiento

### Asignación de Bodega

1. **Prioridad 1**: Busca ciudad normalizada en `CITY_TO_HUB`
2. **Prioridad 2**: Busca departamento normalizado en `DEPT_TO_HUB`
3. **Prioridad 3**: Busca keywords en ciudad o departamento
4. **Fallback**: Asigna primera bodega de `WAREHOUSES`

**Nota importante:** La ciudad se mantiene exactamente como viene del origen, solo se usa para determinar la bodega.

### Indicativo

- Solo se llena la **columna C** (índice 3) con valor **57**
- Si hay múltiples columnas "Indicativo", las demás se dejan vacías
- Si no existe columna C, se usa la primera columna "Indicativo" encontrada

### Consolidación

- Agrupa por **(Brand Slug, Nombre de la empresa)**
- Suma "Número de tiendas" de cada grupo
- Aplica tope máximo de **4 unidades** por grupo
- Genera "Número de orden externo" como `brand-empresa` (slug)

### Empaquetado Inteligente (Opcional)

- Agrupa registros por combinación (ej: Brand Slug + Nombre de la empresa)
- Controla máximo de **100 registros por archivo** (obligatorio)
- **Opcional:** Controla máximo de unidades totales por combinación en el archivo
- El límite de unidades se aplica POR COMBINACIÓN individualmente, no al total del archivo
- Diferentes combinaciones pueden coexistir en el mismo archivo siempre que cada una cumpla su límite individual

### Limpieza de Datos

- **Teléfonos vacíos**: Se autocompletan con número aleatorio (10 dígitos, inicia en 3)
- **Correos**: Solo se mantienen `@gmail.com` y `@hotmail.com` (minúscula), otros se eliminan
- **Número de orden externo**: Se genera automáticamente como slug `brand-empresa`

---

## 🐛 Solución de Problemas

### Error al leer archivo origen
- Verificar que el archivo sea `.xlsx` válido
- Verificar que la hoja seleccionada tenga datos
- Verificar que los nombres de columnas no tengan caracteres especiales problemáticos

### Error al leer template
- Verificar que el archivo sea `.xlsx` válido
- Verificar que la fila de encabezados (`header_row`) contenga los nombres de columnas
- Verificar que la hoja seleccionada exista

### Bodega no se asigna correctamente
- Verificar que la ciudad/departamento esté en `CITY_TO_HUB` o `DEPT_TO_HUB`
- Verificar que la ciudad en `WAREHOUSES` coincida con los valores en los diccionarios de mapeo
- Verificar normalización: los valores se comparan en minúscula y sin acentos

### Archivos generados vacíos
- Verificar que el mapeo de columnas esté correctamente configurado
- Verificar que las columnas del origen coincidan con las especificadas en el mapeo
- Verificar que la fila inicial de escritura (`start_row`) no sobrescriba encabezados

---

## 📚 Documentación Adicional

- **`Modificacion.md`**: Guía detallada para solicitar modificaciones al código
- **`reglas.md`**: Reglas técnicas para agentes que modifiquen el código

---

## 📄 Licencia y Créditos

Aplicación desarrollada para procesamiento de órdenes de Seller Addi.

**Versión:** 2.0

---

**FIN DEL DOCUMENTO**
