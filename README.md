
# Herramienta para crear órdenes – Seller Addi (Streamlit)

App de Streamlit para cargar un Excel **origen** y un **template**, mapear columnas y generar archivos en lotes de 100. 
Incluye asignación **rápida** de bodega por ciudad/departamento, respeta la **Ciudad** del origen, 
llena solo la **columna C (Indicativo)** con 57 (dejando las demás "Indicativo" vacías) y muestra una **barra de progreso**.

## Despliegue en Streamlit Cloud
1. Crea un repo con:
   - `app_streamlit_addi.py`
   - `requirements.txt`
   - (opcional) `README.md`
2. En Streamlit Cloud, selecciona el repo y el archivo principal `app_streamlit_addi.py`.
3. (Opcional, recomendado) Configura un **secret** para la contraseña:
   - Key: `APP_PASSWORD`
   - Value: `addi2025*` (o la que prefieras)
   - Si no configuras secret, la app usará `addi2025*` por defecto (hardcoded).

## Uso local
```bash
pip install -r requirements.txt
streamlit run app_streamlit_addi.py
```

## Notas
- La asignación de bodega es **rápida** (no geocodifica), basada en listas y heurística.
- Para cambiar el tamaño de lote, fila de encabezado o fila inicial, usa la barra lateral.
