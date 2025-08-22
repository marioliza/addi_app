
import io
from io import BytesIO
import zipfile
from pathlib import Path
from typing import Dict, Any, List, Tuple
import math
import os

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# =========================
# TITLE & AUTH
# =========================
APP_TITLE = "🧩 Herramienta para crear órdenes – Seller Addi"
APP_SUBTITLE = "Carga tu Excel origen y template, mapea columnas y descarga los archivos listos (100 por archivo)."
REQUIRED_PASSWORD = "addi2025*"  # Cambia esto si deseas otra clave. Puedes sobreescribir con st.secrets['APP_PASSWORD']

st.set_page_config(page_title="Seller Addi – Crear Órdenes", layout="wide")
st.title(APP_TITLE)
st.caption(APP_SUBTITLE)

# Password gate (simple)
try:
    pw = st.secrets["APP_PASSWORD"]  # usa secrets si existe
except Exception:
    pw = os.environ.get("APP_PASSWORD", REQUIRED_PASSWORD)  # o variable de entorno o fallback
if "auth_ok" not in st.session_state:
    st.session_state.auth_ok = False

if not st.session_state.auth_ok:
    st.subheader("🔒 Acceso")
    typed = st.text_input("Contraseña", type="password", placeholder="Ingresa la contraseña")
    if st.button("Entrar", type="primary"):
        if typed == pw:
            st.session_state.auth_ok = True
            st.success("Acceso concedido.")
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    st.stop()

# =========================
# PROGRESS
# =========================
class ProgressTracker:
    def __init__(self, total_rows, label="Procesando"):
        self.total = max(int(total_rows), 1)
        self.done = 0
        self.pbar = st.progress(0, text=f"{label} 0% (0/{self.total})")
    def add(self, n=1, label="Procesando"):
        self.done += int(n)
        if self.done > self.total:
            self.done = self.total
        frac = self.done / self.total
        pct = int(frac * 100)
        self.pbar.progress(frac, text=f"{label} {pct}% ({self.done}/{self.total})")
    def finish(self, label="Completado"):
        self.pbar.progress(1.0, text=f"{label} 100% ({self.total}/{self.total})")

# =========================
# CONFIG (ajustable en código)
# =========================
# Coordenadas/etiquetas de bodegas (solo para referencias; el modo rápido no las usa).
WAREHOUSES = [
    {"label": "Bogotá #2 - Montevideo", "city": "Bogotá"},
    {"label": "Medellín #2 - Sabaneta Mayorca", "city": "Medellín"},
]

# ===== Asignación RÁPIDA de bodega (sin geocodificación) =====
CITIES_MEDELLIN = {
    "medellin","medellín","sabaneta","itagui","itagüí","envigado","bello","la estrella","caldas",
    "girardota","copacabana","rionegro","el santuario","santuario","guarne","santo domingo",
    "santa fe de antioquia","amaga","barbosa","marinilla"
}
CITIES_BOGOTA = {
    "bogota","bogotá","funza","mosquera","soacha","chia","chía","cota","tocancipa","tocancipá",
    "zipaquira","zipaquirá","facatativa","facatativá","tenjo","siberia"
}
DEPT_MEDELLIN = {"antioquia"}
DEPT_BOGOTA = {"bogotá, d.c.","cundinamarca","bogota d.c.","bogotá d.c.","bogota, d.c."}

def assign_bodega_fast(row: pd.Series) -> str:
    city_val = str(row.get("Ciudad", "") or "").strip().lower()
    dept_val = str(row.get("Departamento", "") or "").strip().lower()
    if city_val in CITIES_MEDELLIN:
        return WAREHOUSES[1]["label"]
    if city_val in CITIES_BOGOTA:
        return WAREHOUSES[0]["label"]
    if dept_val in DEPT_MEDELLIN:
        return WAREHOUSES[1]["label"]
    if dept_val in DEPT_BOGOTA:
        return WAREHOUSES[0]["label"]
    if any(k in city_val for k in ["medellin","medellín","sabaneta","itagui","itagüí","envigado","bello"]):
        return WAREHOUSES[1]["label"]
    if any(k in city_val for k in ["bogota","bogotá","funza","mosquera","soacha","chia","chía","cota","zipaquira","zipaquirá"]):
        return WAREHOUSES[0]["label"]
    return WAREHOUSES[0]["label"]

# =========================
# SIDEBAR
# =========================
with st.sidebar:
    st.header("⚙️ Configuración")
    chunk_size = st.number_input("Tamaño por archivo (máx. registros)", min_value=1, value=100, step=1)
    header_row = st.number_input("Fila de encabezados del template", min_value=1, value=1, step=1)
    start_row = st.number_input("Fila inicial de escritura", min_value=1, value=3, step=1)  # Escribir desde A3
    default_prefix = st.text_input("Prefijo del nombre de salida", value="template_part")
    st.caption("La asignación de **Bodega** es en modo RÁPIDO por ciudad/departamento. **La Ciudad se mantiene tal cual del origen**.")

col_u1, col_u2 = st.columns(2)

# =========================
# UPLOAD SOURCE
# =========================
with col_u1:
    src_file = st.file_uploader("📥 Excel **origen** (.xlsx)", type=["xlsx"], key="src")
    src_sheet = None
    if src_file:
        try:
            xls = pd.ExcelFile(src_file)
            src_sheet = st.selectbox("Hoja de origen", xls.sheet_names, index=0, key="src_sheet")
            src_df = xls.parse(src_sheet, dtype=object)
            src_df.columns = [str(c).strip() for c in src_df.columns]
            st.success(f"Origen cargado. Filas: {len(src_df):,}. Columnas: {len(src_df.columns)}")
            with st.expander("Vista previa origen", expanded=False):
                st.dataframe(src_df.head(20))
        except Exception as e:
            st.error(f"Error leyendo origen: {e}")
            src_df = None
    else:
        src_df = None

# =========================
# UPLOAD TEMPLATE
# =========================
with col_u2:
    tmpl_file = st.file_uploader("📥 Excel **template** (.xlsx)", type=["xlsx"], key="tmpl")
    target_sheet = None
    headers: List[str] = []
    header_index: Dict[str, int] = {}
    header_positions: Dict[str, List[int]] = {}
    if tmpl_file:
        try:
            tmpl_bytes = tmpl_file.getvalue()
            wb = load_workbook(filename=BytesIO(tmpl_bytes), data_only=True)
            target_sheet = st.selectbox("Hoja del template", wb.sheetnames, index=0, key="tmpl_sheet")
            ws = wb[target_sheet]
            headers = []
            header_positions = {}
            for idx, cell in enumerate(ws[header_row], start=1):
                v = cell.value
                if v is None:
                    headers.append("")
                else:
                    name = str(v).strip()
                    headers.append(name)
                    header_positions.setdefault(name, []).append(idx)
            header_index = {name: i+1 for i, name in enumerate(headers) if name}
            st.success(f"Template cargado. Hoja '{target_sheet}'. Encabezados encontrados: {len(header_index)}")
            with st.expander(f"Encabezados del template (fila {header_row})", expanded=False):
                st.write([h for h in headers])
        except Exception as e:
            st.error(f"Error leyendo template: {e}")
            tmpl_bytes = None
    else:
        tmpl_bytes = None

st.markdown("---")
st.subheader("🧭 Mapeo de columnas (destino → origen / constante)")

# Mapeo por defecto (ajústalo en la UI si lo necesitas)
preset_mapping = {
    "Plantilla": {"mode": "template_name"},
    "Número de orden externo": {"mode": "source", "source_col": "Nombre de la empresa"},
    "Nombre completo del comprador": {"mode": "source", "source_col": "Nombre completo"},
    # "Indicativo": se forzará abajo (columna C) con 57, y otras se vacían
    "Teléfono de contacto": {"mode": "source", "source_col": "Celular"},
    "Correo electrónico": {"mode": "source", "source_col": "Correo electrónico"},
    "Tipo de empacado": {"mode": "const", "const_value": "Estandar"},
    "Igual al comprador": {"mode": "const", "const_value": "SI"},
    "Dirección": {"mode": "source", "source_col": "Dirección"},
    "Ciudad": {"mode": "source", "source_col": "Ciudad"},  # NO se cambia
    "Región": {"mode": "source", "source_col": "Departamento"},
    "País": {"mode": "const", "const_value": "Colombia"},
    "Método de envío": {"mode": "const", "const_value": "Estándar (Local y Nacional)"},
    "Tipo de recaudo": {"mode": "const", "const_value": "NO APLICA"},
    "SKU o Código Melonn del producto": {"mode": "source", "source_col": "Referencia"},
    "Cantidad": {"mode": "source", "source_col": "Número de tiendas"},
    # Bodega/CEDIS se llenará automáticamente por reglas
}

if "mapping_state" not in st.session_state:
    st.session_state.mapping_state = {}

def draw_mapping_ui(headers: List[str], src_cols: List[str]) -> Dict[str, Any]:
    mapping: Dict[str, Any] = {}
    modes = ["(no escribir)", "source", "const", "template_name", "source_filename"]
    col_left, col_right = st.columns(2)
    with col_left:
        st.caption("Destino (Template)")
    with col_right:
        st.caption("Origen / Valor")

    for dest in headers:
        if not dest:
            continue
        prev = st.session_state.mapping_state.get(dest, {})
        default_mode = "(no escribir)"
        default_source = src_cols[0] if src_cols else ""
        default_const = ""
        if dest in preset_mapping:
            pm = preset_mapping[dest]
            default_mode = pm.get("mode", default_mode)
            default_source = pm.get("source_col", default_source)
            default_const = str(pm.get("const_value", default_const))
        if prev:
            default_mode = prev.get("mode", default_mode)
            default_source = prev.get("source_col", default_source)
            default_const = str(prev.get("const_value", default_const))

        lock_dest = dest in ("Bodega", "CEDIS de origen")  # estos se calculan automático
        c1, c2 = st.columns([1, 2])
        with c1:
            mode = st.selectbox(
                f"{dest}",
                options=modes,
                index=modes.index(default_mode) if default_mode in modes else 0,
                key=f"mode::{dest}",
                disabled=lock_dest
            )
        with c2:
            if mode == "source":
                source_col = st.selectbox(
                    f"Columna origen para '{dest}'",
                    options=src_cols,
                    index=src_cols.index(default_source) if default_source in src_cols else 0 if src_cols else 0,
                    key=f"src::{dest}",
                    disabled=lock_dest
                )
                mapping[dest] = {"mode": "source", "source_col": source_col}
            elif mode == "const":
                const_val = st.text_input(
                    f"Valor fijo para '{dest}'",
                    value=default_const,
                    key=f"const::{dest}",
                    disabled=lock_dest
                )
                mapping[dest] = {"mode": "const", "const_value": const_val}
            elif mode == "template_name":
                st.text("(escribirá el nombre del template)")
                mapping[dest] = {"mode": "template_name"}
            elif mode == "source_filename":
                st.text("(escribirá el nombre del archivo origen)")
                mapping[dest] = {"mode": "source_filename"}
            else:
                st.text("—")
                mapping[dest] = {"mode": "(no escribir)"}

        st.session_state.mapping_state[dest] = mapping[dest]

    return mapping

src_cols_list = list(src_df.columns) if src_df is not None else []
mapping = draw_mapping_ui(list(header_index.keys()) if header_index else [], src_cols_list)

# =========================
# HELPERS
# =========================
def resolve_value(spec: Dict[str, Any], row: pd.Series, template_name: str, source_name: str):
    mode = spec.get("mode", "(no escribir)")
    if mode == "source":
        col = spec.get("source_col", "")
        return row.get(col, None)
    elif mode == "const":
        val = spec.get("const_value", "")
        # intenta parsear numérico si aplica
        try:
            f = float(val)
            if f.is_integer():
                return int(f)
            return f
        except Exception:
            return val
    elif mode == "template_name":
        return template_name
    elif mode == "source_filename":
        return source_name
    else:
        return None

def fill_one_chunk(
    tmpl_bytes: bytes,
    target_sheet: str,
    header_index: Dict[str, int],
    header_positions: Dict[str, List[int]],
    start_row: int,
    chunk_df: pd.DataFrame,
    mapping: Dict[str, Any],
    template_name: str,
    source_name: str,
    prog: ProgressTracker,
) -> Tuple[bytes, Dict[str, int]]:
    wb = load_workbook(filename=BytesIO(tmpl_bytes))
    if target_sheet not in wb.sheetnames:
        raise KeyError(f"La hoja '{target_sheet}' no existe en el template.")
    ws = wb[target_sheet]

    stats = {"rows": 0, "nw_written": 0, "no_dest_bodega": 0}

    # Detectar columna destino para bodega (prioriza 'Bodega', luego 'CEDIS de origen')
    dest_bodega = None
    if "Bodega" in header_index:
        dest_bodega = "Bodega"
    elif "CEDIS de origen" in header_index:
        dest_bodega = "CEDIS de origen"

    for r_offset, (_, row) in enumerate(chunk_df.iterrows()):
        row_idx = start_row + r_offset

        # 1) Mapeo normal (Ciudad queda del origen)
        for dest, spec in mapping.items():
            if spec.get("mode") == "(no escribir)":
                continue
            if dest not in header_index:
                continue
            c_idx = header_index[dest]
            value = resolve_value(spec, row, template_name, source_name)
            ws.cell(row=row_idx, column=c_idx, value=value)

        # 2) Bodega automática (rápida)
        b_label = assign_bodega_fast(row)
        if dest_bodega and dest_bodega in header_index:
            c_bod = header_index[dest_bodega]
            ws.cell(row=row_idx, column=c_bod, value=b_label)
            stats["nw_written"] += 1
        else:
            stats["no_dest_bodega"] += 1

        # 3) Indicativo: llenar SOLO la primera 'Indicativo' en C (col 3). Vaciar otras 'Indicativo'.
        indic_idxs = []
        for name, idxs in header_positions.items():
            if str(name).strip().lower() == "indicativo":
                indic_idxs.extend(idxs)
        if indic_idxs:
            keep_idx = 3 if 3 in indic_idxs else indic_idxs[0]
            for c_idx in indic_idxs:
                if c_idx == keep_idx:
                    ws.cell(row=row_idx, column=c_idx, value=57)
                else:
                    ws.cell(row=row_idx, column=c_idx, value=None)

        stats["rows"] += 1
        try:
            prog.add(1, label="Procesando")
        except Exception:
            pass

    out_buf = BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue(), stats

# =========================
# GENERATE
# =========================
st.markdown("---")
st.subheader("🚀 Generar archivos")
do_run = st.button("Procesar y generar ZIP", type="primary", disabled=(src_df is None or tmpl_bytes is None or not header_index))

if do_run:
    try:
        if src_df is None or tmpl_bytes is None:
            st.stop()
        total = len(src_df)
        if total == 0:
            st.warning("El origen no tiene filas para procesar.")
            st.stop()

        num_parts = (total + chunk_size - 1) // chunk_size
        st.info(f"Total filas: {total}. Tamaño de bloque: {chunk_size}. Partes a generar: {num_parts}.")

        template_stem = Path(getattr(tmpl_file, "name", "template.xlsx")).stem
        source_stem = Path(getattr(src_file, "name", "origen.xlsx")).stem

        zip_buf = BytesIO()
        agg = {"rows": 0, "nw_written": 0, "no_dest_bodega": 0}
        prog = ProgressTracker(total_rows=total, label="Procesando")

        with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for i in range(num_parts):
                start = i * chunk_size
                end = min(start + chunk_size, total)
                chunk = src_df.iloc[start:end].copy()
                out_xlsx, stats = fill_one_chunk(
                    tmpl_bytes=tmpl_bytes,
                    target_sheet=target_sheet,
                    header_index=header_index,
                    header_positions=header_positions,
                    start_row=int(start_row),
                    chunk_df=chunk,
                    mapping=mapping,
                    template_name=template_stem,
                    source_name=source_stem,
                    prog=prog,
                )
                for k in agg:
                    agg[k] += stats.get(k, 0)
                part_name = f"{default_prefix}{i+1:02d}.xlsx"
                zf.writestr(part_name, out_xlsx)

        try:
            prog.finish(label="Completado")
        except Exception:
            pass

        zip_buf.seek(0)
        st.success("¡Listo! Descarga tu archivo ZIP con los templates llenos (Bodega auto-asignada; Ciudad intacta).")
        st.download_button(
            "⬇️ Descargar ZIP",
            data=zip_buf.getvalue(),
            file_name=f"{default_prefix}_lotes.zip",
            mime="application/zip",
        )

        with st.expander("Resumen de procesamiento", expanded=False):
            st.write(agg)
    except Exception as e:
        st.error(f"ERROR: {e}")

with st.expander("📝 Notas", expanded=False):
    st.markdown("""
    - **Seller**: Addi · **Función**: crear órdenes desde un Excel.
    - **Bodega**: se escribe en 'Bodega' (si existe) o 'CEDIS de origen'. Se decide por reglas rápidas (ciudad/departamento).
    - **Ciudad**: se mantiene exactamente como viene del **origen**.
    - **Indicativo**: solo se llena la **columna C** (si su encabezado es 'Indicativo') con **57**; otras 'Indicativo' se dejan vacías.
    - **Escritura**: inicia en **A3** (configurable) y divide en archivos de **100** registros por defecto.
    """)
