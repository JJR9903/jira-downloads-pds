import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Procesador Jira CSV", page_icon="📊", layout="wide")

# ── Schema & constants ─────────────────────────────────────────────────────────

FINAL_SCHEMA = [
    "Issue Type",
    "Issue key",
    "Issue id",
    "Summary",
    "Custom field (Project / Work ID)",
    "Status",
    "Custom field (Start date)",
    "Due date",
    "Custom field (Story Points)",
    "Assignee",
    "Assignee Id",
    "Reporter",
    "Reporter Id",
    "Created",
    "Updated",
    "Priority",
    "Resolved",
    "Parent",
    "Parent key",
    "Parent summary",
    "Custom field (Project Phase)",
    "Custom field (Execution Status)",
]

DATE_COLUMNS = {
    "Custom field (Start date)",
    "Due date",
    "Created",
    "Updated",
    "Resolved",
}

# ── Chip CSS ───────────────────────────────────────────────────────────────────

CHIP_CSS = """
<style>
div[data-chip-area] button, .chip-area button {
    border-radius: 2rem !important;
    padding: 0 10px !important;
    font-size: 0.78rem !important;
    height: 1.9rem !important;
    background: #1c3f6e !important;
    border: 1px solid #2d6aa0 !important;
    color: #90caff !important;
    min-height: 0 !important;
}
</style>
"""

# ── Core processing helpers ────────────────────────────────────────────────────

def read_csv_safe(file) -> pd.DataFrame:
    try:
        return pd.read_csv(file, dtype=str)
    except UnicodeDecodeError:
        file.seek(0)
        return pd.read_csv(file, dtype=str, encoding="latin-1")


def parse_date_column(series: pd.Series) -> pd.Series:
    """Parse Jira date strings → tz-naive datetime.

    Handles (in the same column):
      - "16/Apr/26 4:02 PM"   DD/MMM/YY h:mm AM/PM
      - "01/Jan/26 12:00 AM"  DD/MMM/YY 12:00 AM/PM
      - "16/Apr/26"           DD/MMM/YY  (date only)
      - ISO-8601 variants, timezone-aware strings, empty / NaN cells
    """
    try:
        parsed = pd.to_datetime(series, format="mixed", dayfirst=True, errors="coerce")
    except TypeError:
        parsed = pd.to_datetime(series, infer_datetime_format=True, dayfirst=True, errors="coerce")

    if parsed.dt.tz is not None:
        parsed = parsed.dt.tz_convert("UTC").dt.tz_localize(None)

    return parsed.dt.normalize()


def apply_filters(df: pd.DataFrame, filters: list) -> pd.DataFrame:
    for f in filters:
        col   = f.get("column", "")
        mode  = f.get("mode", "include")
        match = f.get("match", "exact")
        vals  = f.get("values", [])

        if not col or col not in df.columns or not vals:
            continue

        s = df[col].fillna("").astype(str)

        if match == "exact":
            mask = s.isin(vals)
        elif match == "contains":
            mask = s.apply(lambda x: any(v.lower() in x.lower() for v in vals))
        elif match == "startswith":
            mask = s.apply(lambda x: any(x.lower().startswith(v.lower()) for v in vals))
        else:
            continue

        df = df[mask] if mode == "include" else df[~mask]

    return df.reset_index(drop=True)


def apply_replacements(df: pd.DataFrame, replacements: list) -> pd.DataFrame:
    for r in replacements:
        col  = r.get("column", "")
        mode = r.get("mode", "specific")
        old  = r.get("old", "")
        new  = r.get("new", "")
        if not col or col not in df.columns:
            continue
        if mode == "overwrite":
            df[col] = new
        elif old != "":
            df[col] = df[col].replace(old, new)
    return df


def process_dataframe(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    df = df.copy()

    # 1. Column mapping
    col_mapping = {
        src: tgt
        for src, tgt in config.get("column_mapping", {}).items()
        if src and tgt and src in df.columns
    }
    if not col_mapping:
        return pd.DataFrame(columns=FINAL_SCHEMA)
    df = df[list(col_mapping.keys())].rename(columns=col_mapping)
    for col in FINAL_SCHEMA:
        if col not in df.columns:
            df[col] = pd.NA

    # 2. Filters — on schema columns (after mapping)
    df = apply_filters(df, config.get("filters", []))

    # 3. Replacements — on schema columns (after mapping + filters)
    df = apply_replacements(df, config.get("replacements", []))

    for col in DATE_COLUMNS:
        if col in df.columns:
            df[col] = parse_date_column(df[col])

    return df[FINAL_SCHEMA]


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.date
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Datos")
    return buf.getvalue()


# ── Session state ──────────────────────────────────────────────────────────────

for key, default in [
    ("file_configs", {}),
    ("file_data",    {}),
    ("xlsx_bytes",   None),
    ("final_shape",  None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ── Encabezado & navegación ────────────────────────────────────────────────────

st.title("📊 Procesador Jira CSV")

tab_procesador, tab_baseline = st.tabs(["📊 Procesador", "📋 Baseline"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB: Procesador
# ══════════════════════════════════════════════════════════════════════════════

with tab_procesador:
    st.markdown(
        "Sube exportaciones CSV de Jira → configura **mapeo de columnas**, **filtros** y "
        "**reemplazos de valores** por archivo → descarga un **XLSX** combinado con el esquema estándar."
    )

    # ── Carga de archivos ──────────────────────────────────────────────────────

    uploaded_files = st.file_uploader(
        "Sube uno o más archivos CSV",
        type=["csv"],
        accept_multiple_files=True,
        key="uploader",
    )

    if not uploaded_files:
        st.info("Sube al menos un archivo CSV para comenzar.")
    else:
        # Sincronizar session state con los archivos cargados
        current_names = {f.name for f in uploaded_files}

        for name in list(st.session_state.file_data.keys()):
            if name not in current_names:
                del st.session_state.file_data[name]
                del st.session_state.file_configs[name]

        for f in uploaded_files:
            if f.name not in st.session_state.file_data:
                df = read_csv_safe(f)
                st.session_state.file_data[f.name] = df
                schema_set   = set(FINAL_SCHEMA)
                auto_mapping = {col: col for col in df.columns if col in schema_set}
                st.session_state.file_configs[f.name] = {
                    "column_mapping": auto_mapping,
                    "filters":        [],
                    "replacements":   [],
                }

        # ── Paneles de configuración por archivo ───────────────────────────────

        st.markdown(CHIP_CSS, unsafe_allow_html=True)

        for f in uploaded_files:
            filename = f.name
            df       = st.session_state.file_data[filename]
            config   = st.session_state.file_configs[filename]
            src_opts = [""] + list(df.columns)

            with st.expander(
                f"**{filename}** · {len(df):,} filas × {len(df.columns)} cols",
                expanded=True,
            ):
                tab_map, tab_filt, tab_rep, tab_prev = st.tabs(
                    ["🗂 Mapeo de Columnas", "🔍 Filtros", "✏️ Reemplazos de Valores", "👁 Vista Previa"]
                )

                # ── Mapeo de Columnas ──────────────────────────────────────────
                with tab_map:
                    st.caption(
                        "Para cada columna del esquema destino, selecciona la columna origen correspondiente. "
                        "Déjala en blanco para escribir una columna vacía en el resultado."
                    )
                    rev_map     = {tgt: src for src, tgt in config["column_mapping"].items()}
                    new_mapping = {}

                    pairs = list(zip(FINAL_SCHEMA[::2], FINAL_SCHEMA[1::2]))
                    if len(FINAL_SCHEMA) % 2:
                        pairs.append((FINAL_SCHEMA[-1], None))

                    for left, right in pairs:
                        col_a, col_b = st.columns(2)
                        for target_col, col_widget in [(left, col_a), (right, col_b)]:
                            if target_col is None:
                                continue
                            with col_widget:
                                cur_src = rev_map.get(target_col, "")
                                idx = src_opts.index(cur_src) if cur_src in src_opts else 0
                                chosen = st.selectbox(
                                    target_col,
                                    src_opts,
                                    index=idx,
                                    key=f"{filename}__map__{target_col}",
                                )
                                if chosen:
                                    new_mapping[chosen] = target_col

                    config["column_mapping"] = new_mapping

                # ── Filtros ───────────────────────────────────────────────────
                with tab_filt:
                    st.caption("Los filtros se aplican **después** del mapeo de columnas, sobre las columnas del esquema destino.")

                    if st.button("＋ Agregar filtro", key=f"{filename}__add_filter"):
                        config["filters"].append(
                            {"column": "", "mode": "include", "match": "exact", "values": []}
                        )

                    filt_opts = [""] + FINAL_SCHEMA
                    rev_map   = {tgt: src for src, tgt in config["column_mapping"].items()}

                    to_remove = []
                    for i, filt in enumerate(config["filters"]):

                        with st.container(border=True):
                            c1, c2, c3, c4 = st.columns([2.8, 1.4, 1.8, 0.5])

                            with c1:
                                prev_col = filt.get("column", "")
                                idx = filt_opts.index(prev_col) if prev_col in filt_opts else 0
                                new_col = st.selectbox(
                                    "Columna destino",
                                    filt_opts,
                                    index=idx,
                                    key=f"{filename}__filt_{i}__col",
                                )
                                if new_col != prev_col:
                                    filt["values"] = []
                                filt["column"] = new_col

                            with c2:
                                modes     = ["include", "exclude"]
                                mode_lbls = {"include": "Incluir", "exclude": "Excluir"}
                                midx = modes.index(filt.get("mode", "include"))
                                filt["mode"] = st.selectbox(
                                    "Modo",
                                    modes,
                                    index=midx,
                                    format_func=lambda k: mode_lbls[k],
                                    key=f"{filename}__filt_{i}__mode",
                                )

                            with c3:
                                match_opts = ["exact", "contains", "startswith"]
                                match_lbls = {
                                    "exact":      "Es (exacto)",
                                    "contains":   "Contiene",
                                    "startswith": "Comienza con",
                                }
                                prev_match = filt.get("match", "exact")
                                maidx = match_opts.index(prev_match) if prev_match in match_opts else 0
                                new_match = st.selectbox(
                                    "Tipo de coincidencia",
                                    match_opts,
                                    index=maidx,
                                    format_func=lambda k: match_lbls[k],
                                    key=f"{filename}__filt_{i}__match",
                                )
                                if new_match != prev_match:
                                    filt["values"] = []
                                filt["match"] = new_match

                            with c4:
                                st.markdown("<br><br>", unsafe_allow_html=True)
                                if st.button("🗑", key=f"{filename}__filt_{i}__rm", help="Eliminar filtro"):
                                    to_remove.append(i)

                            col_name = filt.get("column", "")

                            if filt["match"] == "exact":
                                src_col = rev_map.get(col_name, col_name)
                                if col_name and src_col in df.columns:
                                    unique_vals   = sorted(df[src_col].dropna().astype(str).unique())
                                    valid_default = [v for v in filt.get("values", []) if v in unique_vals]
                                    filt["values"] = st.multiselect(
                                        "Valores",
                                        options=unique_vals,
                                        default=valid_default,
                                        placeholder="Selecciona uno o más valores…",
                                        key=f"{filename}__filt_{i}__multi__{col_name}",
                                    )
                                    if not filt["values"]:
                                        st.caption("⚠️ Sin valores seleccionados — filtro inactivo.")
                                else:
                                    st.caption("← Selecciona una columna destino para ver los valores disponibles.")

                            else:
                                inp_key = f"{filename}__filt_{i}__chip_input"
                                match_es = {"contains": "contiene", "startswith": "comienza con"}

                                def _add_chip(
                                    _key=inp_key,
                                    _values=filt["values"],
                                ):
                                    val = st.session_state.get(_key, "").strip()
                                    if val and val not in _values:
                                        _values.append(val)
                                    st.session_state[_key] = ""

                                inp_col, btn_col = st.columns([5, 1])
                                with inp_col:
                                    st.text_input(
                                        "Agregar valor",
                                        label_visibility="collapsed",
                                        placeholder=f"Escribe un valor ({match_es.get(filt['match'], filt['match'])}) y haz clic en ＋",
                                        key=inp_key,
                                    )
                                with btn_col:
                                    st.button(
                                        "＋",
                                        key=f"{filename}__filt_{i}__add_chip",
                                        on_click=_add_chip,
                                        use_container_width=True,
                                    )

                                if filt["values"]:
                                    st.markdown("<br>", unsafe_allow_html=True)
                                    per_row = 5
                                    rows = [
                                        filt["values"][s: s + per_row]
                                        for s in range(0, len(filt["values"]), per_row)
                                    ]
                                    for r_idx, row_vals in enumerate(rows):
                                        chip_cols = st.columns(per_row)
                                        for j, val in enumerate(row_vals):
                                            global_idx = r_idx * per_row + j
                                            with chip_cols[j]:
                                                if st.button(
                                                    f"✕  {val}",
                                                    key=f"{filename}__filt_{i}__chip_{global_idx}",
                                                    use_container_width=True,
                                                ):
                                                    filt["values"].pop(global_idx)
                                else:
                                    st.caption("⚠️ Sin valores agregados — filtro inactivo.")

                    for i in reversed(to_remove):
                        config["filters"].pop(i)

                # ── Reemplazos de Valores ──────────────────────────────────────
                with tab_rep:
                    st.caption(
                        "Los reemplazos se aplican **después** del mapeo de columnas y los filtros, sobre las columnas del esquema destino."
                    )
                    if st.button("＋ Agregar reemplazo", key=f"{filename}__add_rep"):
                        config["replacements"].append({"column": "", "mode": "specific", "old": "", "new": ""})

                    rep_opts = [""] + FINAL_SCHEMA
                    to_remove_rep = []
                    for i, rep in enumerate(config["replacements"]):
                        with st.container(border=True):
                            r1, r2, r3 = st.columns([2.5, 2.5, 0.5])

                            with r1:
                                idx = rep_opts.index(rep["column"]) if rep["column"] in rep_opts else 0
                                rep["column"] = st.selectbox(
                                    "Columna destino",
                                    rep_opts,
                                    index=idx,
                                    key=f"{filename}__rep_{i}__col",
                                )
                            with r2:
                                modes     = ["specific", "overwrite"]
                                mode_lbls = {
                                    "specific":  "Reemplazar un valor específico",
                                    "overwrite": "Establecer toda la columna a un valor fijo",
                                }
                                midx = modes.index(rep.get("mode", "specific"))
                                rep["mode"] = st.selectbox(
                                    "Tipo de reemplazo",
                                    modes,
                                    index=midx,
                                    format_func=lambda k: mode_lbls[k],
                                    key=f"{filename}__rep_{i}__mode",
                                )
                            with r3:
                                st.markdown("<br><br>", unsafe_allow_html=True)
                                if st.button("🗑", key=f"{filename}__rep_{i}__rm", help="Eliminar"):
                                    to_remove_rep.append(i)

                            if rep["mode"] == "specific":
                                v1, v2 = st.columns(2)
                                with v1:
                                    rep["old"] = st.text_input(
                                        "Reemplazar este valor",
                                        value=rep.get("old", ""),
                                        key=f"{filename}__rep_{i}__old",
                                    )
                                with v2:
                                    rep["new"] = st.text_input(
                                        "Con este valor",
                                        value=rep.get("new", ""),
                                        key=f"{filename}__rep_{i}__new",
                                    )
                            else:
                                rep["new"] = st.text_input(
                                    "Valor fijo a establecer",
                                    value=rep.get("new", ""),
                                    key=f"{filename}__rep_{i}__new_ow",
                                )

                    for i in reversed(to_remove_rep):
                        config["replacements"].pop(i)

                # ── Vista Previa ───────────────────────────────────────────────
                with tab_prev:
                    sub1, sub2 = st.tabs(["Datos originales (primeras 10 filas)", "Datos procesados (primeras 50 filas)"])

                    with sub1:
                        st.dataframe(df.head(10), use_container_width=True)

                    with sub2:
                        prev_key = f"{filename}__preview_result"
                        prev_err = f"{filename}__preview_error"

                        if st.button("▶ Ver vista previa", key=f"{filename}__preview_btn"):
                            try:
                                out = process_dataframe(df, config)
                                st.session_state[prev_key] = out
                                st.session_state[prev_err] = None
                            except Exception as e:
                                st.session_state[prev_key] = None
                                st.session_state[prev_err] = str(e)

                        if st.session_state.get(prev_err):
                            st.error(f"Error de procesamiento: {st.session_state[prev_err]}")
                        elif prev_key in st.session_state and st.session_state[prev_key] is not None:
                            out = st.session_state[prev_key]
                            st.dataframe(out.head(50), use_container_width=True)
                            st.caption(f"{len(out):,} filas · {len(out.columns)} cols — haz clic en ▶ para actualizar tras los cambios")
                        else:
                            st.info("Haz clic en ▶ Ver vista previa para ver el resultado procesado.")

        # ── Combinar y Descargar ───────────────────────────────────────────────

        st.divider()
        st.subheader("Combinar y Descargar")

        left, right = st.columns(2)

        with left:
            if st.button("⚙️ Procesar y combinar todos los archivos", type="primary", use_container_width=True):
                all_dfs, errors = [], []
                bar = st.progress(0, text="Procesando…")

                for i, f in enumerate(uploaded_files):
                    fname  = f.name
                    df     = st.session_state.file_data[fname]
                    config = st.session_state.file_configs[fname]
                    try:
                        processed = process_dataframe(df, config)
                        all_dfs.append(processed)
                    except Exception as e:
                        errors.append(f"**{fname}**: {e}")
                    bar.progress((i + 1) / len(uploaded_files), text=f"Procesado {fname}")

                bar.empty()

                for err in errors:
                    st.error(err)

                if all_dfs:
                    final_df = pd.concat(all_dfs, ignore_index=True)
                    st.session_state.xlsx_bytes = to_excel_bytes(final_df)
                    st.session_state.final_shape = final_df.shape
                    st.success(
                        f"✅ Combinados **{len(all_dfs)}** archivo(s) → "
                        f"**{final_df.shape[0]:,} filas** × {final_df.shape[1]} columnas"
                    )

        with right:
            if st.session_state.xlsx_bytes:
                rows, cols = st.session_state.final_shape
                st.download_button(
                    label="⬇️ Descargar XLSX combinado",
                    data=st.session_state.xlsx_bytes,
                    file_name="jira_combinado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help=f"{rows:,} filas × {cols} columnas",
                    use_container_width=True,
                )
            else:
                st.info("Haz clic en **Procesar y combinar** primero; luego aparecerá el botón de descarga.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB: Baseline
# ══════════════════════════════════════════════════════════════════════════════

with tab_baseline:
    st.markdown(
        "Sube el archivo Excel de Baseline para cargar las hojas "
        "**Baseline** (columnas A:B) y **FI_Baseline** (columnas A:C)."
    )

    baseline_file = st.file_uploader(
        "Sube el archivo Excel de Baseline",
        type=["xlsx", "xls"],
        key="baseline_uploader",
    )

    if baseline_file:
        try:
            df_baseline = pd.read_excel(baseline_file, sheet_name="Baseline", usecols="A:B", dtype=str)
            baseline_file.seek(0)
            df_fi = pd.read_excel(baseline_file, sheet_name="FI_Baseline", usecols="A:C", dtype=str)

            st.markdown("### Hoja: Baseline")
            st.dataframe(df_baseline, use_container_width=True)
            st.caption(f"{len(df_baseline):,} filas · {len(df_baseline.columns)} columnas")

            st.markdown("### Hoja: FI_Baseline")
            st.dataframe(df_fi, use_container_width=True)
            st.caption(f"{len(df_fi):,} filas · {len(df_fi.columns)} columnas")

        except KeyError as e:
            st.error(
                f"No se encontró la hoja {e} en el archivo. "
                "Verifica que el Excel tenga hojas llamadas 'Baseline' y 'FI_Baseline'."
            )
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")
    else:
        st.info("Sube un archivo Excel (.xlsx) para comenzar.")
