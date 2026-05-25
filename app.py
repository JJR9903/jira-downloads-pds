import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Jira CSV Processor", page_icon="📊", layout="wide")

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
# Makes buttons inside .chip-area look like removable tags.

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
    """Parse heterogeneous date strings → tz-naive datetime."""
    parsed = pd.to_datetime(series, errors="coerce", utc=True)
    return parsed.dt.tz_localize(None)


def apply_filters(df: pd.DataFrame, filters: list) -> pd.DataFrame:
    for f in filters:
        col   = f.get("column", "")
        mode  = f.get("mode", "include")   # "include" | "exclude"
        match = f.get("match", "exact")    # "exact" | "contains" | "startswith"
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
        col = r.get("column", "")
        old = r.get("old", "")
        new = r.get("new", "")
        if col and col in df.columns and old != "":
            df[col] = df[col].replace(old, new)
    return df


def process_dataframe(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    df = df.copy()

    df = apply_filters(df, config.get("filters", []))
    df = apply_replacements(df, config.get("replacements", []))

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

    for col in DATE_COLUMNS:
        if col in df.columns:
            df[col] = parse_date_column(df[col])

    return df[FINAL_SCHEMA]


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="YYYY-MM-DD") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
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

# ── Header ─────────────────────────────────────────────────────────────────────

st.title("📊 Jira CSV Processor")
st.markdown(
    "Upload Jira CSV exports → configure **column mappings**, **filters**, and "
    "**value replacements** per file → download a merged **XLSX** in the standard schema."
)

# ── File upload ────────────────────────────────────────────────────────────────

uploaded_files = st.file_uploader(
    "Upload one or more CSV files",
    type=["csv"],
    accept_multiple_files=True,
    key="uploader",
)

if not uploaded_files:
    st.info("Upload at least one CSV file to get started.")
    st.stop()

# Sync session state with current uploads
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

# ── Per-file configuration panels ─────────────────────────────────────────────

st.markdown(CHIP_CSS, unsafe_allow_html=True)

for f in uploaded_files:
    filename = f.name
    df       = st.session_state.file_data[filename]
    config   = st.session_state.file_configs[filename]
    src_opts = [""] + list(df.columns)

    with st.expander(
        f"**{filename}** · {len(df):,} rows × {len(df.columns)} cols",
        expanded=True,
    ):
        tab_map, tab_filt, tab_rep, tab_prev = st.tabs(
            ["🗂 Column Mapping", "🔍 Filters", "✏️ Value Replacements", "👁 Preview"]
        )

        # ── Column Mapping ─────────────────────────────────────────────────────
        with tab_map:
            st.caption(
                "For each target schema column select the matching source column. "
                "Leave blank to write an empty column in the output."
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

        # ── Filters ───────────────────────────────────────────────────────────
        with tab_filt:
            st.caption("Filters run **before** column mapping, on the original source columns.")

            if st.button("＋ Add filter", key=f"{filename}__add_filter"):
                config["filters"].append(
                    {"column": "", "mode": "include", "match": "exact", "values": []}
                )

            to_remove = []
            for i, filt in enumerate(config["filters"]):

                with st.container(border=True):
                    # ── Selector row ──────────────────────────────────────────
                    c1, c2, c3, c4 = st.columns([2.8, 1.4, 1.8, 0.5])

                    with c1:
                        prev_col = filt.get("column", "")
                        idx = src_opts.index(prev_col) if prev_col in src_opts else 0
                        new_col = st.selectbox(
                            "Source column",
                            src_opts,
                            index=idx,
                            key=f"{filename}__filt_{i}__col",
                        )
                        if new_col != prev_col:
                            filt["values"] = []
                        filt["column"] = new_col

                    with c2:
                        modes     = ["include", "exclude"]
                        mode_lbls = {"include": "Include", "exclude": "Exclude"}
                        midx = modes.index(filt.get("mode", "include"))
                        filt["mode"] = st.selectbox(
                            "Mode",
                            modes,
                            index=midx,
                            format_func=lambda k: mode_lbls[k],
                            key=f"{filename}__filt_{i}__mode",
                        )

                    with c3:
                        match_opts  = ["exact", "contains", "startswith"]
                        match_lbls  = {
                            "exact":      "Is (exact match)",
                            "contains":   "Contains",
                            "startswith": "Starts with",
                        }
                        prev_match = filt.get("match", "exact")
                        maidx = match_opts.index(prev_match) if prev_match in match_opts else 0
                        new_match = st.selectbox(
                            "Match type",
                            match_opts,
                            index=maidx,
                            format_func=lambda k: match_lbls[k],
                            key=f"{filename}__filt_{i}__match",
                        )
                        if new_match != prev_match:
                            filt["values"] = []
                        filt["match"] = new_match

                    with c4:
                        # Align delete button to bottom of row
                        st.markdown("<br><br>", unsafe_allow_html=True)
                        if st.button("🗑", key=f"{filename}__filt_{i}__rm", help="Remove filter"):
                            to_remove.append(i)

                    # ── Values area ───────────────────────────────────────────
                    col_name = filt.get("column", "")

                    if filt["match"] == "exact":
                        # Multi-select from the column's actual unique values
                        if col_name and col_name in df.columns:
                            unique_vals  = sorted(df[col_name].dropna().astype(str).unique())
                            valid_default = [v for v in filt.get("values", []) if v in unique_vals]
                            filt["values"] = st.multiselect(
                                "Values",
                                options=unique_vals,
                                default=valid_default,
                                placeholder="Select one or more values…",
                                key=f"{filename}__filt_{i}__multi__{col_name}",
                            )
                        else:
                            st.caption("← Select a source column to see available values.")

                    else:
                        # Tag / chip input
                        inp_key = f"{filename}__filt_{i}__chip_input"
                        if inp_key not in st.session_state:
                            st.session_state[inp_key] = ""

                        inp_col, btn_col = st.columns([5, 1])
                        with inp_col:
                            st.text_input(
                                "Add value",
                                label_visibility="collapsed",
                                placeholder=f"Type a value to match ({filt['match']}) and click ＋",
                                key=inp_key,
                            )
                        with btn_col:
                            if st.button("＋", key=f"{filename}__filt_{i}__add_chip", use_container_width=True):
                                new_v = st.session_state.get(inp_key, "").strip()
                                if new_v and new_v not in filt["values"]:
                                    filt["values"].append(new_v)
                                st.session_state[inp_key] = ""
                                st.rerun()

                        # Render chips
                        if filt["values"]:
                            st.markdown("<br>", unsafe_allow_html=True)
                            vals_to_drop = []
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
                                            vals_to_drop.append(global_idx)

                            if vals_to_drop:
                                for idx in reversed(vals_to_drop):
                                    filt["values"].pop(idx)
                                st.rerun()

            for i in reversed(to_remove):
                config["filters"].pop(i)

        # ── Value Replacements ────────────────────────────────────────────────
        with tab_rep:
            st.caption(
                "Replacements are applied **before** column mapping, on the original source columns."
            )
            if st.button("＋ Add replacement", key=f"{filename}__add_rep"):
                config["replacements"].append({"column": "", "old": "", "new": ""})

            to_remove_rep = []
            for i, rep in enumerate(config["replacements"]):
                r1, r2, r3, r4 = st.columns([2, 2.5, 2.5, 0.7])

                with r1:
                    idx = src_opts.index(rep["column"]) if rep["column"] in src_opts else 0
                    rep["column"] = st.selectbox(
                        "Source column",
                        src_opts,
                        index=idx,
                        key=f"{filename}__rep_{i}__col",
                    )
                with r2:
                    rep["old"] = st.text_input(
                        "Replace this value",
                        value=rep["old"],
                        key=f"{filename}__rep_{i}__old",
                    )
                with r3:
                    rep["new"] = st.text_input(
                        "With this value",
                        value=rep["new"],
                        key=f"{filename}__rep_{i}__new",
                    )
                with r4:
                    st.markdown("&nbsp;", unsafe_allow_html=True)
                    st.markdown("&nbsp;", unsafe_allow_html=True)
                    if st.button("🗑", key=f"{filename}__rep_{i}__rm", help="Remove"):
                        to_remove_rep.append(i)

            for i in reversed(to_remove_rep):
                config["replacements"].pop(i)

        # ── Preview ───────────────────────────────────────────────────────────
        with tab_prev:
            sub1, sub2 = st.tabs(["Raw data (first 10 rows)", "Processed data (first 50 rows)"])

            with sub1:
                st.dataframe(df.head(10), use_container_width=True)

            with sub2:
                if st.button("▶ Run preview", key=f"{filename}__preview_btn"):
                    try:
                        out = process_dataframe(df, config)
                        st.dataframe(out.head(50), use_container_width=True)
                        st.caption(f"{len(out):,} rows after filters · {len(out.columns)} columns")
                    except Exception as e:
                        st.error(f"Processing error: {e}")

# ── Merge & Download ───────────────────────────────────────────────────────────

st.divider()
st.subheader("Merge & Download")

left, right = st.columns(2)

with left:
    if st.button("⚙️ Process & Merge all files", type="primary", use_container_width=True):
        all_dfs, errors = [], []
        bar = st.progress(0, text="Processing…")

        for i, f in enumerate(uploaded_files):
            fname  = f.name
            df     = st.session_state.file_data[fname]
            config = st.session_state.file_configs[fname]
            try:
                processed = process_dataframe(df, config)
                all_dfs.append(processed)
            except Exception as e:
                errors.append(f"**{fname}**: {e}")
            bar.progress((i + 1) / len(uploaded_files), text=f"Processed {fname}")

        bar.empty()

        for err in errors:
            st.error(err)

        if all_dfs:
            final_df = pd.concat(all_dfs, ignore_index=True)
            st.session_state.xlsx_bytes = to_excel_bytes(final_df)
            st.session_state.final_shape = final_df.shape
            st.success(
                f"✅ Merged **{len(all_dfs)}** file(s) → "
                f"**{final_df.shape[0]:,} rows** × {final_df.shape[1]} columns"
            )

with right:
    if st.session_state.xlsx_bytes:
        rows, cols = st.session_state.final_shape
        st.download_button(
            label="⬇️ Download merged XLSX",
            data=st.session_state.xlsx_bytes,
            file_name="jira_merged.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help=f"{rows:,} rows × {cols} columns",
            use_container_width=True,
        )
    else:
        st.info("Click **Process & Merge** first, then the download button will appear here.")
