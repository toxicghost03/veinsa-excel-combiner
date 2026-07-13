import io
import re
import pandas as pd
import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Veinsa – Combinador de Excels",
    page_icon="🚗",
    layout="centered",
)

COLUMNA_PLACA = "TAG"

# ── Header ────────────────────────────────────────────────────────────────────
st.title("🚗 Combinador de Excels — Veinsa")
st.markdown(
    "Sube los exports de DealerApps, combínalos en un solo archivo y "
    "descárgalo listo para usar."
)
st.divider()


# ── File upload ───────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Sube los archivos Excel / CSV de DealerApps",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
    help="Puedes subir varios archivos a la vez (Open RO, Finished RO, etc.)",
)

if not uploaded:
    st.info("Sube al menos un archivo para comenzar.")
    st.stop()

st.markdown(f"**{len(uploaded)} archivo(s) cargado(s)**")


# ── Process ───────────────────────────────────────────────────────────────────
def read_file(f):
    name = f.name.lower()
    raw = f.read()
    f.seek(0)

    if name.endswith(".csv"):
        return pd.read_csv(f)

    elif name.endswith(".xls"):
        # DealerApps exports .xls files that are actually HTML tables.
        # Detect by sniffing the first bytes — real XLS starts with 0xD0CF.
        is_html = not raw[:8].startswith(b"\xd0\xcf\x11\xe0")
        if is_html:
            import io
            dfs = pd.read_html(io.BytesIO(raw), header=0)
            return dfs[0] if dfs else None
        else:
            return pd.read_excel(f, engine="xlrd")

    else:
        return pd.read_excel(f)


def find_tag_column(df):
    df.columns = df.columns.astype(str).str.strip()
    if COLUMNA_PLACA in df.columns:
        return df, None
    for i in range(min(10, len(df))):
        if COLUMNA_PLACA in [str(v).strip() for v in df.iloc[i].values]:
            df.columns = df.iloc[i].values
            df = df.iloc[i + 1:].reset_index(drop=True)
            df.columns = df.columns.astype(str).str.strip()
            return df, None
    return None, df.columns.tolist()


lista_dfs = []
errores = []

for f in uploaded:
    try:
        df = read_file(f)
        df, cols_encontradas = find_tag_column(df)
        if df is None:
            errores.append((f.name, cols_encontradas))
        else:
            lista_dfs.append((f.name, df))
    except Exception as e:
        errores.append((f.name, [f"ERROR: {e}"]))

for nombre, cols in errores:
    if cols:
        st.warning(
            f"**{nombre}** — ignorado, no tiene columna **TAG** (placa). "
            f"Este archivo es probablemente un reporte diferente. "
            f"Columnas encontradas: `{', '.join(str(c) for c in cols[:12])}`\n\n"
            f"Asegurate de exportar el reporte **Open Repair Orders** o **Finished Repair Orders** "
            f"que incluya la columna TAG."
        )
    else:
        st.error(f"**{nombre}** — no se pudo leer el archivo.")

if not lista_dfs:
    st.error("Ningún archivo tiene la columna TAG. Revisa los archivos subidos.")
    st.stop()

# ── Summary cards ─────────────────────────────────────────────────────────────
cols = st.columns(len(lista_dfs))
for col, (nombre, df) in zip(cols, lista_dfs):
    col.metric(nombre, f"{len(df):,} filas")

# ── Combine ───────────────────────────────────────────────────────────────────
df_combined = pd.concat([df for _, df in lista_dfs], ignore_index=True)

df_combined["Placa_Match"] = (
    df_combined[COLUMNA_PLACA]
    .astype(str)
    .str.replace(r"[^A-Z0-9]", "", regex=True)
    .str.upper()
)
if "VIN NO" in df_combined.columns:
    df_combined["VIN_Match"] = (
        df_combined["VIN NO"]
        .astype(str)
        .str.replace(r"[^A-HJ-NPR-Z0-9]", "", regex=True)
        .str.upper()
    )

st.divider()
st.subheader(f"Vista previa — {len(df_combined):,} registros en total")

search = st.text_input("Buscar placa o cliente...", placeholder="ej. ABC123")
df_view = df_combined.copy()
if search:
    mask = df_view.astype(str).apply(
        lambda col: col.str.contains(search.upper(), case=False, na=False)
    ).any(axis=1)
    df_view = df_view[mask]
    st.caption(f"{len(df_view):,} resultado(s) para '{search}'")

st.dataframe(df_view.fillna(""), use_container_width=True, height=380)

# ── Build Excel in memory ─────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def build_excel(data_json: str) -> bytes:
    df = pd.read_json(io.StringIO(data_json), orient="records")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df_out = df.fillna("")
        df_out.to_excel(writer, sheet_name="Sistema_Completo", index=False)
        ws = writer.sheets["Sistema_Completo"]
        wb = writer.book
        for i, col in enumerate(df_out.columns):
            ancho = max(df_out[col].astype(str).map(len).max(), len(str(col))) + 2
            ws.set_column(i, i, min(ancho, 50))
        ws.add_table(0, 0, len(df_out), len(df_out.columns) - 1, {
            "columns": [{"header": c} for c in df_out.columns],
            "style": "Table Style Medium 9",
        })

        # ── Hoja dedicada: solo OTs FACTURADAS ──
        _status_col = next((c for c in df_out.columns if str(c).strip().upper() in ("STATUS", "ESTATUS", "ESTADO")), None)
        if _status_col is not None:
            df_fact = df_out[df_out[_status_col].astype(str).str.upper().str.contains("FACTURAD", na=False)]
        else:
            df_fact = df_out.iloc[0:0]
        df_fact.to_excel(writer, sheet_name="Facturadas", index=False)
        if len(df_fact) > 0:
            ws_f = writer.sheets["Facturadas"]
            for i, col in enumerate(df_fact.columns):
                ancho = max(df_fact[col].astype(str).map(len).max(), len(str(col))) + 2
                ws_f.set_column(i, i, min(ancho, 50))
            ws_f.add_table(0, 0, len(df_fact), len(df_fact.columns) - 1, {
                "columns": [{"header": c} for c in df_fact.columns],
                "style": "Table Style Medium 3",
            })
    return buf.getvalue()


st.divider()

with st.spinner("Generando archivo..."):
    excel_bytes = build_excel(df_combined.to_json(orient="records"))

_status_col_ui = next((c for c in df_combined.columns if str(c).strip().upper() in ("STATUS", "ESTATUS", "ESTADO")), None)
if _status_col_ui is not None:
    _n_fact = int(df_combined[_status_col_ui].astype(str).str.upper().str.contains("FACTURAD", na=False).sum())
    st.caption(f"📄 El archivo incluye la hoja **Facturadas** con {_n_fact:,} OT(s) facturadas — el Car Tracker las archiva automáticamente al importar.")

st.download_button(
    label="⬇️  Descargar Excels_Combinados.xlsx",
    data=excel_bytes,
    file_name="Excels_Combinados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary",
)
