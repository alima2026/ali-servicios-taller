from __future__ import annotations

import re
import unicodedata
from io import BytesIO
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Dashboard Taller Magna",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_EXCEL_PATH = Path(__file__).with_name("vehiculos_en_reparacion_magna.xlsx")
PREFERRED_SHEETS = ["SINIESTROS", "PARTICULAR Y GARANTIAS"]

CANONICAL_COLUMNS = [
    "FECHA",
    "CANAL",
    "DIAS EN TALLER",
    "COMPANIA",
    "NRO SINIESTRO",
    "PROVEEDOR",
    "CHASIS",
    "MATRICULA",
    "NOMBRE CLIENTE",
    "TELEFONO",
    "MARCA",
    "MODELO",
    "CODIGO",
    "REPUESTOS SOLICITADO",
    "MONTO PIEZA",
    "MONTO M OBRA",
    "STATUS DEL REPUESTO",
    "STATUS DEL VEHICULO",
    "FECHA ENTREGA PIEZA",
    "COMENTARIOS",
]

COLUMN_ALIASES = {
    "FECHA": "FECHA",
    "CANAL": "CANAL",
    "DIAS EN TALLER": "DIAS EN TALLER",
    "COMPANIA": "COMPANIA",
    "COMPANIAS": "COMPANIA",
    "N SINIESTRO": "NRO SINIESTRO",
    "NRO SINIESTRO": "NRO SINIESTRO",
    "NRO DE SINIESTRO": "NRO SINIESTRO",
    "NUMERO SINIESTRO": "NRO SINIESTRO",
    "PROVEEDOR": "PROVEEDOR",
    "CHASIS": "CHASIS",
    "MATRICULA": "MATRICULA",
    "NOMBRE CLIENTE": "NOMBRE CLIENTE",
    "TELEFONO": "TELEFONO",
    "MARCA": "MARCA",
    "MODELO": "MODELO",
    "CODIGO": "CODIGO",
    "REPUESTOS SOLICITADO": "REPUESTOS SOLICITADO",
    "REPUESTO SOLICITADO": "REPUESTOS SOLICITADO",
    "MONTO PIEZA": "MONTO PIEZA",
    "MONTO M OBRA": "MONTO M OBRA",
    "MONTO M OBRAS": "MONTO M OBRA",
    "STATUS DEL REPUESTO": "STATUS DEL REPUESTO",
    "STATUS REPUESTO": "STATUS DEL REPUESTO",
    "STATUS DEL VEHICULO": "STATUS DEL VEHICULO",
    "STATUS VEHICULO": "STATUS DEL VEHICULO",
    "FECHA ENTREGA PIEZA": "FECHA ENTREGA PIEZA",
    "COMENTARIOS": "COMENTARIOS",
}


def inject_css() -> None:
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 1rem;
            padding-bottom: 2rem;
        }
        .title-card {
            background: linear-gradient(135deg, rgba(255,255,255,0.98) 0%, rgba(240,249,255,0.98) 100%);
            border: 1px solid rgba(15,23,42,0.08);
            border-radius: 24px;
            padding: 1.3rem 1.5rem;
            margin-bottom: 1rem;
            box-shadow: 0 12px 28px rgba(15,23,42,0.08);
        }
        .title-card h1 {
            margin: 0;
            font-size: 2.7rem;
            line-height: 1;
            color: #0f172a;
            letter-spacing: -0.03em;
        }
        .title-card h1 span {
            color: #0f766e;
        }
        .title-card p {
            margin: 0.55rem 0 0 0;
            color: #475569;
            font-size: 1rem;
            font-weight: 500;
        }
        .hero-card {
            background: linear-gradient(135deg, #0f172a 0%, #0f766e 100%);
            color: white;
            border-radius: 22px;
            padding: 1.15rem 1.3rem;
            box-shadow: 0 10px 24px rgba(15,23,42,0.18);
            margin-bottom: 1rem;
        }
        .hero-title {
            font-size: 1.12rem;
            font-weight: 800;
            margin-bottom: 0.25rem;
        }
        .hero-text {
            font-size: 0.96rem;
            opacity: 0.94;
        }
        .metric-card {
            background: rgba(255,255,255,0.99);
            border: 1px solid rgba(15,23,42,0.06);
            border-radius: 18px;
            padding: 0.95rem 1rem;
            box-shadow: 0 4px 12px rgba(15,23,42,0.05);
            min-height: 118px;
        }
        .metric-label {
            font-size: 0.88rem;
            color: #64748b;
            margin-bottom: 0.25rem;
            font-weight: 700;
        }
        .metric-value {
            font-size: 1.7rem;
            line-height: 1.1;
            font-weight: 900;
            color: #0f172a;
        }
        .metric-help {
            margin-top: 0.35rem;
            color: #475569;
            font-size: 0.88rem;
        }
        .status-good {
            background: #ecfdf5;
            color: #166534;
            border-left: 6px solid #22c55e;
            padding: 0.85rem 1rem;
            border-radius: 14px;
            font-weight: 700;
            margin-bottom: 1rem;
        }
        .status-mid {
            background: #fffbeb;
            color: #92400e;
            border-left: 6px solid #f59e0b;
            padding: 0.85rem 1rem;
            border-radius: 14px;
            font-weight: 700;
            margin-bottom: 1rem;
        }
        [data-testid="metric-container"] {
            background: rgba(255,255,255,0.99);
            border: 1px solid rgba(15,23,42,0.06);
            padding: 1rem !important;
            border-radius: 18px !important;
            box-shadow: 0 4px 12px rgba(15,23,42,0.05);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def slug_text(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def first_non_empty(series: pd.Series) -> object:
    for value in series:
        if isinstance(value, pd.Timestamp) and not pd.isna(value):
            return value
        text = normalize_text(value)
        if text:
            return text
    return pd.NA


def count_non_empty(series: pd.Series) -> int:
    return int(sum(bool(normalize_text(value)) for value in series))


def unique_join(series: pd.Series) -> str:
    seen: set[str] = set()
    values: list[str] = []
    for value in series:
        text = normalize_text(value)
        if not text:
            continue
        key = slug_text(text)
        if key in seen:
            continue
        seen.add(key)
        values.append(text)
    return " | ".join(values)


def detect_header_row(raw_df: pd.DataFrame, max_rows: int = 15) -> int:
    for idx in range(min(max_rows, len(raw_df))):
        row_tokens = [slug_text(value) for value in raw_df.iloc[idx].tolist()]
        token_set = set(token for token in row_tokens if token)
        has_fecha = "FECHA" in token_set
        has_chasis = "CHASIS" in token_set
        has_repuesto = any(token.startswith("REPUESTOS SOLICITADO") for token in token_set)
        if has_fecha and has_chasis and has_repuesto:
            return idx
    return 0


def standardize_columns(columns: list[object]) -> list[str]:
    normalized: list[str] = []
    for idx, value in enumerate(columns):
        key = slug_text(value)
        normalized.append(COLUMN_ALIASES.get(key, normalize_text(value) or f"COL_{idx}"))
    return normalized


def clean_taller_dataframe(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    df = df.copy()
    for col in CANONICAL_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA

    ordered_cols = CANONICAL_COLUMNS + [col for col in df.columns if col not in CANONICAL_COLUMNS]
    df = df[ordered_cols]

    text_cols = [col for col in df.columns if col not in {"FECHA", "FECHA ENTREGA PIEZA"}]
    for col in text_cols:
        df[col] = df[col].apply(lambda value: normalize_text(value) if isinstance(value, str) else value)

    fillable_cols = [
        "FECHA",
        "CANAL",
        "DIAS EN TALLER",
        "COMPANIA",
        "NRO SINIESTRO",
        "PROVEEDOR",
        "CHASIS",
        "MATRICULA",
        "NOMBRE CLIENTE",
        "TELEFONO",
        "MARCA",
        "MODELO",
        "STATUS DEL REPUESTO",
        "STATUS DEL VEHICULO",
        "FECHA ENTREGA PIEZA",
        "COMENTARIOS",
    ]
    for col in fillable_cols + ["CODIGO", "REPUESTOS SOLICITADO", "MONTO PIEZA", "MONTO M OBRA"]:
        df[col] = df[col].apply(lambda value: pd.NA if normalize_text(value) == "" else value)

    df["CHASIS"] = df["CHASIS"].ffill()
    df["MATRICULA"] = df["MATRICULA"].ffill()
    df["VEHICULO_ID"] = df["CHASIS"].where(df["CHASIS"].notna(), df["MATRICULA"]).ffill()
    df = df[df["VEHICULO_ID"].notna()].copy()

    grouped = df.groupby("VEHICULO_ID", dropna=False, sort=False)
    for col in [
        "FECHA",
        "CANAL",
        "DIAS EN TALLER",
        "COMPANIA",
        "NRO SINIESTRO",
        "PROVEEDOR",
        "CHASIS",
        "MATRICULA",
        "NOMBRE CLIENTE",
        "TELEFONO",
        "MARCA",
        "MODELO",
        "STATUS DEL VEHICULO",
    ]:
        df[col] = grouped[col].transform(lambda series: series.ffill().bfill())

    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    df["FECHA ENTREGA PIEZA"] = pd.to_datetime(df["FECHA ENTREGA PIEZA"], errors="coerce")
    df["DIAS EN TALLER"] = pd.to_numeric(df["DIAS EN TALLER"], errors="coerce")

    for col in ["CANAL", "COMPANIA", "PROVEEDOR", "MARCA", "STATUS DEL REPUESTO", "STATUS DEL VEHICULO"]:
        df[col] = df[col].apply(normalize_text).str.upper()

    for col in ["MODELO", "CHASIS", "MATRICULA", "NOMBRE CLIENTE", "TELEFONO", "CODIGO", "REPUESTOS SOLICITADO", "COMENTARIOS", "NRO SINIESTRO"]:
        df[col] = df[col].apply(normalize_text)

    df["PROVEEDOR_NORMALIZADO"] = df["PROVEEDOR"].apply(slug_text)
    df["MAGNA_ADJUDICADO"] = df["PROVEEDOR_NORMALIZADO"] == "MAGNA"
    df["SOURCE_SHEET"] = sheet_name
    return df.reset_index(drop=True)


def read_sheet_smart(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    raw_df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    header_row = detect_header_row(raw_df)
    data = raw_df.iloc[header_row + 1 :].copy()
    data.columns = standardize_columns(raw_df.iloc[header_row].tolist())
    data = data.dropna(how="all").reset_index(drop=True)
    return clean_taller_dataframe(data, sheet_name)


@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    workbook = pd.ExcelFile(BytesIO(file_bytes))
    data: dict[str, pd.DataFrame] = {}
    for sheet_name in workbook.sheet_names:
        try:
            data[sheet_name] = read_sheet_smart(file_bytes, sheet_name)
        except Exception:
            continue
    return data


def build_vehicle_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            columns=[
                "VEHICULO_ID",
                "FECHA",
                "DIAS EN TALLER",
                "COMPANIA",
                "NRO SINIESTRO",
                "PROVEEDOR",
                "CHASIS",
                "MATRICULA",
                "NOMBRE CLIENTE",
                "MARCA",
                "MODELO",
                "STATUS DEL VEHICULO",
                "REPUESTOS SOLICITADOS",
                "REPUESTOS ENTREGADOS",
                "REPUESTOS PENDIENTES",
                "LISTA REPUESTOS",
                "MAGNA ADJUDICADO",
            ]
        )

    grouped = df.groupby("VEHICULO_ID", dropna=False, sort=False)
    summary = grouped.agg(
        FECHA=("FECHA", first_non_empty),
        DIAS_EN_TALLER=("DIAS EN TALLER", "max"),
        COMPANIA=("COMPANIA", first_non_empty),
        NRO_SINIESTRO=("NRO SINIESTRO", first_non_empty),
        PROVEEDOR=("PROVEEDOR", first_non_empty),
        CHASIS=("CHASIS", first_non_empty),
        MATRICULA=("MATRICULA", first_non_empty),
        NOMBRE_CLIENTE=("NOMBRE CLIENTE", first_non_empty),
        MARCA=("MARCA", first_non_empty),
        MODELO=("MODELO", first_non_empty),
        STATUS_DEL_VEHICULO=("STATUS DEL VEHICULO", first_non_empty),
        REPUESTOS_SOLICITADOS=("REPUESTOS SOLICITADO", count_non_empty),
        REPUESTOS_ENTREGADOS=("STATUS DEL REPUESTO", lambda series: int(sum(slug_text(value) == "ENTREGADO" for value in series))),
        LISTA_REPUESTOS=("REPUESTOS SOLICITADO", unique_join),
        MAGNA_ADJUDICADO=("MAGNA_ADJUDICADO", "max"),
    ).reset_index()

    summary["REPUESTOS_PENDIENTES"] = summary["REPUESTOS_SOLICITADOS"] - summary["REPUESTOS_ENTREGADOS"]
    return summary


def provider_summary(df: pd.DataFrame, vehicle_summary: pd.DataFrame) -> pd.DataFrame:
    if vehicle_summary.empty:
        return pd.DataFrame(columns=["PROVEEDOR", "VEHICULOS", "REPUESTOS"])
    data = (
        vehicle_summary.groupby("PROVEEDOR", dropna=False)
        .agg(
            VEHICULOS=("VEHICULO_ID", "count"),
            REPUESTOS=("REPUESTOS_SOLICITADOS", "sum"),
        )
        .reset_index()
        .sort_values(["REPUESTOS", "VEHICULOS"], ascending=False)
    )
    data["PROVEEDOR"] = data["PROVEEDOR"].replace("", "SIN PROVEEDOR")
    return data


def status_summary(vehicle_summary: pd.DataFrame) -> pd.DataFrame:
    if vehicle_summary.empty:
        return pd.DataFrame(columns=["STATUS DEL VEHICULO", "VEHICULOS"])
    data = (
        vehicle_summary.groupby("STATUS_DEL_VEHICULO", dropna=False)
        .agg(VEHICULOS=("VEHICULO_ID", "count"))
        .reset_index()
        .sort_values("VEHICULOS", ascending=False)
    )
    data["STATUS DEL VEHICULO"] = data["STATUS_DEL_VEHICULO"].replace("", "SIN ESTADO")
    return data[["STATUS DEL VEHICULO", "VEHICULOS"]]


def brand_or_model_summary(vehicle_summary: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    if vehicle_summary.empty:
        return pd.DataFrame(columns=["CATEGORIA", "VEHICULOS"]), "Marca"

    if vehicle_summary["MARCA"].nunique(dropna=True) > 1:
        data = (
            vehicle_summary.groupby("MARCA", dropna=False)
            .agg(VEHICULOS=("VEHICULO_ID", "count"))
            .reset_index()
            .sort_values("VEHICULOS", ascending=False)
        )
        data["CATEGORIA"] = data["MARCA"].replace("", "SIN MARCA")
        return data[["CATEGORIA", "VEHICULOS"]], "Marca"

    data = (
        vehicle_summary.groupby("MODELO", dropna=False)
        .agg(VEHICULOS=("VEHICULO_ID", "count"))
        .reset_index()
        .sort_values("VEHICULOS", ascending=False)
    )
    data["CATEGORIA"] = data["MODELO"].replace("", "SIN MODELO")
    return data[["CATEGORIA", "VEHICULOS"]], "Modelo"


def top_vehicles_summary(vehicle_summary: pd.DataFrame) -> pd.DataFrame:
    if vehicle_summary.empty:
        return pd.DataFrame(columns=["VEHICULO", "REPUESTOS"])
    data = vehicle_summary.copy()
    data["VEHICULO"] = data.apply(
        lambda row: f"{row['MATRICULA'] or row['CHASIS']} | {row['MODELO'] or 'Sin modelo'}",
        axis=1,
    )
    return data.sort_values("REPUESTOS_SOLICITADOS", ascending=False)[["VEHICULO", "REPUESTOS_SOLICITADOS"]].rename(
        columns={"REPUESTOS_SOLICITADOS": "REPUESTOS"}
    )


def format_date(value: object) -> str:
    if isinstance(value, pd.Timestamp) and not pd.isna(value):
        return value.strftime("%Y-%m-%d")
    return normalize_text(value)


def metric_card(title: str, value: object, help_text: str) -> None:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{title}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-help">{help_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def horizontal_bar(
    df: pd.DataFrame,
    category_col: str,
    value_col: str,
    title: str,
    color: str,
) -> None:
    if df.empty:
        st.info("No hay datos para este grafico con los filtros actuales.")
        return

    chart = (
        alt.Chart(df)
        .mark_bar(cornerRadiusTopRight=7, cornerRadiusBottomRight=7)
        .encode(
            x=alt.X(f"{value_col}:Q", title=""),
            y=alt.Y(f"{category_col}:N", sort="-x", title=""),
            color=alt.value(color),
            tooltip=[category_col, value_col],
        )
        .properties(height=max(220, 36 * len(df)), title=title)
    )
    st.altair_chart(chart, use_container_width=True)


def dataframe_to_excel_bytes(
    vehicle_summary_df: pd.DataFrame,
    repuestos_df: pd.DataFrame,
    provider_df: pd.DataFrame,
    status_df: pd.DataFrame,
) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        vehicle_summary_df.to_excel(writer, sheet_name="Vehiculos", index=False)
        repuestos_df.to_excel(writer, sheet_name="Repuestos", index=False)
        provider_df.to_excel(writer, sheet_name="Proveedores", index=False)
        status_df.to_excel(writer, sheet_name="Estados", index=False)
    return buffer.getvalue()


def get_input_file() -> tuple[bytes | None, str]:
    uploaded_file = st.sidebar.file_uploader("Subir Excel del taller", type=["xlsx", "xls"])
    if uploaded_file is not None:
        return uploaded_file.getvalue(), uploaded_file.name
    if DEFAULT_EXCEL_PATH.exists():
        return DEFAULT_EXCEL_PATH.read_bytes(), DEFAULT_EXCEL_PATH.name
    return None, ""


def build_search_mask(vehicle_summary: pd.DataFrame, search_text: str) -> pd.Series:
    if not search_text:
        return pd.Series(True, index=vehicle_summary.index)
    query = search_text.strip().lower()
    searchable_cols = ["CHASIS", "MATRICULA", "NOMBRE_CLIENTE", "MODELO", "LISTA_REPUESTOS", "NRO_SINIESTRO"]
    combined = (
        vehicle_summary[searchable_cols]
        .fillna("")
        .astype(str)
        .agg(" ".join, axis=1)
        .str.lower()
    )
    return combined.str.contains(query, na=False)


inject_css()

st.markdown(
    """
    <div class="title-card">
        <h1>Dashboard <span>Taller Magna</span></h1>
        <p>Conteo de vehiculos, repuestos por unidad, marcas y adjudicacion al proveedor MAGNA.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

file_bytes, active_file_name = get_input_file()

if not file_bytes:
    st.error("No se encontro un archivo Excel para analizar.")
    st.stop()

workbook_data = load_workbook(file_bytes)
valid_sheets = [sheet for sheet in PREFERRED_SHEETS if sheet in workbook_data and not workbook_data[sheet].empty]
valid_sheets += [sheet for sheet in workbook_data if sheet not in valid_sheets and not workbook_data[sheet].empty]

if not valid_sheets:
    st.error("No se detectaron hojas con estructura valida de taller en el archivo.")
    st.stop()

default_sheet = "SINIESTROS" if "SINIESTROS" in valid_sheets else valid_sheets[0]

st.sidebar.subheader("Fuente")
selected_sheet = st.sidebar.selectbox(
    "Hoja a analizar",
    valid_sheets,
    index=valid_sheets.index(default_sheet),
    help="La hoja SINIESTROS suele ser la mejor referencia para revisar adjudicaciones a MAGNA.",
)

raw_df = workbook_data[selected_sheet].copy()
vehicle_summary_df = build_vehicle_summary(raw_df)

st.sidebar.markdown("---")
st.sidebar.subheader("Filtros")

brand_options = sorted([value for value in vehicle_summary_df["MARCA"].dropna().unique().tolist() if value])
status_options = sorted([value for value in vehicle_summary_df["STATUS_DEL_VEHICULO"].dropna().unique().tolist() if value])
provider_options = sorted([value for value in vehicle_summary_df["PROVEEDOR"].dropna().unique().tolist() if value])

selected_brands = st.sidebar.multiselect("Marca", brand_options, default=brand_options)
selected_status = st.sidebar.multiselect("Estado del vehiculo", status_options, default=status_options)
selected_providers = st.sidebar.multiselect("Proveedor", provider_options, default=provider_options)
only_magna = st.sidebar.checkbox("Solo adjudicados a MAGNA", value=False)
search_text = st.sidebar.text_input("Buscar cliente, chasis, matricula o repuesto")

vehicle_mask = pd.Series(True, index=vehicle_summary_df.index)
if brand_options:
    vehicle_mask &= vehicle_summary_df["MARCA"].isin(selected_brands)
if status_options:
    vehicle_mask &= vehicle_summary_df["STATUS_DEL_VEHICULO"].isin(selected_status)
if provider_options:
    vehicle_mask &= vehicle_summary_df["PROVEEDOR"].isin(selected_providers)
if only_magna:
    vehicle_mask &= vehicle_summary_df["MAGNA_ADJUDICADO"]
vehicle_mask &= build_search_mask(vehicle_summary_df, search_text)

filtered_vehicle_summary = vehicle_summary_df.loc[vehicle_mask].copy()
filtered_vehicle_ids = filtered_vehicle_summary["VEHICULO_ID"].tolist()
filtered_df = raw_df[raw_df["VEHICULO_ID"].isin(filtered_vehicle_ids)].copy()

if selected_sheet != "SINIESTROS":
    st.info(
        "Estas viendo una hoja distinta a SINIESTROS. Para revisar adjudicaciones a MAGNA, la referencia mas util suele ser SINIESTROS."
    )

total_vehicles = len(filtered_vehicle_summary)
vehicles_in_shop = int((filtered_vehicle_summary["STATUS_DEL_VEHICULO"] == "EN TALLER").sum())
total_parts = int(filtered_vehicle_summary["REPUESTOS_SOLICITADOS"].sum())
magna_vehicles = int(filtered_vehicle_summary["MAGNA_ADJUDICADO"].sum())
magna_parts = int(filtered_vehicle_summary.loc[filtered_vehicle_summary["MAGNA_ADJUDICADO"], "REPUESTOS_SOLICITADOS"].sum())
unique_brands = max(filtered_vehicle_summary["MARCA"].replace("", pd.NA).nunique(dropna=True), 0)

hero_message = (
    f"Archivo activo: <strong>{active_file_name}</strong> | Hoja: <strong>{selected_sheet}</strong> | "
    f"Vehiculos filtrados: <strong>{total_vehicles}</strong> | Repuestos filtrados: <strong>{total_parts}</strong>"
)
st.markdown(
    f"""
    <div class="hero-card">
        <div class="hero-title">Resumen operativo del taller</div>
        <div class="hero-text">{hero_message}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if total_parts > 0 and magna_parts > 0:
    part_share = magna_parts / total_parts * 100
    vehicle_share = (magna_vehicles / total_vehicles * 100) if total_vehicles else 0
    st.markdown(
        f"""
        <div class="status-good">
            MAGNA tiene adjudicados {magna_parts} repuestos en {magna_vehicles} vehiculos
            ({part_share:.1f}% de las piezas y {vehicle_share:.1f}% de los vehiculos filtrados).
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        """
        <div class="status-mid">
            Con los filtros actuales no aparecen repuestos adjudicados a MAGNA.
        </div>
        """,
        unsafe_allow_html=True,
    )

metric_cols = st.columns(5)
with metric_cols[0]:
    metric_card("Vehiculos unicos", total_vehicles, "Vehiculos distintos detectados en la hoja elegida.")
with metric_cols[1]:
    metric_card("Estado EN TALLER", vehicles_in_shop, "Vehiculos cuyo estado actual figura como EN TALLER.")
with metric_cols[2]:
    metric_card("Repuestos solicitados", total_parts, "Cantidad total de piezas asociadas a los vehiculos filtrados.")
with metric_cols[3]:
    metric_card("Vehiculos MAGNA", magna_vehicles, "Vehiculos cuyo proveedor adjudicado es MAGNA.")
with metric_cols[4]:
    metric_card("Marcas detectadas", unique_brands, "Cantidad de marcas presentes luego de aplicar los filtros.")

provider_df = provider_summary(filtered_df, filtered_vehicle_summary)
status_df = status_summary(filtered_vehicle_summary)
brand_df, brand_label = brand_or_model_summary(filtered_vehicle_summary)
top_vehicle_df = top_vehicles_summary(filtered_vehicle_summary).head(10)

chart_col_1, chart_col_2 = st.columns(2)
with chart_col_1:
    horizontal_bar(provider_df, "PROVEEDOR", "REPUESTOS", "Repuestos por proveedor", "#0f766e")
with chart_col_2:
    horizontal_bar(status_df, "STATUS DEL VEHICULO", "VEHICULOS", "Vehiculos por estado", "#1d4ed8")

chart_col_3, chart_col_4 = st.columns(2)
with chart_col_3:
    horizontal_bar(brand_df, "CATEGORIA", "VEHICULOS", f"Vehiculos por {brand_label.lower()}", "#0f172a")
with chart_col_4:
    horizontal_bar(top_vehicle_df, "VEHICULO", "REPUESTOS", "Top vehiculos por cantidad de repuestos", "#f59e0b")

vehicle_display = filtered_vehicle_summary.copy()
vehicle_display["FECHA"] = vehicle_display["FECHA"].apply(format_date)
vehicle_display["MAGNA ADJUDICADO"] = vehicle_display["MAGNA_ADJUDICADO"].map({True: "SI", False: "NO"})
vehicle_display = vehicle_display.rename(
    columns={
        "DIAS_EN_TALLER": "DIAS EN TALLER",
        "NRO_SINIESTRO": "NRO SINIESTRO",
        "NOMBRE_CLIENTE": "NOMBRE CLIENTE",
        "STATUS_DEL_VEHICULO": "STATUS DEL VEHICULO",
        "REPUESTOS_SOLICITADOS": "REPUESTOS SOLICITADOS",
        "REPUESTOS_ENTREGADOS": "REPUESTOS ENTREGADOS",
        "REPUESTOS_PENDIENTES": "REPUESTOS PENDIENTES",
        "LISTA_REPUESTOS": "LISTA REPUESTOS",
    }
)

repuestos_display = filtered_df.copy()
repuestos_display["FECHA"] = repuestos_display["FECHA"].apply(format_date)
repuestos_display["FECHA ENTREGA PIEZA"] = repuestos_display["FECHA ENTREGA PIEZA"].apply(format_date)
repuestos_display["MAGNA ADJUDICADO"] = repuestos_display["MAGNA_ADJUDICADO"].map({True: "SI", False: "NO"})

summary_export = vehicle_display[
    [
        "FECHA",
        "COMPANIA",
        "NRO SINIESTRO",
        "PROVEEDOR",
        "MAGNA ADJUDICADO",
        "CHASIS",
        "MATRICULA",
        "NOMBRE CLIENTE",
        "MARCA",
        "MODELO",
        "STATUS DEL VEHICULO",
        "DIAS EN TALLER",
        "REPUESTOS SOLICITADOS",
        "REPUESTOS ENTREGADOS",
        "REPUESTOS PENDIENTES",
        "LISTA REPUESTOS",
    ]
].copy()

repuestos_export = repuestos_display[
    [
        "FECHA",
        "COMPANIA",
        "NRO SINIESTRO",
        "PROVEEDOR",
        "MAGNA ADJUDICADO",
        "CHASIS",
        "MATRICULA",
        "NOMBRE CLIENTE",
        "MARCA",
        "MODELO",
        "CODIGO",
        "REPUESTOS SOLICITADO",
        "STATUS DEL REPUESTO",
        "STATUS DEL VEHICULO",
        "FECHA ENTREGA PIEZA",
        "COMENTARIOS",
    ]
].copy()

download_bytes = dataframe_to_excel_bytes(summary_export, repuestos_export, provider_df, status_df)

st.subheader("Detalle por vehiculo")
st.dataframe(
    summary_export.sort_values(["PROVEEDOR", "REPUESTOS SOLICITADOS"], ascending=[True, False]),
    use_container_width=True,
    hide_index=True,
)

st.subheader("Detalle de repuestos")
st.dataframe(
    repuestos_export.sort_values(["PROVEEDOR", "CHASIS", "REPUESTOS SOLICITADO"], ascending=[True, True, True]),
    use_container_width=True,
    hide_index=True,
)

st.download_button(
    "Descargar resumen en Excel",
    data=download_bytes,
    file_name=f"resumen_taller_magna_{selected_sheet.lower().replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
