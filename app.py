# app.py
# Proyecto independiente: WhatsApp + Dashboard de avance a objetivo
# Desarrollado para separar ventas objetivo y crosselling.

from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime
from io import BytesIO
from urllib.parse import quote

import pandas as pd
import streamlit as st

try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None


# =========================================================
# UTILIDADES
# =========================================================
def hoy_bolivia() -> date:
    if ZoneInfo is None:
        return date.today()
    return datetime.now(ZoneInfo("America/La_Paz")).date()


def sin_tildes(valor: object) -> str:
    if valor is None or pd.isna(valor):
        return ""
    texto = str(valor)
    return "".join(c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c))


def normalizar_columna(columna: object) -> str:
    texto = sin_tildes(columna).strip().lower()
    reemplazos = {
        " ": "_", "-": "_", "/": "_", ".": "_", "(": "", ")": "", "#": "", ":": "",
    }
    for a, b in reemplazos.items():
        texto = texto.replace(a, b)
    texto = re.sub(r"_+", "_", texto)
    return texto.strip("_")


def normalizar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalizar_columna(c) for c in df.columns]
    return df


def buscar_columna(df: pd.DataFrame, opciones: list[str]) -> str | None:
    columnas = set(df.columns)
    for opcion in opciones:
        op = normalizar_columna(opcion)
        if op in columnas:
            return op
    return None


def limpiar_codigo(valor: object) -> str:
    if valor is None or pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return re.sub(r"[^0-9A-Za-z-]", "", texto)


def limpiar_eh(valor: object) -> str:
    if valor is None or pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return re.sub(r"\D", "", texto)


def limpiar_texto(valor: object) -> str:
    if valor is None or pd.isna(valor):
        return ""
    return str(valor).strip()


def leer_archivo(archivo) -> pd.DataFrame:
    nombre = getattr(archivo, "name", "").lower()
    if nombre.endswith(".csv"):
        return pd.read_csv(archivo, sep=None, engine="python")
    return pd.read_excel(archivo, engine="openpyxl")


def formatear_fecha(valor: object) -> str:
    if valor is None or pd.isna(valor) or str(valor).strip() == "":
        return "S/F"
    fecha = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    if pd.isna(fecha):
        return str(valor).strip()
    return fecha.strftime("%d-%m-%Y")


def whatsapp_link(texto: str) -> str:
    return "https://wa.me/?text=" + quote(texto)


def excel_bytes(hojas: dict[str, pd.DataFrame]) -> BytesIO:
    salida = BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        for nombre, df in hojas.items():
            df.to_excel(writer, index=False, sheet_name=str(nombre)[:31])
        for hoja in writer.book.worksheets:
            hoja.freeze_panes = "A2"
            for col in hoja.columns:
                letra = col[0].column_letter
                ancho = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                hoja.column_dimensions[letra].width = min(max(ancho + 2, 10), 55)
    salida.seek(0)
    return salida


# =========================================================
# OBJETIVOS
# =========================================================
def detectar_columna_objetivo(df_obj: pd.DataFrame) -> str | None:
    # Primero intenta BU JUNIO / BU_JUNIO y meses comunes.
    candidatos_prioritarios = [
        "BU JUNIO", "BU_JUNIO", "OBJETIVO", "META", "BU", "JUNIO", "OBJ JUNIO", "OBJETIVO JUNIO",
    ]
    for c in candidatos_prioritarios:
        col = buscar_columna(df_obj, [c])
        if col:
            return col

    # Luego cualquier columna numérica que no sea POS_CODE/EH.
    for col in df_obj.columns:
        if col in {"pos_code", "eh", "codigo_eh", "vendedor_eh"}:
            continue
        serie_num = pd.to_numeric(df_obj[col], errors="coerce")
        if serie_num.notna().sum() > 0:
            return col
    return None


def procesar_objetivos(df_obj_original: pd.DataFrame, columna_objetivo_manual: str | None = None) -> pd.DataFrame:
    df = normalizar_dataframe(df_obj_original)
    col_eh = buscar_columna(df, ["POS_CODE", "EH", "VENDEDOR_EH", "CODIGO_EH", "COD_EH"])
    col_socio = buscar_columna(df, ["POS_OWNER", "SOCIO", "VENDEDOR_NOMBRE", "NOMBRE_SOCIO", "NOMBRE"])
    col_obj = normalizar_columna(columna_objetivo_manual) if columna_objetivo_manual else detectar_columna_objetivo(df)

    faltantes = []
    if not col_eh:
        faltantes.append("POS_CODE / EH")
    if not col_obj:
        faltantes.append("BU JUNIO / OBJETIVO")
    if faltantes:
        raise ValueError("No se encontraron columnas de objetivos: " + ", ".join(faltantes))

    salida = pd.DataFrame({
        "EH": df[col_eh].apply(limpiar_eh),
        "Socio objetivo": df[col_socio].apply(limpiar_texto) if col_socio else "",
        "Objetivo": pd.to_numeric(df[col_obj], errors="coerce").fillna(0).astype(int),
    })
    salida = salida[salida["EH"] != ""].drop_duplicates(subset=["EH"], keep="last")
    return salida


# =========================================================
# VENTAS
# =========================================================
def preparar_ventas(df_original: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str | None]]:
    df = normalizar_dataframe(df_original)

    col_codigo = buscar_columna(df, ["CLIENTE_NRO", "CODIGO_CLIENTE", "COD_CLIENTE", "CODIGO", "COD", "CLIENTE"])
    col_eh = buscar_columna(df, ["VENDEDOR_EH", "EH", "POS_CODE", "CODIGO_EH", "EHUMANO", "EH_PROMOTOR"])
    col_socio = buscar_columna(df, ["VENDEDOR_NOMBRE", "SOCIO", "POS_OWNER", "NOMBRE_SOCIO", "VENDEDOR", "EJECUTIVO"])
    col_tipo = buscar_columna(df, ["TIPO_VENTA", "TIPO VENTA", "TIPO", "TIPO_OPERACION", "OPERACION"])
    col_nodo = buscar_columna(df, ["NODO_NOMBRE", "NODO", "NODO_RED", "NODO ACTUAL"])
    col_fecha = buscar_columna(df, ["FECHA_INSTALACION", "FECHA_GENERACION_OT", "FECHA_REPORTE", "FECHA_VENTA", "FECHA", "FECHA_GENERACION"])
    col_cliente = buscar_columna(df, ["CLIENTE_NOMBRE", "NOMBRE_CLIENTE", "CLIENTE", "NOMBRE", "TITULAR"])
    col_tel1 = buscar_columna(df, ["CLIENTE_TELEFONO1", "TELEFONO1", "TEL1", "TELEFONO", "CELULAR"])
    col_tel2 = buscar_columna(df, ["CLIENTE_TELEFONO2", "TELEFONO2", "TEL2", "REFERENCIA", "REF"])

    columnas = {
        "codigo": col_codigo,
        "eh": col_eh,
        "socio": col_socio,
        "tipo_venta": col_tipo,
        "nodo": col_nodo,
        "fecha": col_fecha,
        "cliente": col_cliente,
        "telefono1": col_tel1,
        "telefono2": col_tel2,
    }

    faltantes = [k for k in ["codigo", "eh", "tipo_venta"] if columnas[k] is None]
    if faltantes:
        raise ValueError(
            "No se encontraron columnas obligatorias en el archivo de ventas: " + ", ".join(faltantes) +
            ". Columnas detectadas: " + ", ".join(df.columns[:80])
        )

    salida = pd.DataFrame({
        "Código cliente": df[col_codigo].apply(limpiar_codigo),
        "EH": df[col_eh].apply(limpiar_eh),
        "Socio": df[col_socio].apply(limpiar_texto) if col_socio else "SIN NOMBRE",
        "Tipo venta": df[col_tipo].apply(limpiar_texto).str.upper() if col_tipo else "",
        "Nodo": df[col_nodo].apply(limpiar_texto).str.upper() if col_nodo else "",
        "Fecha": df[col_fecha] if col_fecha else "",
        "Cliente": df[col_cliente].apply(limpiar_texto) if col_cliente else "",
        "Teléfono 1": df[col_tel1].apply(limpiar_texto) if col_tel1 else "",
        "Teléfono 2": df[col_tel2].apply(limpiar_texto) if col_tel2 else "",
    })
    salida = salida[(salida["Código cliente"] != "") & (salida["EH"] != "")].copy()
    salida["Es crosselling"] = salida["Tipo venta"].str.contains("CROSS|CROSSELL|CROSS_SELLING|CROSS SELLING", case=False, na=False)
    salida["Fecha texto"] = salida["Fecha"].apply(formatear_fecha)
    return salida, columnas


def calcular_resumen(ventas: pd.DataFrame, objetivos: pd.DataFrame | None) -> pd.DataFrame:
    if ventas.empty:
        return pd.DataFrame(columns=["EH", "Socio", "Ventas objetivo", "Crosselling", "Total ventas", "Objetivo", "Cumplimiento", "Faltan"])

    base = ventas.copy()
    resumen = base.groupby(["EH", "Socio"], as_index=False).agg(
        **{
            "Ventas objetivo": ("Es crosselling", lambda s: int((~s).sum())),
            "Crosselling": ("Es crosselling", lambda s: int(s.sum())),
            "Total ventas": ("Código cliente", "count"),
        }
    )

    if objetivos is not None and not objetivos.empty:
        resumen = resumen.merge(objetivos[["EH", "Objetivo"]], on="EH", how="left")
    else:
        resumen["Objetivo"] = 0

    resumen["Objetivo"] = pd.to_numeric(resumen["Objetivo"], errors="coerce").fillna(0).astype(int)
    resumen["Cumplimiento"] = resumen.apply(lambda r: (r["Ventas objetivo"] / r["Objetivo"] * 100) if r["Objetivo"] > 0 else 0, axis=1)
    resumen["Faltan"] = (resumen["Objetivo"] - resumen["Ventas objetivo"]).clip(lower=0).astype(int)
    resumen = resumen.sort_values(["Cumplimiento", "Ventas objetivo"], ascending=[False, False]).reset_index(drop=True)
    return resumen


# =========================================================
# MENSAJES WHATSAPP
# =========================================================
def mensaje_avance_socio(socio_row: pd.Series, ventas_socio: pd.DataFrame) -> str:
    eh = str(socio_row["EH"])
    socio = str(socio_row["Socio"])
    ventas_obj = int(socio_row["Ventas objetivo"])
    cross = int(socio_row["Crosselling"])
    total = int(socio_row["Total ventas"])
    objetivo = int(socio_row["Objetivo"])
    cumplimiento = float(socio_row["Cumplimiento"])
    faltan = int(socio_row["Faltan"])

    ventas_objetivo = ventas_socio[~ventas_socio["Es crosselling"]].copy()
    ventas_cross = ventas_socio[ventas_socio["Es crosselling"]].copy()

    lineas = [
        "📊 *AVANCE DE VENTAS*",
        "",
        f"👤 *{socio}*",
        f"EH: *{eh}*",
        "",
        f"✅ Ventas objetivo: *{ventas_obj}*",
        f"🔄 Crosselling: *{cross}*",
        f"📊 Total ventas: *{total}*",
        "",
        f"🎯 Objetivo: *{objetivo}*",
        f"📈 Cumplimiento: *{cumplimiento:.1f}%*",
        f"⏳ Faltan: *{faltan}*",
        "",
    ]

    if faltan == 0 and objetivo > 0:
        lineas.append("🏆 ¡Objetivo alcanzado! Sigamos asegurando instalaciones y calidad.")
    else:
        lineas.append("💪 Sigamos avanzando hacia el objetivo del mes.")

    lineas += ["", "✅ *CÓDIGOS QUE CUENTAN AL OBJETIVO:*", ""]
    if ventas_objetivo.empty:
        lineas.append("Sin códigos objetivo registrados.")
    else:
        for _, r in ventas_objetivo.sort_values("Fecha texto").iterrows():
            lineas.append(f"🔹 {r['Código cliente']} | {r['Cliente'] or 'S/N'} | {r['Fecha texto']}")

    if not ventas_cross.empty:
        lineas += ["", "🔄 *CROSSSELLING - NO SUMA AL OBJETIVO:*", ""]
        for _, r in ventas_cross.sort_values("Fecha texto").iterrows():
            lineas.append(f"🔸 {r['Código cliente']} | {r['Cliente'] or 'S/N'} | {r['Fecha texto']}")

    return "\n".join(lineas)


def mensaje_ranking_general(resumen: pd.DataFrame) -> str:
    total_obj = int(resumen["Ventas objetivo"].sum()) if not resumen.empty else 0
    total_cross = int(resumen["Crosselling"].sum()) if not resumen.empty else 0
    total_ventas = int(resumen["Total ventas"].sum()) if not resumen.empty else 0
    total_objetivo = int(resumen["Objetivo"].sum()) if not resumen.empty else 0
    cumplimiento_global = (total_obj / total_objetivo * 100) if total_objetivo > 0 else 0

    lineas = [
        "📊 *RANKING DE VENTAS*",
        "",
        f"✅ Ventas objetivo: *{total_obj}*",
        f"🔄 Crosselling: *{total_cross}*",
        f"📊 Total ventas: *{total_ventas}*",
        f"🎯 Objetivo global: *{total_objetivo}*",
        f"📈 Cumplimiento global: *{cumplimiento_global:.1f}%*",
        "",
        "🏆 *Ranking por cumplimiento:*",
    ]

    for i, (_, r) in enumerate(resumen.iterrows(), start=1):
        icono = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "🔹"
        lineas.append(
            f"{icono} {r['EH']} - {r['Socio']}: *{int(r['Ventas objetivo'])}/{int(r['Objetivo'])}* "
            f"({float(r['Cumplimiento']):.1f}%) | Cross: *{int(r['Crosselling'])}*"
        )

    lineas += ["", "💪 Equipo, enfoquémonos en ventas nuevas que cuentan al objetivo y mantengamos el ritmo del mes."]
    return "\n".join(lineas)


def mensaje_todos_los_socios(resumen: pd.DataFrame, ventas: pd.DataFrame) -> str:
    bloques = []
    for _, row in resumen.iterrows():
        det = ventas[ventas["EH"] == str(row["EH"])]
        bloques.append(mensaje_avance_socio(row, det))
    return "\n\n------------------------------\n\n".join(bloques)


def mensaje_solo_codigos_objetivo(row: pd.Series, ventas_socio: pd.DataFrame) -> str:
    ventas_obj = ventas_socio[~ventas_socio["Es crosselling"]].copy()
    lineas = [
        "✅ *CÓDIGOS QUE CUENTAN AL OBJETIVO*",
        "",
        f"👤 *{row['Socio']}*",
        f"EH: *{row['EH']}*",
        f"Total: *{len(ventas_obj)}*",
        "",
    ]
    if ventas_obj.empty:
        lineas.append("Sin códigos objetivo registrados.")
    else:
        for _, r in ventas_obj.sort_values("Fecha texto").iterrows():
            lineas.append(f"🔹 {r['Código cliente']} | {r['Cliente'] or 'S/N'} | {r['Fecha texto']}")
    return "\n".join(lineas)


# =========================================================
# STREAMLIT APP
# =========================================================
def cargar_datos() -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame]:
    st.sidebar.subheader("Archivos")
    archivo_ventas = st.sidebar.file_uploader("1) Subir base de ventas", type=["csv", "xlsx", "xls"], key="ventas")
    archivo_objetivos = st.sidebar.file_uploader("2) Subir Objetivo.xlsx", type=["xlsx", "xls", "csv"], key="objetivos")

    if archivo_ventas is None:
        st.info("Sube primero la base de ventas.")
        return pd.DataFrame(), None, pd.DataFrame()

    try:
        df_ventas_raw = leer_archivo(archivo_ventas)
        ventas, columnas_ventas = preparar_ventas(df_ventas_raw)
    except Exception as e:
        st.error(f"No se pudo leer la base de ventas: {e}")
        return pd.DataFrame(), None, pd.DataFrame()

    objetivos = None
    if archivo_objetivos is not None:
        try:
            df_obj_raw = leer_archivo(archivo_objetivos)
            df_obj_norm = normalizar_dataframe(df_obj_raw)
            col_default = detectar_columna_objetivo(df_obj_norm)
            columnas_objetivo = list(df_obj_norm.columns)
            col_sel = st.sidebar.selectbox(
                "Columna de objetivo",
                columnas_objetivo,
                index=columnas_objetivo.index(col_default) if col_default in columnas_objetivo else 0,
                key="col_objetivo",
            )
            objetivos = procesar_objetivos(df_obj_raw, col_sel)
        except Exception as e:
            st.sidebar.error(f"No se pudo leer Objetivo.xlsx: {e}")
            objetivos = None

    resumen = calcular_resumen(ventas, objetivos)
    return ventas, objetivos, resumen


def mostrar_dashboard(ventas: pd.DataFrame, resumen: pd.DataFrame) -> None:
    st.title("📊 Dashboard de avance")
    if ventas.empty:
        return

    total_obj = int(resumen["Ventas objetivo"].sum())
    total_cross = int(resumen["Crosselling"].sum())
    total_ventas = int(resumen["Total ventas"].sum())
    total_objetivo = int(resumen["Objetivo"].sum())
    cumplimiento = (total_obj / total_objetivo * 100) if total_objetivo > 0 else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Ventas objetivo", total_obj)
    c2.metric("Crosselling", total_cross)
    c3.metric("Total ventas", total_ventas)
    c4.metric("Objetivo", total_objetivo)
    c5.metric("Cumplimiento", f"{cumplimiento:.1f}%")

    st.subheader("Ranking por socio")
    st.dataframe(
        resumen[["EH", "Socio", "Ventas objetivo", "Crosselling", "Total ventas", "Objetivo", "Cumplimiento", "Faltan"]],
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("Detalle de ventas")
    st.dataframe(ventas, use_container_width=True, hide_index=True)

    excel = excel_bytes({"Ranking": resumen, "Detalle ventas": ventas})
    st.download_button(
        "⬇️ Descargar Excel",
        excel,
        file_name=f"ranking_ventas_{hoy_bolivia().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def mostrar_whatsapp(ventas: pd.DataFrame, resumen: pd.DataFrame) -> None:
    st.title("📲 WhatsApp - Avance de ventas")
    if ventas.empty:
        return

    tipo = st.selectbox(
        "Tipo de mensaje",
        ["Avance por socio", "Ranking general", "Todos los socios separados", "Solo códigos objetivo"],
        key="tipo_mensaje",
    )

    if tipo in {"Avance por socio", "Solo códigos objetivo"}:
        opciones = [f"{r['EH']} - {r['Socio']}" for _, r in resumen.iterrows()]
        seleccion = st.selectbox("Seleccionar socio", opciones, key="sel_socio")
        eh = seleccion.split(" - ")[0]
        row = resumen[resumen["EH"].astype(str) == eh].iloc[0]
        detalle = ventas[ventas["EH"].astype(str) == eh]

        st.subheader("Ventas del socio")
        st.dataframe(detalle, use_container_width=True, hide_index=True)

        if tipo == "Avance por socio":
            texto = mensaje_avance_socio(row, detalle)
        else:
            texto = mensaje_solo_codigos_objetivo(row, detalle)

    elif tipo == "Ranking general":
        st.subheader("Ranking")
        st.dataframe(resumen, use_container_width=True, hide_index=True)
        texto = mensaje_ranking_general(resumen)

    else:
        st.subheader("Todos los socios")
        st.dataframe(resumen, use_container_width=True, hide_index=True)
        texto = mensaje_todos_los_socios(resumen, ventas)

    st.subheader("Mensaje WhatsApp")
    st.text_area("Copiar mensaje", texto, height=430, key="mensaje_whatsapp")
    st.markdown(f"[📲 Enviar por WhatsApp]({whatsapp_link(texto)})")


def mostrar_inicio() -> None:
    st.markdown(
        """
        <div style="padding: 28px; border-radius: 18px; background: linear-gradient(135deg, #0033A0, #00AEEF); color: white;">
            <h1>📊 WhatsApp Objetivos Tigo</h1>
            <p style="font-size: 18px;">Ranking, avance a objetivo y mensajes WhatsApp por socio.</p>
            <hr style="border: 1px solid rgba(255,255,255,.25);">
            <p>👷 Desarrollado para seguimiento comercial</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.info("Sube la base de ventas y el archivo Objetivo.xlsx desde el panel lateral.")


def main() -> None:
    st.set_page_config(page_title="WhatsApp Objetivos Tigo", page_icon="📲", layout="wide")
    st.sidebar.title("📲 WhatsApp Objetivos")
    modulo = st.sidebar.radio("Módulo", ["Inicio", "Dashboard", "WhatsApp"], key="menu")

    ventas, objetivos, resumen = cargar_datos()

    if objetivos is None:
        st.sidebar.warning("Objetivos no cargados. El cumplimiento saldrá en 0 hasta subir Objetivo.xlsx.")

    if modulo == "Inicio":
        mostrar_inicio()
    elif modulo == "Dashboard":
        mostrar_dashboard(ventas, resumen)
    elif modulo == "WhatsApp":
        mostrar_whatsapp(ventas, resumen)


if __name__ == "__main__":
    main()
