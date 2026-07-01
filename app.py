# app.py
# Proyecto independiente: WhatsApp + Dashboard de avance a objetivo
# V4: ranking motivacional con podio, colores, semáforo, retos y WhatsApp por socio.

from __future__ import annotations

import hashlib
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


def fecha_orden(valor: object) -> pd.Timestamp:
    fecha = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    if pd.isna(fecha):
        return pd.Timestamp.max
    return fecha


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


def key_mensaje(prefix: str, texto: str) -> str:
    digest = hashlib.md5(texto.encode("utf-8", errors="ignore")).hexdigest()[:10]
    safe = re.sub(r"[^A-Za-z0-9_]+", "_", prefix)[:40]
    return f"msg_{safe}_{digest}"


# =========================================================
# ESTILO VISUAL / MOTIVACIÓN
# =========================================================
def aplicar_estilos() -> None:
    st.markdown(
        """
        <style>
        .main .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
        .hero-box {
            padding: 28px;
            border-radius: 24px;
            background: linear-gradient(135deg, #0033A0 0%, #006BEA 45%, #00AEEF 100%);
            color: white;
            box-shadow: 0 12px 30px rgba(0, 51, 160, .22);
            margin-bottom: 18px;
        }
        .hero-box h1 {font-size: 2.25rem; margin: 0 0 8px 0;}
        .hero-box p {font-size: 1.05rem; margin: 0; opacity: .96;}
        .metric-card {
            padding: 18px 18px;
            border-radius: 20px;
            color: white;
            min-height: 118px;
            box-shadow: 0 8px 22px rgba(15, 23, 42, .13);
            border: 1px solid rgba(255, 255, 255, .26);
        }
        .metric-card .label {font-size: .88rem; opacity: .92; margin-bottom: 8px;}
        .metric-card .value {font-size: 2rem; font-weight: 800; line-height: 1.05;}
        .metric-card .note {font-size: .82rem; opacity: .94; margin-top: 8px;}
        .podio-card {
            padding: 20px;
            border-radius: 24px;
            color: white;
            box-shadow: 0 10px 28px rgba(15, 23, 42, .16);
            min-height: 210px;
            border: 1px solid rgba(255, 255, 255, .28);
        }
        .podio-card .medal {font-size: 2.2rem; margin-bottom: 8px;}
        .podio-card .name {font-size: 1.15rem; font-weight: 800; margin-bottom: 4px;}
        .podio-card .eh {font-size: .82rem; opacity: .9; margin-bottom: 12px;}
        .podio-card .big {font-size: 1.65rem; font-weight: 900;}
        .podio-card .small {font-size: .88rem; opacity: .95; margin-top: 8px;}
        .badge-box {
            padding: 16px;
            border-radius: 18px;
            background: #F8FAFC;
            border: 1px solid #E2E8F0;
            min-height: 118px;
        }
        .badge-box .title {font-weight: 800; font-size: 1rem; color: #0F172A; margin-bottom: 6px;}
        .badge-box .text {font-size: .92rem; color: #334155;}
        .challenge-box {
            padding: 18px;
            border-radius: 22px;
            background: linear-gradient(135deg, #FFF7ED, #FFEDD5);
            border: 1px solid #FDBA74;
            color: #7C2D12;
            margin: 12px 0 18px 0;
        }
        .challenge-box b {color: #9A3412;}
        .stProgress > div > div > div > div {border-radius: 999px;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def tarjeta_metric(label: str, valor: str | int | float, nota: str, color1: str, color2: str) -> None:
    st.markdown(
        f"""
        <div class="metric-card" style="background: linear-gradient(135deg, {color1}, {color2});">
            <div class="label">{label}</div>
            <div class="value">{valor}</div>
            <div class="note">{nota}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def estado_cumplimiento(cumplimiento: float, ventas_obj: int, objetivo: int) -> tuple[str, str, str, str, str]:
    if objetivo <= 0:
        return "⚪", "Sin objetivo", "#64748B", "Objetivo no cargado", "Cargar objetivo para medir avance real."
    if ventas_obj >= objetivo:
        return "🏆", "Objetivo cumplido", "#16A34A", "Meta cumplida", "Mantener ritmo y cuidar calidad de instalación."
    if cumplimiento >= 80:
        return "🔥", "Muy cerca", "#2563EB", "A punto de cumplir", "Cerrar las ventas faltantes hoy o mañana."
    if cumplimiento >= 60:
        return "💪", "En carrera", "#EAB308", "Buen avance", "Necesita empuje diario para llegar a tiempo."
    if ventas_obj > 0:
        return "⚡", "Necesita impulso", "#F97316", "Hay oportunidad", "Activar seguimiento de prospectos y cierres rápidos."
    return "🚨", "Sin avance", "#DC2626", "Sin ventas objetivo", "Prioridad: conseguir la primera venta objetivo."


def mensaje_motivacional_por_estado(cumplimiento: float, ventas_obj: int, objetivo: int, faltan: int) -> str:
    if objetivo <= 0:
        return "Carguemos su objetivo para medir el avance correctamente."
    if ventas_obj >= objetivo:
        extra = ventas_obj - objetivo
        return f"🏆 Objetivo cumplido. Excelente trabajo; ya está {extra} venta(s) arriba de la meta."
    if faltan <= 3:
        return f"🔥 Está a solo {faltan} venta(s) de cumplir. Hoy puede cerrar el objetivo."
    if cumplimiento >= 80:
        return f"💪 Muy buen avance. Faltan {faltan} venta(s), hay que empujar el cierre."
    if cumplimiento >= 60:
        return f"⚡ Va en carrera. Faltan {faltan} venta(s), necesita constancia diaria."
    if ventas_obj > 0:
        return f"📌 Ya empezó, ahora toca acelerar. Faltan {faltan} venta(s) para cumplir."
    return "🚨 Todavía no registra ventas objetivo. Hoy la meta mínima es lograr la primera venta."


def reto_sugerido(cumplimiento: float, ventas_obj: int, objetivo: int, faltan: int) -> str:
    if objetivo <= 0:
        return "Revisar objetivo cargado."
    if ventas_obj >= objetivo:
        return "Reto: mantener calidad y sumar ventas extra sin descuidar instalaciones."
    if faltan <= 3:
        return f"Reto: cerrar {faltan} venta(s) para entrar al grupo de objetivo cumplido."
    if cumplimiento >= 80:
        return "Reto: mínimo 1 venta objetivo diaria hasta cumplir."
    if cumplimiento >= 60:
        return "Reto: recuperar 2 ventas objetivo esta semana."
    if ventas_obj > 0:
        return "Reto: revisar cartera, llamar pendientes y buscar 2 cierres rápidos."
    return "Reto: primera venta objetivo del día antes del mediodía."


def agregar_columnas_visuales(resumen: pd.DataFrame) -> pd.DataFrame:
    if resumen.empty:
        return resumen.copy()
    salida = resumen.copy().reset_index(drop=True)
    salida.insert(0, "Puesto", range(1, len(salida) + 1))
    estados = salida.apply(lambda r: estado_cumplimiento(float(r["Cumplimiento"]), int(r["Ventas objetivo"]), int(r["Objetivo"])), axis=1)
    salida["Estado"] = [f"{e[0]} {e[1]}" for e in estados]
    salida["Motivación"] = salida.apply(
        lambda r: mensaje_motivacional_por_estado(float(r["Cumplimiento"]), int(r["Ventas objetivo"]), int(r["Objetivo"]), int(r["Faltan"])),
        axis=1,
    )
    salida["Reto sugerido"] = salida.apply(
        lambda r: reto_sugerido(float(r["Cumplimiento"]), int(r["Ventas objetivo"]), int(r["Objetivo"]), int(r["Faltan"])),
        axis=1,
    )
    salida["Avance"] = salida.apply(lambda r: f"{int(r['Ventas objetivo'])}/{int(r['Objetivo'])}", axis=1)
    salida["Cumplimiento"] = salida["Cumplimiento"].round(1)
    return salida


def color_fila_ranking(row: pd.Series) -> list[str]:
    cumplimiento = float(row.get("Cumplimiento", 0) or 0)
    ventas_obj = int(row.get("Ventas objetivo", 0) or 0)
    objetivo = int(row.get("Objetivo", 0) or 0)
    _, _, color, _, _ = estado_cumplimiento(cumplimiento, ventas_obj, objetivo)
    fondo = {
        "#16A34A": "background-color: #DCFCE7; color: #14532D; font-weight: 700;",
        "#2563EB": "background-color: #DBEAFE; color: #1E3A8A; font-weight: 700;",
        "#EAB308": "background-color: #FEF9C3; color: #713F12; font-weight: 700;",
        "#F97316": "background-color: #FFEDD5; color: #7C2D12; font-weight: 700;",
        "#DC2626": "background-color: #FEE2E2; color: #7F1D1D; font-weight: 700;",
        "#64748B": "background-color: #F1F5F9; color: #334155; font-weight: 700;",
    }.get(color, "")
    return [fondo for _ in row]


def render_podio(resumen: pd.DataFrame) -> None:
    if resumen.empty:
        return
    st.subheader("🏆 Podio motivacional")
    top = resumen.head(3).copy().reset_index(drop=True)
    cols = st.columns(3)
    medallas = ["🥇", "🥈", "🥉"]
    colores = [
        ("#F59E0B", "#F97316"),
        ("#64748B", "#94A3B8"),
        ("#B45309", "#D97706"),
    ]
    for i, (_, r) in enumerate(top.iterrows()):
        emoji, estado, _, _, _ = estado_cumplimiento(float(r["Cumplimiento"]), int(r["Ventas objetivo"]), int(r["Objetivo"]))
        mensaje = mensaje_motivacional_por_estado(float(r["Cumplimiento"]), int(r["Ventas objetivo"]), int(r["Objetivo"]), int(r["Faltan"]))
        c1, c2 = colores[i]
        cols[i].markdown(
            f"""
            <div class="podio-card" style="background: linear-gradient(135deg, {c1}, {c2});">
                <div class="medal">{medallas[i]}</div>
                <div class="name">{r['Socio']}</div>
                <div class="eh">EH {r['EH']}</div>
                <div class="big">{int(r['Ventas objetivo'])}/{int(r['Objetivo'])}</div>
                <div class="small">{emoji} {estado} · {float(r['Cumplimiento']):.1f}%</div>
                <div class="small">🔄 Crosselling: {int(r['Crosselling'])}</div>
                <div class="small">{mensaje}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_insignias(resumen: pd.DataFrame) -> None:
    if resumen.empty:
        return
    st.subheader("🎖️ Reconocimientos para motivar")
    con_obj = resumen[resumen["Objetivo"] > 0].copy()
    cols = st.columns(4)

    if not con_obj.empty:
        lider = con_obj.sort_values(["Cumplimiento", "Ventas objetivo"], ascending=[False, False]).iloc[0]
        cerca = con_obj[(con_obj["Faltan"] > 0) & (con_obj["Faltan"] <= 3)].sort_values("Faltan").head(1)
        impulso = con_obj[con_obj["Ventas objetivo"] > 0].sort_values(["Faltan", "Cumplimiento"], ascending=[True, False]).head(1)
        cross = resumen[resumen["Crosselling"] > 0].sort_values("Crosselling", ascending=False).head(1)

        cols[0].markdown(f"""<div class="badge-box"><div class="title">🏆 Líder del ranking</div><div class="text">{lider['Socio']} · {float(lider['Cumplimiento']):.1f}%</div></div>""", unsafe_allow_html=True)
        if not cerca.empty:
            r = cerca.iloc[0]
            cols[1].markdown(f"""<div class="badge-box"><div class="title">🔥 Más cerca de cumplir</div><div class="text">{r['Socio']} · faltan {int(r['Faltan'])}</div></div>""", unsafe_allow_html=True)
        else:
            cols[1].markdown("""<div class="badge-box"><div class="title">🔥 Más cerca de cumplir</div><div class="text">Sin socios a 3 ventas o menos.</div></div>""", unsafe_allow_html=True)
        if not impulso.empty:
            r = impulso.iloc[0]
            cols[2].markdown(f"""<div class="badge-box"><div class="title">🚀 Socio en impulso</div><div class="text">{r['Socio']} · {int(r['Ventas objetivo'])} ventas objetivo</div></div>""", unsafe_allow_html=True)
        else:
            cols[2].markdown("""<div class="badge-box"><div class="title">🚀 Socio en impulso</div><div class="text">Pendiente de activar.</div></div>""", unsafe_allow_html=True)
        if not cross.empty:
            r = cross.iloc[0]
            cols[3].markdown(f"""<div class="badge-box"><div class="title">🔄 Mayor crosselling</div><div class="text">{r['Socio']} · {int(r['Crosselling'])} crosselling</div></div>""", unsafe_allow_html=True)
        else:
            cols[3].markdown("""<div class="badge-box"><div class="title">🔄 Mayor crosselling</div><div class="text">Sin crosselling registrado.</div></div>""", unsafe_allow_html=True)


def render_reto_equipo(resumen: pd.DataFrame) -> None:
    if resumen.empty:
        return
    total_obj = int(resumen["Ventas objetivo"].sum())
    total_meta = int(resumen["Objetivo"].sum())
    faltan = max(total_meta - total_obj, 0)
    cumplimiento = (total_obj / total_meta * 100) if total_meta > 0 else 0
    if total_meta <= 0:
        texto = "Carga el archivo de objetivos para activar el reto del equipo."
    elif faltan == 0:
        texto = "🏆 El equipo ya cumplió el objetivo global. Reto: sostener calidad, instalaciones y ventas extra."
    elif cumplimiento >= 80:
        texto = f"🔥 Reto del equipo: faltan {faltan} ventas objetivo. Cada socio debe empujar mínimo 1 cierre."
    else:
        texto = f"💪 Reto del equipo: faltan {faltan} ventas objetivo. Enfoque diario en ventas nuevas, no en crosselling."
    st.markdown(f"""<div class="challenge-box"><b>🎯 Reto del día:</b> {texto}</div>""", unsafe_allow_html=True)


def render_barra_equipo(total_obj: int, total_objetivo: int) -> None:
    if total_objetivo <= 0:
        st.progress(0, text="Objetivo global no cargado")
        return
    avance = min(max(total_obj / total_objetivo, 0), 1)
    st.progress(avance, text=f"Avance global: {total_obj}/{total_objetivo} ventas objetivo ({avance * 100:.1f}%)")


# =========================================================
# OBJETIVOS
# =========================================================
def detectar_columna_objetivo(df_obj: pd.DataFrame) -> str | None:
    candidatos_prioritarios = [
        "BU JUNIO", "BU_JUNIO", "OBJETIVO", "META", "BU", "JUNIO", "OBJ JUNIO", "OBJETIVO JUNIO",
    ]
    for c in candidatos_prioritarios:
        col = buscar_columna(df_obj, [c])
        if col:
            return col

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
# CROSSSELLING MANUAL / ADICIONAL
# =========================================================
def extraer_codigos_de_texto(texto: str) -> set[str]:
    codigos: set[str] = set()
    for token in re.split(r"[\s,;|]+", str(texto)):
        codigo = limpiar_codigo(token)
        if codigo:
            codigos.add(codigo)
    return codigos


def extraer_codigos_de_archivo(archivo) -> set[str]:
    if archivo is None:
        return set()
    try:
        df = leer_archivo(archivo)
    except Exception:
        return set()
    if df.empty:
        return set()

    df_norm = normalizar_dataframe(df)
    col_codigo = buscar_columna(df_norm, ["CLIENTE_NRO", "CODIGO", "CODIGO_CLIENTE", "COD_CLIENTE", "CLIENTE", "Código cliente"])
    if col_codigo is None:
        col_codigo = df_norm.columns[0]
    return {limpiar_codigo(v) for v in df_norm[col_codigo].tolist() if limpiar_codigo(v)}


# =========================================================
# VENTAS
# =========================================================
def preparar_ventas(
    df_original: pd.DataFrame,
    codigos_cross_manual: set[str] | None = None,
    detectar_cross_por_tipo: bool = True,
    codigo_cross_desde: str = "",
) -> tuple[pd.DataFrame, dict[str, str | None]]:
    df = normalizar_dataframe(df_original)
    codigos_cross_manual = codigos_cross_manual or set()

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

    tipo_cross = salida["Tipo venta"].str.contains(
        r"CROSS|CROSSELL|CROSS_SELLING|CROSS SELLING|CROSSSELLING|CROSSEL|CROSELL",
        case=False,
        na=False,
        regex=True,
    ) if detectar_cross_por_tipo else pd.Series([False] * len(salida), index=salida.index)

    codigo_cross = salida["Código cliente"].isin(codigos_cross_manual)

    codigo_desde_limpio = limpiar_codigo(codigo_cross_desde)
    if codigo_desde_limpio:
        cod_num = pd.to_numeric(salida["Código cliente"], errors="coerce")
        desde_num = pd.to_numeric(pd.Series([codigo_desde_limpio]), errors="coerce").iloc[0]
        if pd.notna(desde_num):
            codigo_cross_desde_mask = cod_num >= desde_num
        else:
            codigo_cross_desde_mask = pd.Series([False] * len(salida), index=salida.index)
    else:
        codigo_cross_desde_mask = pd.Series([False] * len(salida), index=salida.index)

    salida["Es crosselling"] = tipo_cross | codigo_cross | codigo_cross_desde_mask
    salida["Motivo crosselling"] = ""
    salida.loc[tipo_cross, "Motivo crosselling"] = "Tipo venta"
    salida.loc[codigo_cross, "Motivo crosselling"] = salida.loc[codigo_cross, "Motivo crosselling"].replace("", "Código manual")
    salida.loc[codigo_cross_desde_mask, "Motivo crosselling"] = salida.loc[codigo_cross_desde_mask, "Motivo crosselling"].replace("", f"Desde código {codigo_desde_limpio}")
    salida.loc[tipo_cross & codigo_cross, "Motivo crosselling"] = "Tipo venta + código manual"
    salida.loc[tipo_cross & codigo_cross_desde_mask, "Motivo crosselling"] = salida.loc[tipo_cross & codigo_cross_desde_mask, "Motivo crosselling"].apply(lambda x: x if "+" in x else f"{x} + desde código")

    salida["Fecha texto"] = salida["Fecha"].apply(formatear_fecha)
    salida["Fecha orden"] = salida["Fecha"].apply(fecha_orden)
    return salida, columnas


def calcular_resumen(ventas: pd.DataFrame, objetivos: pd.DataFrame | None) -> pd.DataFrame:
    columnas_base = ["EH", "Socio", "Ventas objetivo", "Crosselling", "Total ventas", "Objetivo", "Cumplimiento", "Faltan"]
    if ventas.empty:
        return pd.DataFrame(columns=columnas_base)

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
    resumen = resumen.sort_values(["Cumplimiento", "Ventas objetivo", "Crosselling"], ascending=[False, False, False]).reset_index(drop=True)
    return resumen[columnas_base]


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

    ventas_objetivo = ventas_socio[~ventas_socio["Es crosselling"]].copy().sort_values(["Fecha orden", "Código cliente"])
    ventas_cross = ventas_socio[ventas_socio["Es crosselling"]].copy().sort_values(["Fecha orden", "Código cliente"])

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

    emoji_estado, estado, _, _, _ = estado_cumplimiento(cumplimiento, ventas_obj, objetivo)
    lineas.append(f"{emoji_estado} Estado: *{estado}*")
    lineas.append(mensaje_motivacional_por_estado(cumplimiento, ventas_obj, objetivo, faltan))
    lineas.append("🎯 " + reto_sugerido(cumplimiento, ventas_obj, objetivo, faltan))

    lineas += ["", "✅ *CÓDIGOS QUE CUENTAN AL OBJETIVO:*", ""]
    if ventas_objetivo.empty:
        lineas.append("Sin códigos objetivo registrados.")
    else:
        for _, r in ventas_objetivo.iterrows():
            lineas.append(f"🔹 {r['Código cliente']} | {r['Cliente'] or 'S/N'} | {r['Fecha texto']}")

    lineas += ["", "🔄 *CROSSSELLING - NO SUMA AL OBJETIVO:*", ""]
    if ventas_cross.empty:
        lineas.append("Sin crosselling registrado.")
    else:
        for _, r in ventas_cross.iterrows():
            lineas.append(f"🔸 {r['Código cliente']} | {r['Cliente'] or 'S/N'} | {r['Fecha texto']}")

    return "\n".join(lineas)


def mensaje_ranking_general(resumen: pd.DataFrame) -> str:
    total_obj = int(resumen["Ventas objetivo"].sum()) if not resumen.empty else 0
    total_cross = int(resumen["Crosselling"].sum()) if not resumen.empty else 0
    total_ventas = int(resumen["Total ventas"].sum()) if not resumen.empty else 0
    total_objetivo = int(resumen["Objetivo"].sum()) if not resumen.empty else 0
    faltan_global = max(total_objetivo - total_obj, 0)
    cumplimiento_global = (total_obj / total_objetivo * 100) if total_objetivo > 0 else 0

    lineas = [
        "🏆 *RANKING MOTIVACIONAL DE VENTAS*",
        "",
        f"✅ Ventas que cuentan al objetivo: *{total_obj}*",
        f"🔄 Crosselling separado: *{total_cross}*",
        f"📊 Total gestiones: *{total_ventas}*",
        f"🎯 Objetivo global: *{total_objetivo}*",
        f"📈 Cumplimiento global: *{cumplimiento_global:.1f}%*",
        f"⏳ Faltan para el equipo: *{faltan_global}*",
        "",
        "🥇🥈🥉 *Ranking por cumplimiento:*",
    ]

    for i, (_, r) in enumerate(resumen.iterrows(), start=1):
        icono = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "🔹"
        emoji_estado, estado, _, _, _ = estado_cumplimiento(float(r["Cumplimiento"]), int(r["Ventas objetivo"]), int(r["Objetivo"]))
        lineas.append(
            f"{icono} {r['EH']} - {r['Socio']}: *{int(r['Ventas objetivo'])}/{int(r['Objetivo'])}* "
            f"({float(r['Cumplimiento']):.1f}%) | {emoji_estado} {estado} | Cross: *{int(r['Crosselling'])}* | Faltan: *{int(r['Faltan'])}*"
        )

    cercanos = resumen[(resumen["Faltan"] > 0) & (resumen["Faltan"] <= 3)] if not resumen.empty else pd.DataFrame()
    if not cercanos.empty:
        lineas += ["", "🔥 *Socios cerca del objetivo:*"]
        for _, r in cercanos.iterrows():
            lineas.append(f"➡️ {r['Socio']} está a *{int(r['Faltan'])}* venta(s) de cumplir.")

    sin_avance = resumen[(resumen["Ventas objetivo"] == 0) & (resumen["Objetivo"] > 0)] if not resumen.empty else pd.DataFrame()
    if not sin_avance.empty:
        lineas += ["", "🚨 *Prioridad de activación:*"]
        for _, r in sin_avance.head(5).iterrows():
            lineas.append(f"➡️ {r['Socio']} necesita su primera venta objetivo.")

    if total_objetivo > 0 and faltan_global == 0:
        cierre = "🏆 Equipo, objetivo global cumplido. Mantengamos calidad y sigamos sumando ventas nuevas."
    elif total_objetivo > 0 and cumplimiento_global >= 80:
        cierre = f"🔥 Equipo, estamos cerca. Faltan {faltan_global} ventas objetivo para cerrar con fuerza."
    else:
        cierre = "💪 Equipo, enfoquémonos en ventas nuevas que cuentan al objetivo. Cada cierre suma."

    lineas += ["", f"🎯 *Reto del día:* {cierre}"]
    return "\n".join(lineas)


def mensaje_todos_los_socios(resumen: pd.DataFrame, ventas: pd.DataFrame) -> str:
    bloques = []
    for _, row in resumen.iterrows():
        det = ventas[ventas["EH"].astype(str) == str(row["EH"])]
        bloques.append(mensaje_avance_socio(row, det))
    return "\n\n------------------------------\n\n".join(bloques)


def mensaje_solo_codigos_objetivo(row: pd.Series, ventas_socio: pd.DataFrame) -> str:
    ventas_obj = ventas_socio[~ventas_socio["Es crosselling"]].copy().sort_values(["Fecha orden", "Código cliente"])
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
        for _, r in ventas_obj.iterrows():
            lineas.append(f"🔹 {r['Código cliente']} | {r['Cliente'] or 'S/N'} | {r['Fecha texto']}")
    return "\n".join(lineas)


# =========================================================
# STREAMLIT APP
# =========================================================
def cargar_datos() -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame, set[str]]:
    st.sidebar.subheader("Archivos")
    archivo_ventas = st.sidebar.file_uploader("1) Subir base de ventas", type=["csv", "xlsx", "xls"], key="ventas")
    archivo_objetivos = st.sidebar.file_uploader("2) Subir Objetivo.xlsx", type=["xlsx", "xls", "csv"], key="objetivos")

    st.sidebar.subheader("Crosselling")
    detectar_cross_tipo = st.sidebar.checkbox("Detectar CROSS_SELLING por TIPO_VENTA", value=True, key="detectar_cross_tipo")
    with st.sidebar.expander("Configurar crosselling", expanded=True):
        st.caption("El crosselling aparece en el reporte, pero NO suma al objetivo de ventas nuevas.")
        texto_cross = st.text_area("Pegar códigos crosselling", height=100, key="texto_cross_manual", placeholder="2671233\n2674567")
        archivo_cross = st.file_uploader("Subir archivo con códigos crosselling", type=["csv", "xlsx", "xls"], key="archivo_cross_manual")
        codigo_cross_desde = st.text_input("Desde este código considerar crosselling", key="codigo_cross_desde", placeholder="Ej.: 2680000")
        st.caption("Úsalo solo si desde cierto código en adelante todos deben tratarse como crosselling.")

    codigos_cross = extraer_codigos_de_texto(texto_cross) | extraer_codigos_de_archivo(archivo_cross)
    if codigos_cross:
        st.sidebar.success(f"Códigos crosselling manuales: {len(codigos_cross)}")
    if limpiar_codigo(codigo_cross_desde):
        st.sidebar.info(f"Crosselling desde código: {limpiar_codigo(codigo_cross_desde)}")

    if archivo_ventas is None:
        st.info("Sube primero la base de ventas.")
        return pd.DataFrame(), None, pd.DataFrame(), codigos_cross

    try:
        df_ventas_raw = leer_archivo(archivo_ventas)
        ventas, columnas_ventas = preparar_ventas(
            df_ventas_raw,
            codigos_cross_manual=codigos_cross,
            detectar_cross_por_tipo=detectar_cross_tipo,
            codigo_cross_desde=codigo_cross_desde,
        )
    except Exception as e:
        st.error(f"No se pudo leer la base de ventas: {e}")
        return pd.DataFrame(), None, pd.DataFrame(), codigos_cross

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
    return ventas, objetivos, resumen, codigos_cross


def mostrar_dashboard(ventas: pd.DataFrame, resumen: pd.DataFrame) -> None:
    st.markdown(
        """
        <div class="hero-box">
            <h1>📊 Ranking motivacional</h1>
            <p>Ventas objetivo separadas de crosselling, podio, semáforo y retos para activar al equipo.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if ventas.empty:
        return

    total_obj = int(resumen["Ventas objetivo"].sum())
    total_cross = int(resumen["Crosselling"].sum())
    total_ventas = int(resumen["Total ventas"].sum())
    total_objetivo = int(resumen["Objetivo"].sum())
    faltan_global = max(total_objetivo - total_obj, 0)
    cumplimiento = (total_obj / total_objetivo * 100) if total_objetivo > 0 else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        tarjeta_metric("✅ Ventas objetivo", total_obj, "Sí suman al objetivo", "#16A34A", "#22C55E")
    with c2:
        tarjeta_metric("🔄 Crosselling", total_cross, "Separado, no suma", "#7C3AED", "#A855F7")
    with c3:
        tarjeta_metric("📊 Total gestiones", total_ventas, "Objetivo + crosselling", "#0EA5E9", "#38BDF8")
    with c4:
        tarjeta_metric("🎯 Objetivo", total_objetivo, f"Faltan {faltan_global}", "#F97316", "#FB923C")
    with c5:
        tarjeta_metric("📈 Cumplimiento", f"{cumplimiento:.1f}%", "Avance global", "#0033A0", "#2563EB")

    render_barra_equipo(total_obj, total_objetivo)
    render_reto_equipo(resumen)
    render_podio(resumen)
    render_insignias(resumen)

    st.subheader("📋 Ranking completo por socio")
    st.caption("Ordenado por cumplimiento. El crosselling se muestra separado y no suma al objetivo.")
    ranking_visual = agregar_columnas_visuales(resumen)
    columnas_ranking = [
        "Puesto", "Estado", "EH", "Socio", "Avance", "Ventas objetivo", "Crosselling",
        "Total ventas", "Objetivo", "Cumplimiento", "Faltan", "Motivación", "Reto sugerido"
    ]
    st.dataframe(
        ranking_visual[columnas_ranking].style.apply(color_fila_ranking, axis=1),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("🔄 Resumen crosselling")
    cross = ventas[ventas["Es crosselling"]].copy()
    if cross.empty:
        st.info("No se detectó crosselling con la configuración actual.")
    else:
        cross_resumen = cross.groupby(["EH", "Socio"], as_index=False).agg(
            Crosselling=("Código cliente", "count")
        ).sort_values("Crosselling", ascending=False)
        st.dataframe(cross_resumen, use_container_width=True, hide_index=True)
        with st.expander("Ver detalle de crosselling", expanded=False):
            st.dataframe(
                cross[["Código cliente", "EH", "Socio", "Tipo venta", "Nodo", "Fecha texto", "Cliente", "Motivo crosselling"]],
                use_container_width=True,
                hide_index=True,
            )

    with st.expander("Detalle completo de ventas", expanded=False):
        st.dataframe(ventas, use_container_width=True, hide_index=True)

    excel = excel_bytes({"Ranking": ranking_visual, "Detalle ventas": ventas, "Crosselling": cross})
    st.download_button(
        "⬇️ Descargar Excel",
        excel,
        file_name=f"ranking_motivacional_{hoy_bolivia().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def mostrar_whatsapp(ventas: pd.DataFrame, resumen: pd.DataFrame) -> None:
    st.title("📲 WhatsApp - Avance de ventas")
    if ventas.empty:
        return

    tipo = st.selectbox(
        "Tipo de mensaje",
        ["Avance por socio", "Ranking motivacional", "Ranking general", "Todos los socios separados", "Solo códigos objetivo"],
        key="tipo_mensaje",
    )

    detalle = pd.DataFrame()
    prefijo_key = tipo

    if tipo in {"Avance por socio", "Solo códigos objetivo"}:
        opciones = [f"{r['EH']} - {r['Socio']}" for _, r in resumen.iterrows()]
        seleccion = st.selectbox("Seleccionar socio", opciones, key=f"sel_socio_{tipo}")
        eh = seleccion.split(" - ")[0]
        row = resumen[resumen["EH"].astype(str) == eh].iloc[0]
        detalle = ventas[ventas["EH"].astype(str) == eh].copy()
        prefijo_key = f"{tipo}_{eh}_{int(row['Ventas objetivo'])}_{int(row['Crosselling'])}_{int(row['Total ventas'])}"

        if tipo == "Avance por socio":
            texto = mensaje_avance_socio(row, detalle)
        else:
            texto = mensaje_solo_codigos_objetivo(row, detalle)

    elif tipo in {"Ranking general", "Ranking motivacional"}:
        texto = mensaje_ranking_general(resumen)
        detalle = agregar_columnas_visuales(resumen)
        prefijo_key = f"ranking_{tipo}_{len(resumen)}_{int(resumen['Ventas objetivo'].sum())}_{int(resumen['Crosselling'].sum())}"

    else:
        texto = mensaje_todos_los_socios(resumen, ventas)
        detalle = resumen.copy()
        prefijo_key = f"todos_{len(resumen)}_{int(resumen['Ventas objetivo'].sum())}_{int(resumen['Crosselling'].sum())}"

    st.subheader("Mensaje WhatsApp")
    st.text_area("Copiar mensaje", value=texto, height=430, key=key_mensaje(prefijo_key, texto))
    st.markdown(f"[📲 Enviar por WhatsApp]({whatsapp_link(texto)})")

    # El detalle va debajo del mensaje para no distraer del texto a copiar.
    st.subheader("Ventas / detalle del socio")
    if tipo in {"Avance por socio", "Solo códigos objetivo"}:
        st.dataframe(
            detalle[["Código cliente", "EH", "Socio", "Tipo venta", "Nodo", "Fecha texto", "Cliente", "Es crosselling", "Motivo crosselling"]],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.dataframe(detalle, use_container_width=True, hide_index=True)


def mostrar_inicio() -> None:
    st.markdown(
        """
        <div class="hero-box">
            <h1>📲 WhatsApp Objetivos Tigo V4</h1>
            <p>Ranking motivacional con podio, colores, semáforo, retos y mensajes para tus socios.</p>
            <hr style="border: 1px solid rgba(255,255,255,.25);">
            <p>👷 Desarrollado para seguimiento comercial diario</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.info("Sube la base de ventas y el archivo Objetivo.xlsx desde el panel lateral.")
    st.markdown(
        """
        ### Flujo recomendado
        1. Sube la base **GROSSADD**.
        2. Sube **Objetivo.xlsx**.
        3. Revisa **Dashboard** para ver podio, semáforo y ranking.
        4. Entra a **WhatsApp** y usa **Ranking motivacional** para el grupo.
        5. Usa **Avance por socio** para enviar mensajes individuales con reto sugerido.

        ### Nuevas mejoras V4
        - 🏆 Podio Top 3.
        - 🟢🟡🟠🔴 Semáforo por cumplimiento.
        - 🎯 Reto del día para el equipo.
        - 🎖️ Reconocimientos: líder, socio cerca del objetivo, socio impulso y mayor crosselling.
        - 📲 Mensajes WhatsApp más motivadores.
        - 🔄 Crosselling visible, pero separado del objetivo.
        """
    )


def main() -> None:
    st.set_page_config(page_title="WhatsApp Objetivos Tigo", page_icon="📲", layout="wide")
    aplicar_estilos()
    st.sidebar.title("📲 WhatsApp Objetivos")
    modulo = st.sidebar.radio("Módulo", ["Inicio", "Dashboard", "WhatsApp"], key="menu")

    ventas, objetivos, resumen, codigos_cross = cargar_datos()

    if objetivos is None:
        st.sidebar.warning("Objetivos no cargados. El cumplimiento saldrá en 0 hasta subir Objetivo.xlsx.")

    if not ventas.empty:
        st.sidebar.metric("Ventas objetivo", int((~ventas["Es crosselling"]).sum()))
        st.sidebar.metric("Crosselling", int(ventas["Es crosselling"].sum()))

    if modulo == "Inicio":
        mostrar_inicio()
    elif modulo == "Dashboard":
        mostrar_dashboard(ventas, resumen)
    elif modulo == "WhatsApp":
        mostrar_whatsapp(ventas, resumen)


if __name__ == "__main__":
    main()
