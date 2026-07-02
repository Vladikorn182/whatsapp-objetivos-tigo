# app.py
# WhatsApp Objetivos Tigo - V9
# Ranking de Ventas Top 5 por GROSSADD, crosselling automático por TIPO_VENTA y mensajes WhatsApp.

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
    for a, b in {" ": "_", "-": "_", "/": "_", ".": "_", "(": "", ")": "", "#": "", ":": ""}.items():
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
    texto = str(valor).strip()
    return "" if texto.lower() == "nan" else texto


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


def key_mensaje(prefix: str, texto: str) -> str:
    digest = hashlib.md5(texto.encode("utf-8", errors="ignore")).hexdigest()[:10]
    safe = re.sub(r"[^A-Za-z0-9_]+", "_", prefix)[:40]
    return f"msg_{safe}_{digest}"


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
# ESTILO
# =========================================================
def aplicar_estilos() -> None:
    st.markdown(
        """
        <style>
        .main .block-container {padding-top: 1.1rem; padding-bottom: 2rem;}
        .hero-box {
            padding: 26px 28px;
            border-radius: 26px;
            background: linear-gradient(135deg, #0033A0 0%, #006BEA 45%, #00AEEF 100%);
            color: white;
            box-shadow: 0 12px 28px rgba(0, 51, 160, .22);
            margin-bottom: 20px;
        }
        .hero-box h1 {font-size: 2.25rem; margin: 0 0 8px 0;}
        .hero-box p {font-size: 1.04rem; margin: 0; opacity: .96;}
        .metric-card {
            padding: 18px;
            border-radius: 21px;
            color: white;
            min-height: 116px;
            box-shadow: 0 8px 22px rgba(15, 23, 42, .16);
            border: 1px solid rgba(255, 255, 255, .24);
        }
        .metric-card .label {font-size: .86rem; opacity: .96; margin-bottom: 8px;}
        .metric-card .value {font-size: 2.05rem; font-weight: 900; line-height: 1.05;}
        .metric-card .note {font-size: .80rem; opacity: .94; margin-top: 8px;}
        .top-card {
            padding: 20px;
            border-radius: 24px;
            color: white;
            min-height: 265px;
            box-shadow: 0 10px 26px rgba(15, 23, 42, .18);
            border: 1px solid rgba(255, 255, 255, .25);
        }
        .top-card .rank {font-size: 2.1rem; font-weight: 950; margin-bottom: 12px;}
        .top-card .name {font-size: 1.05rem; font-weight: 950; line-height: 1.18; min-height: 44px;}
        .top-card .eh {font-size: .78rem; opacity: .92; margin: 9px 0 15px 0;}
        .top-card .sales {font-size: 2.45rem; font-weight: 950; line-height: 1;}
        .top-card .sub {font-size: .84rem; opacity: .97; margin-top: 8px;}
        .mini-progress-bg {height: 9px; border-radius: 999px; background: rgba(255,255,255,.28); overflow: hidden; margin-top: 12px;}
        .mini-progress-fill {height: 9px; border-radius: 999px; background: rgba(255,255,255,.92);}
        .section-card {
            padding: 16px;
            border-radius: 18px;
            background: #F8FAFC;
            border: 1px solid #E2E8F0;
            min-height: 128px;
            margin-bottom: 10px;
        }
        .section-card .title {font-weight: 900; font-size: .98rem; color: #0F172A; margin-bottom: 8px;}
        .section-card .text {font-size: .90rem; color: #334155; line-height: 1.38;}
        .info-box {
            padding: 16px 18px;
            border-radius: 18px;
            background: #ECFDF5;
            border: 1px solid #86EFAC;
            color: #14532D;
            margin: 10px 0 18px 0;
        }
        .alert-box {
            padding: 16px 18px;
            border-radius: 18px;
            background: #FFF7ED;
            border: 1px solid #FDBA74;
            color: #7C2D12;
            margin: 10px 0 18px 0;
        }
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


# =========================================================
# OBJETIVOS
# =========================================================
def columnas_objetivo_candidatas(df_obj_norm: pd.DataFrame, objetivo_global: int) -> list[tuple[str, int, int]]:
    """Devuelve columnas numéricas candidatas: (columna, suma, diferencia)."""
    excluir = {
        "pos_code", "eh", "codigo_eh", "vendedor_eh", "cod_eh", "cod_pdv", "pdv", "codigo_pos",
        "telefono", "celular", "ci", "documento", "id", "cliente_nro",
    }
    candidatos: list[tuple[str, int, int]] = []
    for col in df_obj_norm.columns:
        nombre = normalizar_columna(col)
        if nombre in excluir or any(x in nombre for x in ["nombre", "socio", "owner", "vendedor", "fecha"]):
            continue
        serie = pd.to_numeric(df_obj_norm[col], errors="coerce")
        if serie.notna().sum() == 0:
            continue
        # Evita columnas que parecen EH/POS o códigos grandes.
        valores = serie.dropna()
        if valores.empty:
            continue
        if valores.max() > 1000:
            continue
        suma = int(valores.fillna(0).sum())
        diff = abs(suma - int(objetivo_global))
        candidatos.append((col, suma, diff))
    candidatos.sort(key=lambda x: (x[2], -x[1]))
    return candidatos


def procesar_objetivos(df_obj_original: pd.DataFrame, columna_objetivo: str | None, objetivo_global: int) -> pd.DataFrame:
    df = normalizar_dataframe(df_obj_original)
    col_eh = buscar_columna(df, ["POS_CODE", "EH", "VENDEDOR_EH", "CODIGO_EH", "COD_EH", "CODIGO_POS", "COD_PDV", "PDV"])
    col_socio = buscar_columna(df, ["POS_OWNER", "SOCIO", "VENDEDOR_NOMBRE", "NOMBRE_SOCIO", "NOMBRE", "EJECUTIVO"])

    if not col_eh:
        raise ValueError("No se encontró columna de EH/POS en Objetivo.xlsx.")

    candidatos = columnas_objetivo_candidatas(df, objetivo_global)
    if columna_objetivo is None or columna_objetivo == "AUTOMATICO":
        if not candidatos:
            raise ValueError("No se encontró una columna numérica válida para el objetivo.")
        col_obj = candidatos[0][0]
    else:
        col_obj = normalizar_columna(columna_objetivo)
        if col_obj not in df.columns:
            raise ValueError(f"No existe la columna de objetivo seleccionada: {columna_objetivo}")

    salida = pd.DataFrame({
        "EH": df[col_eh].apply(limpiar_eh),
        "Socio objetivo": df[col_socio].apply(limpiar_texto) if col_socio else "",
        "Objetivo": pd.to_numeric(df[col_obj], errors="coerce").fillna(0).astype(int),
    })
    salida = salida[(salida["EH"] != "") & (salida["Objetivo"] >= 0)].copy()
    salida = salida.drop_duplicates(subset=["EH"], keep="last")
    salida.attrs["columna_objetivo_usada"] = col_obj
    return salida


# =========================================================
# VENTAS
# =========================================================
def preparar_ventas(df_original: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str | None]]:
    df = normalizar_dataframe(df_original)

    col_codigo = buscar_columna(df, ["CLIENTE_NRO", "CODIGO_CLIENTE", "COD_CLIENTE", "CODIGO", "COD", "CLIENTE"])
    col_eh = buscar_columna(df, ["VENDEDOR_EH", "EH", "POS_CODE", "CODIGO_EH", "COD_EH", "CODIGO_POS", "COD_PDV", "PDV", "EHUMANO", "EH_PROMOTOR"])
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
            "No se encontraron columnas obligatorias en la base de ventas: " + ", ".join(faltantes) +
            ". La app necesita CLIENTE_NRO, VENDEDOR_EH y TIPO_VENTA."
        )

    salida = pd.DataFrame({
        "Código cliente": df[col_codigo].apply(limpiar_codigo),
        "EH": df[col_eh].apply(limpiar_eh),
        "Socio": df[col_socio].apply(limpiar_texto) if col_socio else "SIN NOMBRE",
        "Tipo venta": df[col_tipo].apply(limpiar_texto).str.upper().str.strip(),
        "Nodo": df[col_nodo].apply(limpiar_texto).str.upper() if col_nodo else "",
        "Fecha": df[col_fecha] if col_fecha else "",
        "Cliente": df[col_cliente].apply(limpiar_texto) if col_cliente else "",
        "Teléfono 1": df[col_tel1].apply(limpiar_texto) if col_tel1 else "",
        "Teléfono 2": df[col_tel2].apply(limpiar_texto) if col_tel2 else "",
    })
    salida = salida[(salida["Código cliente"] != "") & (salida["EH"] != "")].copy()

    tipo_norm = salida["Tipo venta"].apply(lambda x: normalizar_columna(x).upper())
    salida["Es crosselling"] = tipo_norm.str.contains("CROSS", na=False)
    salida["Es venta objetivo"] = tipo_norm.str.contains("GROSS", na=False) & ~salida["Es crosselling"]
    # Para evitar distorsión: solo GROSSADD suma. Tipos desconocidos no suman ni como cross.
    salida.loc[~salida["Es venta objetivo"] & ~salida["Es crosselling"], "Es crosselling"] = False
    salida["Motivo"] = ""
    salida.loc[salida["Es venta objetivo"], "Motivo"] = "TIPO_VENTA = GROSSADD"
    salida.loc[salida["Es crosselling"], "Motivo"] = "TIPO_VENTA = CROSS_SELLING"
    salida.loc[~salida["Es venta objetivo"] & ~salida["Es crosselling"], "Motivo"] = "TIPO_VENTA no identificado"

    salida["Fecha texto"] = salida["Fecha"].apply(formatear_fecha)
    salida["Fecha orden"] = salida["Fecha"].apply(fecha_orden)
    return salida, columnas


# =========================================================
# RESUMEN
# =========================================================
def nombre_preferido(serie: pd.Series) -> str:
    for valor in serie.tolist():
        texto = limpiar_texto(valor)
        if texto and texto.upper() not in {"SIN NOMBRE", "S/N", "NAN"}:
            return texto
    return "SIN NOMBRE"


def calcular_resumen(ventas: pd.DataFrame, objetivos: pd.DataFrame | None) -> pd.DataFrame:
    columnas = ["EH", "Socio", "Ventas objetivo", "Crosselling", "Total gestiones", "Objetivo", "Cumplimiento", "Faltan"]
    if ventas.empty and (objetivos is None or objetivos.empty):
        return pd.DataFrame(columns=columnas)

    if not ventas.empty:
        ventas_resumen = ventas.groupby("EH", as_index=False).agg(
            Socio=("Socio", nombre_preferido),
            **{
                "Ventas objetivo": ("Es venta objetivo", lambda s: int(s.sum())),
                "Crosselling": ("Es crosselling", lambda s: int(s.sum())),
                "Total gestiones": ("Código cliente", "count"),
            },
        )
    else:
        ventas_resumen = pd.DataFrame(columns=["EH", "Socio", "Ventas objetivo", "Crosselling", "Total gestiones"])

    if objetivos is not None and not objetivos.empty:
        resumen = ventas_resumen.merge(objetivos[["EH", "Socio objetivo", "Objetivo"]], on="EH", how="outer")
        resumen["Socio"] = resumen.apply(
            lambda r: r["Socio"] if limpiar_texto(r.get("Socio", "")).upper() not in {"", "SIN NOMBRE", "S/N", "NAN"}
            else (limpiar_texto(r.get("Socio objetivo", "")) or "SIN NOMBRE"),
            axis=1,
        )
    else:
        resumen = ventas_resumen.copy()
        resumen["Objetivo"] = 0

    for col in ["Ventas objetivo", "Crosselling", "Total gestiones", "Objetivo"]:
        resumen[col] = pd.to_numeric(resumen.get(col, 0), errors="coerce").fillna(0).astype(int)
    resumen["EH"] = resumen["EH"].apply(limpiar_eh)
    resumen["Socio"] = resumen["Socio"].apply(limpiar_texto).replace("", "SIN NOMBRE")
    resumen["Cumplimiento"] = resumen.apply(lambda r: (r["Ventas objetivo"] / r["Objetivo"] * 100) if r["Objetivo"] > 0 else 0, axis=1)
    resumen["Faltan"] = (resumen["Objetivo"] - resumen["Ventas objetivo"]).clip(lower=0).astype(int)
    resumen = resumen.sort_values(["Ventas objetivo", "Crosselling", "Cumplimiento"], ascending=[False, False, False]).reset_index(drop=True)
    return resumen[columnas]


def resumen_visual(resumen: pd.DataFrame) -> pd.DataFrame:
    if resumen.empty:
        return resumen.copy()
    df = resumen.copy().reset_index(drop=True)
    df.insert(0, "Puesto", range(1, len(df) + 1))
    df["Avance"] = df.apply(lambda r: f"{int(r['Ventas objetivo'])}/{int(r['Objetivo'])}", axis=1)
    df["Cumplimiento"] = df["Cumplimiento"].round(1)
    df["Estado"] = df.apply(lambda r: estado_texto(int(r["Ventas objetivo"]), int(r["Objetivo"]), float(r["Cumplimiento"])), axis=1)
    return df


def estado_texto(ventas_obj: int, objetivo: int, cumplimiento: float) -> str:
    if objetivo <= 0:
        return "⚪ Sin objetivo"
    if ventas_obj >= objetivo:
        return "🏆 Cumplido"
    if cumplimiento >= 80:
        return "🔥 Cerca"
    if cumplimiento >= 60:
        return "💪 En carrera"
    if ventas_obj > 0:
        return "⚡ Necesita impulso"
    return "🚨 Sin avance"


def color_fila(row: pd.Series) -> list[str]:
    ventas = int(row.get("Ventas objetivo", 0) or 0)
    objetivo = int(row.get("Objetivo", 0) or 0)
    if objetivo > 0 and ventas >= objetivo:
        style = "background-color: #DCFCE7; color: #14532D; font-weight: 700;"
    elif objetivo > 0 and ventas / objetivo >= .8:
        style = "background-color: #DBEAFE; color: #1E3A8A; font-weight: 700;"
    elif objetivo > 0 and ventas / objetivo >= .6:
        style = "background-color: #FEF9C3; color: #713F12; font-weight: 700;"
    elif ventas > 0:
        style = "background-color: #FFEDD5; color: #7C2D12; font-weight: 700;"
    else:
        style = "background-color: #FEE2E2; color: #7F1D1D; font-weight: 700;"
    return [style for _ in row]


# =========================================================
# WHATSAPP
# =========================================================
def mensaje_recomendacion(row: pd.Series) -> str:
    ventas = int(row["Ventas objetivo"])
    objetivo = int(row["Objetivo"])
    faltan = int(row["Faltan"])
    cumplimiento = float(row["Cumplimiento"])
    if objetivo <= 0:
        return "Revisar objetivo cargado para medir su avance."
    if ventas >= objetivo:
        return "Mantener calidad, cuidar instalaciones y seguir sumando ventas nuevas."
    if faltan <= 3:
        return f"Está a solo {faltan} venta(s). Priorizar cierres calientes hoy."
    if cumplimiento >= 80:
        return "Está cerca. Meta diaria sugerida: mínimo 1 venta nueva hasta cumplir."
    if cumplimiento >= 60:
        return "Buen avance, necesita empuje. Revisar cartera y recuperar 2 cierres esta semana."
    if ventas > 0:
        return "Necesita acelerar. Enfocar seguimiento a pendientes y prospectos con mayor intención."
    return "Prioridad: lograr la primera venta nueva del día."


def mensaje_ranking(resumen: pd.DataFrame) -> str:
    total_obj = int(resumen["Ventas objetivo"].sum()) if not resumen.empty else 0
    total_cross = int(resumen["Crosselling"].sum()) if not resumen.empty else 0
    total = int(resumen["Total gestiones"].sum()) if not resumen.empty else 0
    objetivo_global = int(st.session_state.get("objetivo_global", 400))
    cumplimiento = (total_obj / objetivo_global * 100) if objetivo_global > 0 else 0
    lineas = [
        "🏆 *RANKING DE VENTAS - TOP 5*",
        "",
        f"✅ Ventas nuevas que suman: *{total_obj}*",
        f"🔄 Crosselling separado: *{total_cross}*",
        f"📊 Total gestiones: *{total}*",
        f"🎯 Objetivo global: *{objetivo_global}*",
        f"📈 Cumplimiento global: *{cumplimiento:.1f}%*",
        "",
        "*Top 5 por ventas nuevas:*",
    ]
    iconos = ["🥇", "🥈", "🥉", "🏅", "🏅"]
    for i, (_, r) in enumerate(resumen.head(5).iterrows()):
        lineas.append(
            f"{iconos[i]} {r['Socio']} | EH {r['EH']} | *{int(r['Ventas objetivo'])} ventas* "
            f"| Obj. {int(r['Objetivo'])} | Cross {int(r['Crosselling'])}"
        )
    return "\n".join(lineas)


def mensaje_meta_superada(resumen: pd.DataFrame) -> str:
    df = resumen[(resumen["Objetivo"] > 0) & (resumen["Ventas objetivo"] >= resumen["Objetivo"])].copy()
    lineas = ["🚀 *SOCIOS QUE SOBREPASARON SU OBJETIVO*", ""]
    if df.empty:
        lineas.append("Todavía no hay socios con objetivo sobrepasado.")
    else:
        df["Arriba"] = df["Ventas objetivo"] - df["Objetivo"]
        for _, r in df.sort_values(["Arriba", "Ventas objetivo"], ascending=[False, False]).head(10).iterrows():
            lineas.append(f"🏆 {r['Socio']} | EH {r['EH']} | {int(r['Ventas objetivo'])}/{int(r['Objetivo'])} | +{int(r['Arriba'])} venta(s)")
    return "\n".join(lineas)


def mensaje_top_cross(resumen: pd.DataFrame) -> str:
    df = resumen[resumen["Crosselling"] > 0].sort_values(["Crosselling", "Ventas objetivo"], ascending=[False, False]).head(10)
    lineas = ["🔄 *TOP CROSSSELLING*", ""]
    if df.empty:
        lineas.append("No hay crosselling registrado.")
    else:
        for i, (_, r) in enumerate(df.iterrows(), start=1):
            lineas.append(f"{i}. {r['Socio']} | EH {r['EH']} | Crosselling: *{int(r['Crosselling'])}* | Ventas nuevas: {int(r['Ventas objetivo'])}")
    return "\n".join(lineas)


def mensaje_pendientes(resumen: pd.DataFrame) -> str:
    df = resumen[(resumen["Objetivo"] > 0) & (resumen["Faltan"] > 0)].sort_values(["Faltan", "Ventas objetivo"], ascending=[True, False]).head(15)
    lineas = ["📌 *RECOMENDACIONES PARA CUMPLIR OBJETIVO*", ""]
    if df.empty:
        lineas.append("Todos los socios con objetivo cargado ya cumplieron.")
    else:
        for _, r in df.iterrows():
            lineas.append(
                f"👤 *{r['Socio']}* | EH {r['EH']}\n"
                f"✅ Avance: {int(r['Ventas objetivo'])}/{int(r['Objetivo'])} ({float(r['Cumplimiento']):.1f}%) | Faltan: *{int(r['Faltan'])}*\n"
                f"🎯 {mensaje_recomendacion(r)}\n"
            )
    return "\n".join(lineas)


def mensaje_avance_socio(row: pd.Series, ventas_socio: pd.DataFrame) -> str:
    ventas_obj = ventas_socio[ventas_socio["Es venta objetivo"]].copy().sort_values(["Fecha orden", "Código cliente"])
    ventas_cross = ventas_socio[ventas_socio["Es crosselling"]].copy().sort_values(["Fecha orden", "Código cliente"])
    lineas = [
        "📊 *AVANCE DE VENTAS*",
        "",
        f"👤 *{row['Socio']}*",
        f"EH: *{row['EH']}*",
        "",
        f"✅ Ventas nuevas: *{int(row['Ventas objetivo'])}*",
        f"🔄 Crosselling: *{int(row['Crosselling'])}*",
        f"📊 Total gestiones: *{int(row['Total gestiones'])}*",
        "",
        f"🎯 Objetivo: *{int(row['Objetivo'])}*",
        f"📈 Cumplimiento: *{float(row['Cumplimiento']):.1f}%*",
        f"⏳ Faltan: *{int(row['Faltan'])}*",
        f"📌 {mensaje_recomendacion(row)}",
        "",
        "✅ *CÓDIGOS QUE SUMAN AL OBJETIVO:*",
    ]
    if ventas_obj.empty:
        lineas.append("Sin códigos de venta nueva registrados.")
    else:
        for _, r in ventas_obj.iterrows():
            lineas.append(f"🔹 {r['Código cliente']} | {r['Cliente'] or 'S/N'} | {r['Fecha texto']}")
    lineas += ["", "🔄 *CROSSSELLING - NO SUMA AL OBJETIVO:*"]
    if ventas_cross.empty:
        lineas.append("Sin crosselling registrado.")
    else:
        for _, r in ventas_cross.iterrows():
            lineas.append(f"🔸 {r['Código cliente']} | {r['Cliente'] or 'S/N'} | {r['Fecha texto']}")
    return "\n".join(lineas)


def mensaje_codigos_objetivo(row: pd.Series, ventas_socio: pd.DataFrame) -> str:
    ventas_obj = ventas_socio[ventas_socio["Es venta objetivo"]].copy().sort_values(["Fecha orden", "Código cliente"])
    lineas = ["✅ *CÓDIGOS QUE SUMAN AL OBJETIVO*", "", f"👤 *{row['Socio']}*", f"EH: *{row['EH']}*", f"Total: *{len(ventas_obj)}*", ""]
    if ventas_obj.empty:
        lineas.append("Sin códigos de venta nueva registrados.")
    else:
        for _, r in ventas_obj.iterrows():
            lineas.append(f"🔹 {r['Código cliente']} | {r['Cliente'] or 'S/N'} | {r['Fecha texto']}")
    return "\n".join(lineas)


# =========================================================
# RENDER VISUAL
# =========================================================
def barra_equipo(total_obj: int, objetivo_global: int) -> None:
    avance = (total_obj / objetivo_global) if objetivo_global > 0 else 0
    st.caption(f"Avance global: {total_obj}/{objetivo_global} ventas nuevas ({avance*100:.1f}%)")
    st.progress(min(avance, 1.0))


def render_top5(resumen: pd.DataFrame) -> None:
    st.subheader("🏆 Top 5 socios con más ventas nuevas")
    st.caption("Este Top 5 se ordena solo por ventas nuevas GROSSADD. No usa porcentaje de cumplimiento para definir el puesto.")
    top = resumen.sort_values(["Ventas objetivo", "Crosselling", "Total gestiones"], ascending=[False, False, False]).head(5).reset_index(drop=True)
    if top.empty:
        st.info("No hay ventas para mostrar.")
        return

    st.markdown("""<div class="info-box">✅ Orden aplicado: <b>Ventas nuevas GROSSADD de mayor a menor</b>. El porcentaje y el objetivo solo se muestran como referencia.</div>""", unsafe_allow_html=True)
    cols = st.columns(5)
    medallas = ["🥇", "🥈", "🥉", "🏅", "🏅"]
    colores = [
        ("#F59E0B", "#F97316"),
        ("#64748B", "#94A3B8"),
        ("#B45309", "#D97706"),
        ("#0033A0", "#2563EB"),
        ("#0F766E", "#14B8A6"),
    ]
    for i, (_, r) in enumerate(top.iterrows()):
        ventas = int(r["Ventas objetivo"])
        objetivo = int(r["Objetivo"])
        cumplimiento = float(r["Cumplimiento"])
        cross = int(r["Crosselling"])
        faltan = int(r["Faltan"])
        c1, c2 = colores[i]
        barra = min(max(cumplimiento, 0), 100)
        estado = "Objetivo cumplido" if objetivo > 0 and ventas >= objetivo else f"Faltan {faltan}" if objetivo > 0 else "Sin objetivo"
        cols[i].markdown(
            f"""
            <div class="top-card" style="background: linear-gradient(135deg, {c1}, {c2});">
                <div class="rank">{medallas[i]} #{i + 1}</div>
                <div class="name">{r['Socio']}</div>
                <div class="eh">EH {r['EH']}</div>
                <div class="sales">{ventas}</div>
                <div class="sub"><b>ventas nuevas que suman al ranking</b></div>
                <div class="mini-progress-bg"><div class="mini-progress-fill" style="width:{barra:.0f}%;"></div></div>
                <div class="sub">🎯 Objetivo: {objetivo} | {cumplimiento:.1f}%</div>
                <div class="sub">🔄 Crosselling: {cross}</div>
                <div class="sub">{estado}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_secciones(resumen: pd.DataFrame) -> None:
    st.subheader("🎖️ Reconocimientos y recomendaciones")
    con_obj = resumen[resumen["Objetivo"] > 0].copy()
    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown("### 🚀 Sobrepasaron objetivo")
        cumplidos = con_obj[con_obj["Ventas objetivo"] >= con_obj["Objetivo"]].copy()
        if cumplidos.empty:
            st.info("Aún no hay socios con objetivo cumplido.")
        else:
            cumplidos["Arriba"] = cumplidos["Ventas objetivo"] - cumplidos["Objetivo"]
            for _, r in cumplidos.sort_values(["Arriba", "Ventas objetivo"], ascending=[False, False]).head(5).iterrows():
                st.markdown(
                    f"""<div class="section-card"><div class="title">🏆 {r['Socio']}</div>
                    <div class="text">EH {r['EH']}<br>Avance: <b>{int(r['Ventas objetivo'])}/{int(r['Objetivo'])}</b> ({float(r['Cumplimiento']):.1f}%)<br>Arriba de meta: <b>{int(r['Arriba'])}</b> venta(s)</div></div>""",
                    unsafe_allow_html=True,
                )

    with c2:
        st.markdown("### 🔄 Más crosselling")
        cross_top = resumen[resumen["Crosselling"] > 0].sort_values(["Crosselling", "Ventas objetivo"], ascending=[False, False]).head(5)
        if cross_top.empty:
            st.info("No hay crosselling registrado.")
        else:
            for _, r in cross_top.iterrows():
                st.markdown(
                    f"""<div class="section-card"><div class="title">🔄 {r['Socio']}</div>
                    <div class="text">EH {r['EH']}<br>Crosselling: <b>{int(r['Crosselling'])}</b><br>Ventas nuevas: <b>{int(r['Ventas objetivo'])}</b></div></div>""",
                    unsafe_allow_html=True,
                )

    with c3:
        st.markdown("### 📌 Para cumplir objetivo")
        pendientes = con_obj[con_obj["Faltan"] > 0].sort_values(["Faltan", "Ventas objetivo"], ascending=[True, False]).head(5)
        if pendientes.empty:
            st.success("Todos los socios con objetivo cargado ya cumplieron.")
        else:
            for _, r in pendientes.iterrows():
                st.markdown(
                    f"""<div class="section-card"><div class="title">🎯 {r['Socio']}</div>
                    <div class="text">EH {r['EH']}<br>Avance: <b>{int(r['Ventas objetivo'])}/{int(r['Objetivo'])}</b> · faltan <b>{int(r['Faltan'])}</b><br>{mensaje_recomendacion(r)}</div></div>""",
                    unsafe_allow_html=True,
                )


# =========================================================
# CARGA DE DATOS
# =========================================================
def cargar_datos() -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame]:
    st.sidebar.subheader("Archivos")
    archivo_ventas = st.sidebar.file_uploader("1) Subir base de ventas", type=["csv", "xlsx", "xls"], key="ventas_v9")
    archivo_objetivos = st.sidebar.file_uploader("2) Subir Objetivo.xlsx", type=["xlsx", "xls", "csv"], key="objetivos_v9")

    st.sidebar.subheader("Objetivo mensual")
    objetivo_global = st.sidebar.number_input("Objetivo global del mes", min_value=0, value=400, step=1, key="objetivo_global_v9")
    st.session_state["objetivo_global"] = int(objetivo_global)

    st.sidebar.subheader("Crosselling")
    st.sidebar.success("Automático por TIPO_VENTA")
    st.sidebar.caption("GROSSADD = venta nueva y suma al objetivo. CROSS_SELLING = crosselling separado y no suma.")

    if archivo_ventas is None:
        st.info("Sube primero la base de ventas.")
        return pd.DataFrame(), None, pd.DataFrame()

    try:
        raw_ventas = leer_archivo(archivo_ventas)
        ventas, columnas = preparar_ventas(raw_ventas)
    except Exception as e:
        st.error(f"No se pudo leer la base de ventas: {e}")
        return pd.DataFrame(), None, pd.DataFrame()

    objetivos = None
    if archivo_objetivos is not None:
        try:
            raw_obj = leer_archivo(archivo_objetivos)
            obj_norm = normalizar_dataframe(raw_obj)
            candidatos = columnas_objetivo_candidatas(obj_norm, int(objetivo_global))
            opciones = ["AUTOMATICO"] + [c[0] for c in candidatos]
            etiquetas = {
                "AUTOMATICO": "Automático recomendado",
                **{c: f"{c}  | suma aproximada: {s}" for c, s, _ in candidatos},
            }
            seleccion = st.sidebar.selectbox(
                "Columna de objetivo",
                opciones,
                index=0,
                format_func=lambda x: etiquetas.get(x, x),
                key="col_objetivo_v9",
            )
            objetivos = procesar_objetivos(raw_obj, None if seleccion == "AUTOMATICO" else seleccion, int(objetivo_global))
            suma_obj = int(objetivos["Objetivo"].sum()) if not objetivos.empty else 0
            col_usada = objetivos.attrs.get("columna_objetivo_usada", "")
            st.sidebar.metric("Suma objetivos socios", suma_obj)
            st.sidebar.caption(f"Columna usada: {col_usada}")
            if suma_obj != int(objetivo_global):
                st.sidebar.warning(f"La suma individual ({suma_obj}) no coincide con el objetivo global ({int(objetivo_global)}). Puedes revisar la columna seleccionada.")
        except Exception as e:
            st.sidebar.error(f"No se pudo leer Objetivo.xlsx: {e}")
            objetivos = None

    resumen = calcular_resumen(ventas, objetivos)
    return ventas, objetivos, resumen


# =========================================================
# PANTALLAS
# =========================================================
def mostrar_dashboard(ventas: pd.DataFrame, resumen: pd.DataFrame) -> None:
    st.markdown(
        """
        <div class="hero-box">
            <h1>🏆 Ranking de Ventas - Top 5</h1>
            <p>Top 5 por cantidad de ventas nuevas GROSSADD. El primer lugar es quien más ventas hizo; el crosselling queda separado.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if ventas.empty:
        return

    total_obj = int(resumen["Ventas objetivo"].sum())
    total_cross = int(resumen["Crosselling"].sum())
    total_gest = int(resumen["Total gestiones"].sum())
    objetivo_global = int(st.session_state.get("objetivo_global", 400))
    faltan = max(objetivo_global - total_obj, 0)
    cumplimiento = (total_obj / objetivo_global * 100) if objetivo_global > 0 else 0
    suma_individual = int(resumen["Objetivo"].sum())

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        tarjeta_metric("✅ Ventas nuevas", total_obj, "GROSSADD - sí suma", "#16A34A", "#22C55E")
    with c2:
        tarjeta_metric("🔄 Crosselling", total_cross, "Separado, no suma", "#7C3AED", "#A855F7")
    with c3:
        tarjeta_metric("📊 Total gestiones", total_gest, "Ventas + crosselling", "#0EA5E9", "#38BDF8")
    with c4:
        tarjeta_metric("🎯 Objetivo global", objetivo_global, f"Faltan {faltan}", "#F97316", "#FB923C")
    with c5:
        tarjeta_metric("📈 Cumplimiento", f"{cumplimiento:.1f}%", "Avance global", "#0033A0", "#2563EB")

    barra_equipo(total_obj, objetivo_global)

    if suma_individual and suma_individual != objetivo_global:
        st.markdown(
            f"""<div class="alert-box">⚠️ La suma de objetivos individuales cargada es <b>{suma_individual}</b>, pero el objetivo global está configurado en <b>{objetivo_global}</b>. El ranking por ventas no cambia; revisa la columna de objetivo seleccionada si necesitas que también cuadre por socio.</div>""",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """<div class="info-box">✅ Crosselling automático desde TIPO_VENTA. No se necesita código referencial ni archivo adicional.</div>""",
            unsafe_allow_html=True,
        )

    with st.expander("🔎 Control de conteo", expanded=False):
        conteo_tipo = ventas.groupby("Tipo venta", as_index=False).agg(Cantidad=("Código cliente", "count"))
        st.dataframe(conteo_tipo, use_container_width=True, hide_index=True)
        st.caption("Regla aplicada: GROSSADD suma al ranking y al objetivo. CROSS_SELLING queda separado.")

    render_top5(resumen)
    render_secciones(resumen)

    st.subheader("📲 Envío rápido por WhatsApp")
    w1, w2, w3, w4 = st.columns(4)
    with w1:
        st.markdown(f"[🏆 Ranking Top 5]({whatsapp_link(mensaje_ranking(resumen))})")
    with w2:
        st.markdown(f"[🚀 Meta superada]({whatsapp_link(mensaje_meta_superada(resumen))})")
    with w3:
        st.markdown(f"[🔄 Top crosselling]({whatsapp_link(mensaje_top_cross(resumen))})")
    with w4:
        st.markdown(f"[📌 Por cumplir]({whatsapp_link(mensaje_pendientes(resumen))})")

    st.subheader("📋 Ranking completo por socio")
    ranking = resumen_visual(resumen)
    columnas = ["Puesto", "EH", "Socio", "Ventas objetivo", "Crosselling", "Total gestiones", "Objetivo", "Cumplimiento", "Faltan", "Estado"]
    st.dataframe(ranking[columnas].style.apply(color_fila, axis=1), use_container_width=True, hide_index=True)

    with st.expander("Detalle completo de ventas", expanded=False):
        st.dataframe(ventas, use_container_width=True, hide_index=True)

    cross = ventas[ventas["Es crosselling"]].copy()
    excel = excel_bytes({"Ranking": ranking, "Detalle ventas": ventas, "Crosselling": cross})
    st.download_button(
        "⬇️ Descargar Excel",
        excel,
        file_name=f"ranking_ventas_{hoy_bolivia().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def mostrar_whatsapp(ventas: pd.DataFrame, resumen: pd.DataFrame) -> None:
    st.title("📲 WhatsApp - Ventas")
    if ventas.empty:
        return

    tipo = st.selectbox(
        "Tipo de mensaje",
        ["Ranking Top 5", "Avance por socio", "Socios por cumplir", "Socios con meta superada", "Top crosselling", "Solo códigos objetivo"],
        key="tipo_mensaje_v9",
    )

    detalle = pd.DataFrame()
    prefijo = tipo

    if tipo in {"Avance por socio", "Solo códigos objetivo"}:
        opciones = [f"{r['EH']} - {r['Socio']}" for _, r in resumen.iterrows()]
        seleccion = st.selectbox("Seleccionar socio", opciones, key=f"sel_socio_v9_{tipo}")
        eh = seleccion.split(" - ")[0]
        row = resumen[resumen["EH"].astype(str) == eh].iloc[0]
        detalle = ventas[ventas["EH"].astype(str) == eh].copy()
        texto = mensaje_avance_socio(row, detalle) if tipo == "Avance por socio" else mensaje_codigos_objetivo(row, detalle)
        prefijo = f"{tipo}_{eh}_{int(row['Ventas objetivo'])}_{int(row['Crosselling'])}"
    elif tipo == "Ranking Top 5":
        texto = mensaje_ranking(resumen)
        detalle = resumen_visual(resumen).head(5)
    elif tipo == "Socios por cumplir":
        texto = mensaje_pendientes(resumen)
        detalle = resumen[(resumen["Objetivo"] > 0) & (resumen["Faltan"] > 0)].copy()
    elif tipo == "Socios con meta superada":
        texto = mensaje_meta_superada(resumen)
        detalle = resumen[(resumen["Objetivo"] > 0) & (resumen["Ventas objetivo"] >= resumen["Objetivo"])].copy()
    else:
        texto = mensaje_top_cross(resumen)
        detalle = resumen[resumen["Crosselling"] > 0].sort_values(["Crosselling", "Ventas objetivo"], ascending=[False, False]).copy()

    st.subheader("Mensaje WhatsApp")
    st.text_area("Copiar mensaje", value=texto, height=430, key=key_mensaje(prefijo, texto))
    st.markdown(f"[📲 Enviar por WhatsApp]({whatsapp_link(texto)})")

    st.subheader("Detalle")
    if tipo in {"Avance por socio", "Solo códigos objetivo"}:
        if detalle.empty:
            st.info("Este socio no tiene ventas registradas en la base cargada.")
        else:
            st.dataframe(
                detalle[["Código cliente", "EH", "Socio", "Tipo venta", "Nodo", "Fecha texto", "Cliente", "Es venta objetivo", "Es crosselling"]],
                use_container_width=True,
                hide_index=True,
            )
    else:
        st.dataframe(detalle, use_container_width=True, hide_index=True)


def mostrar_inicio() -> None:
    st.markdown(
        """
        <div class="hero-box">
            <h1>📲 WhatsApp Objetivos Tigo V9</h1>
            <p>Dashboard corregido: Top 5 por ventas nuevas, crosselling separado y WhatsApp para seguimiento.</p>
            <hr style="border: 1px solid rgba(255,255,255,.25);">
            <p>👷 Seguimiento comercial diario</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.info("Sube la base GROSSADD y Objetivo.xlsx desde el panel lateral.")
    st.markdown(
        """
        ### Flujo recomendado
        1. Sube la base **GROSSADD**.
        2. Sube **Objetivo.xlsx**.
        3. Revisa **Dashboard** para ver el **Top 5 por ventas nuevas**.
        4. Usa **WhatsApp** para enviar ranking, recomendaciones y seguimiento por socio.

        ### Regla de conteo
        - **TIPO_VENTA = GROSSADD** → venta nueva, suma al ranking y al objetivo.
        - **TIPO_VENTA = CROSS_SELLING** → crosselling, se muestra separado y no suma.
        """
    )


def main() -> None:
    st.set_page_config(page_title="WhatsApp Objetivos Tigo V9", page_icon="📲", layout="wide")
    aplicar_estilos()
    st.sidebar.title("📲 WhatsApp Objetivos V9")
    modulo = st.sidebar.radio("Módulo", ["Inicio", "Dashboard", "WhatsApp"], key="menu_v9")

    ventas, objetivos, resumen = cargar_datos()

    if objetivos is None:
        st.sidebar.warning("Objetivos no cargados. El ranking por ventas funciona, pero el cumplimiento por socio quedará en 0.")

    if not ventas.empty:
        st.sidebar.metric("Ventas nuevas", int(ventas["Es venta objetivo"].sum()))
        st.sidebar.metric("Crosselling", int(ventas["Es crosselling"].sum()))

    if modulo == "Inicio":
        mostrar_inicio()
    elif modulo == "Dashboard":
        mostrar_dashboard(ventas, resumen)
    elif modulo == "WhatsApp":
        mostrar_whatsapp(ventas, resumen)


if __name__ == "__main__":
    main()
