# app.py
# Autogestión Tigo - App única sin módulos externos ni carpeta pages
# Incluye: Agenda Técnica, Pendientes de Instalación +3 días y Suspendidas +3 días.

from __future__ import annotations

import re
import unicodedata
import warnings
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
# CONFIGURACIÓN GENERAL
# =========================================================
SOCIOS_EH_DEFAULT = {
    "59509": "Adriana Paola Villafuerte Guerra",
    "63483": "Franklin Ramiro Quispe Rosas",
    "89859": "José Pablo Fernández Puente",
    "88463": "Alicia Graciela Zamora Buezo",
    "88426": "Sonia Noemí Mayta",
    "58984": "María Surco Aruquipa",
    "78099": "Geovana Carla Siñani Luna",
    "78340": "Olivia Sánchez Quispe",
    "89326": "Palmira Selaes Herrera",
    "83457": "Estrella Belén Quispe Flores",
    "89231": "Alex Rudy Mamani Guarachi",
    "72210": "Guadalupe Apaza Vila",
    "67755": "Sandro Iván Copa Velasco",
    "86737": "Anahi Oinca",
    "86963": "Teresa Eugenia Chipana Mamani De Sayes",
    "79030": "Víctor Hugo Chambilla Flores",
    "88874": "Gustavo Callejas Mamani",
    "77735": "Claudia Michme Ajno",
    "87933": "Pamela Mery Rojas Alarcón",
    "78272": "My Phone SRL",
}


# =========================================================
# UTILIDADES COMUNES
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
        " ": "_",
        "-": "_",
        "/": "_",
        ".": "_",
        "(": "",
        ")": "",
        "#": "",
        ":": "",
    }
    for origen, destino in reemplazos.items():
        texto = texto.replace(origen, destino)
    texto = re.sub(r"_+", "_", texto)
    return texto.strip("_")


def normalizar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalizar_columna(c) for c in df.columns]
    return df


def buscar_columna(df: pd.DataFrame, opciones: list[str]) -> str | None:
    columnas = set(df.columns)
    for opcion in opciones:
        opcion_norm = normalizar_columna(opcion)
        if opcion_norm in columnas:
            return opcion_norm
    return None


def obtener_serie(df: pd.DataFrame, opciones: list[str]) -> pd.Series:
    columna = buscar_columna(df, opciones)
    if columna is None:
        return pd.Series([""] * len(df), index=df.index)
    return df[columna]


def limpiar_texto(valor: object) -> str:
    if valor is None or pd.isna(valor):
        return ""
    return str(valor).strip()


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


def parsear_fecha(serie: pd.Series) -> pd.Series:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        fecha = pd.to_datetime(serie, errors="coerce", dayfirst=True)
    return fecha.dt.date


def leer_archivo(archivo) -> pd.DataFrame:
    nombre = getattr(archivo, "name", "").lower()
    if nombre.endswith(".csv"):
        return pd.read_csv(archivo, sep=None, engine="python")
    return pd.read_excel(archivo, engine="openpyxl")


def obtener_hojas_excel(archivo) -> list[str]:
    nombre = getattr(archivo, "name", "").lower()
    if nombre.endswith(".csv"):
        return []
    posicion = archivo.tell()
    try:
        archivo.seek(0)
        xls = pd.ExcelFile(archivo, engine="openpyxl")
        return xls.sheet_names
    finally:
        archivo.seek(posicion)


def leer_archivo_con_hoja(archivo, hoja: str | None = None) -> pd.DataFrame:
    nombre = getattr(archivo, "name", "").lower()
    if nombre.endswith(".csv"):
        return pd.read_csv(archivo, sep=None, engine="python")
    return pd.read_excel(archivo, sheet_name=hoja, engine="openpyxl") if hoja else pd.read_excel(archivo, engine="openpyxl")


def construir_socios(texto_extra: str = "") -> dict[str, str]:
    socios = dict(SOCIOS_EH_DEFAULT)
    for linea in str(texto_extra).splitlines():
        linea = linea.strip()
        if not linea:
            continue
        match = re.match(r"^(\d{3,8})\s*[-|,;:]?\s*(.+)$", linea)
        if match:
            socios[match.group(1)] = match.group(2).strip()
    return socios


# Objetivos base tomados del formato Objetivo.xlsx:
# POS_CODE = EH, POS_OWNER = Socio, BU JUNIO = Objetivo.
OBJETIVOS_BASE = [
    ("91207", "GUALBERTO FERNANDO SANJINES", "PYME DIGITAL", 5),
    ("91208", "DANNY QUISBERT MENDOZA", "PYME DIGITAL", 5),
    ("88874", "GUSTAVO CALLEJAS", "PYME DIGITAL", 10),
    ("88463", "ALICIA GRACIELA ZAMORA BUEZO", "PYME DIGITAL", 30),
    ("86963", "TERESA CHIPANA", "PYME DIGITAL", 10),
    ("79030", "VICTOR HUGO CHAMBILLA FLORES", "PYME DIGITAL", 5),
    ("88426", "SONIA NOEMI MAYTA", "PYME DIGITAL", 20),
    ("91262", "WARNES RIVERA CHUQUIMIA", "SOCIO HOGAR", 10),
    ("91283", "SOLAGEL QUENTA GUTIERREZ", "SOCIO HOGAR", 10),
    ("89859", "JOSE PABLO FERNANDEZ", "SOCIO HOGAR", 30),
    ("89326", "PALMIRA SELAES", "SOCIO HOGAR", 20),
    ("83457", "ESTRELLA BELEN QUISPE FLORES", "SOCIO HOGAR", 25),
    ("63483", "FRANKLIN RAMIRO QUISPE ROSAS", "SOCIO HOGAR", 40),
    ("72210", "GUADALUPE APAZA VILA", "SOCIO HOGAR", 20),
    ("59509", "ADRIANA PAOLA VILLAFUERTE GUERRA", "SOCIO HOGAR", 75),
    ("58984", "MARIA SURCO ARUQUIPA", "SOCIO HOGAR", 25),
    ("89231", "ALEX RUDY MAMANI GUARACHI", "SOCIO HOGAR", 15),
    ("86737", "ANAHI OINCA", "SOCIO HOGAR", 10),
    ("78340", "OLIVIA SANCHEZ QUISPE", "SOCIO HOGAR", 15),
    ("78099", "GEOVANA CARLA SIÑANI LUNA", "SOCIO HOGAR", 20),
]


def objetivos_default_df() -> pd.DataFrame:
    return pd.DataFrame(OBJETIVOS_BASE, columns=["EH", "Socio", "Categoria", "Objetivo"])


def normalizar_objetivos_archivo(df_obj_original: pd.DataFrame) -> pd.DataFrame:
    """Lee objetivos en distintos formatos.

    Formato principal esperado del archivo Objetivo.xlsx:
    POS_CODE, POS_OWNER, CATEGORIA, BU JUNIO.
    """
    df_obj = normalizar_dataframe(df_obj_original)

    col_eh = buscar_columna(df_obj, [
        "EH", "VENDEDOR_EH", "CODIGO_EH", "COD_EH", "EHUMANO", "CA",
        "POS_CODE", "POS CODE", "CODIGO", "CODIGO_SOCIO", "CODIGO VENDEDOR",
    ])
    col_socio = buscar_columna(df_obj, [
        "SOCIO", "VENDEDOR_NOMBRE", "NOMBRE_VENDEDOR", "NOMBRE_SOCIO", "VENDEDOR",
        "POS_OWNER", "POS OWNER", "NOMBRE", "CA_NOMBRE",
    ])
    col_categoria = buscar_columna(df_obj, ["CATEGORIA", "CATEGORÍA", "TIPO_SOCIO", "TIPO"])
    col_obj = buscar_columna(df_obj, [
        "OBJETIVO", "META", "CUOTA", "TARGET", "BU JUNIO", "BU_JUNIO", "JUNIO", "BU",
        "OBJETIVO_JUNIO", "META_JUNIO",
    ])

    if col_eh is None:
        raise ValueError(
            "No se encontró la columna de EH/código. Para tu archivo de objetivos se acepta POS_CODE como EH. "
            "Columnas detectadas: " + ", ".join(df_obj.columns[:80])
        )
    if col_obj is None:
        raise ValueError(
            "No se encontró la columna de objetivo/meta. Para tu archivo se acepta BU JUNIO como objetivo. "
            "Columnas detectadas: " + ", ".join(df_obj.columns[:80])
        )

    salida = pd.DataFrame({
        "EH": df_obj[col_eh].apply(limpiar_eh),
        "Socio": df_obj[col_socio].apply(limpiar_texto) if col_socio else "",
        "Categoria": df_obj[col_categoria].apply(limpiar_texto) if col_categoria else "",
        "Objetivo": pd.to_numeric(df_obj[col_obj], errors="coerce").fillna(0).astype(int),
    })
    salida = salida[salida["EH"].astype(str).str.strip().ne("")].copy()
    salida = salida[salida["Objetivo"].fillna(0).astype(int) >= 0].copy()

    # Completar nombre con el maestro si viene vacío.
    salida["Socio"] = salida.apply(
        lambda r: r["Socio"] or SOCIOS_EH_DEFAULT.get(str(r["EH"]), "SIN NOMBRE"), axis=1
    )

    # Si hay EH repetidos, suma objetivos y conserva el primer nombre/categoría.
    salida = (
        salida.groupby("EH", as_index=False)
        .agg({"Socio": "first", "Categoria": "first", "Objetivo": "sum"})
        .sort_values(["Categoria", "Socio", "EH"])
        .reset_index(drop=True)
    )
    return salida


def get_objetivos_df() -> pd.DataFrame:
    if "objetivos_df" not in st.session_state:
        st.session_state["objetivos_df"] = objetivos_default_df()
    return st.session_state["objetivos_df"].copy()


def set_objetivos_df(df_obj: pd.DataFrame) -> None:
    limpio = df_obj.copy()
    if "EH" in limpio.columns:
        limpio["EH"] = limpio["EH"].apply(limpiar_eh)
    if "Objetivo" in limpio.columns:
        limpio["Objetivo"] = pd.to_numeric(limpio["Objetivo"], errors="coerce").fillna(0).astype(int)
    st.session_state["objetivos_df"] = limpio


def objetivos_maps() -> tuple[dict[str, int], dict[str, str], dict[str, str]]:
    obj = get_objetivos_df()
    obj["EH"] = obj["EH"].apply(limpiar_eh)
    obj["Objetivo"] = pd.to_numeric(obj.get("Objetivo", 0), errors="coerce").fillna(0).astype(int)
    objetivo_map = dict(zip(obj["EH"], obj["Objetivo"]))
    socio_map = dict(zip(obj["EH"], obj.get("Socio", pd.Series([""] * len(obj))).fillna("")))
    categoria_map = dict(zip(obj["EH"], obj.get("Categoria", pd.Series([""] * len(obj))).fillna("")))
    return objetivo_map, socio_map, categoria_map


def cargar_objetivos_widget(key_prefix: str) -> None:
    with st.expander("🎯 Objetivos cargados", expanded=False):
        st.caption("Tu archivo Objetivo.xlsx usa POS_CODE como EH, POS_OWNER como socio y BU JUNIO como objetivo.")
        archivo_obj = st.file_uploader(
            "Subir/actualizar archivo de objetivos",
            type=["csv", "xlsx", "xls"],
            key=f"{key_prefix}_objetivos_uploader",
        )
        if archivo_obj is not None:
            try:
                df_obj_original = leer_archivo(archivo_obj)
                df_obj = normalizar_objetivos_archivo(df_obj_original)
                set_objetivos_df(df_obj)
                st.success(f"Objetivos cargados correctamente: {len(df_obj)} socios.")
            except Exception as exc:
                st.error(f"No se pudo cargar objetivos: {exc}")
        st.dataframe(get_objetivos_df(), use_container_width=True, hide_index=True)


def construir_detalle_comercial(df_filtro: pd.DataFrame) -> pd.DataFrame:
    objetivo_map, socio_obj_map, _ = objetivos_maps()
    detalle = pd.DataFrame({
        "Código cliente": df_filtro["_codigo"],
        "EH": df_filtro["_eh"],
        "Socio": df_filtro.apply(
            lambda r: r["_socio"] or socio_obj_map.get(str(r["_eh"]), SOCIOS_EH_DEFAULT.get(str(r["_eh"]), "SIN NOMBRE")),
            axis=1,
        ),
        "Tipo venta": df_filtro["_tipo_venta"],
        "Nodo": df_filtro["_nodo"],
        "Fecha": df_filtro["_fecha_base"],
        "Cliente": df_filtro["_cliente"],
        "Teléfono 1": df_filtro["_telefono1"],
        "Teléfono 2": df_filtro["_telefono2"],
    })
    detalle = detalle[detalle["Código cliente"].astype(str).str.strip().ne("")].copy()
    detalle["Es Crosselling"] = detalle["Tipo venta"].fillna("").astype(str).str.upper().str.contains("CROSS")
    detalle["Es GrossAdd"] = detalle["Tipo venta"].fillna("").astype(str).str.upper().str.contains("GROSS")
    # Regla solicitada: Crosselling NO suma al objetivo.
    detalle["Venta objetivo"] = ~detalle["Es Crosselling"]
    return detalle.reset_index(drop=True)


def construir_ranking_objetivos(detalle: pd.DataFrame) -> pd.DataFrame:
    objetivo_map, socio_obj_map, categoria_map = objetivos_maps()
    columnas = ["EH", "Socio", "Categoria", "Ventas objetivo", "Crosselling", "GrossAdd", "Total ventas", "Objetivo", "Cumplimiento %"]
    if detalle.empty:
        return pd.DataFrame(columns=columnas)

    ranking = (
        detalle.groupby(["EH", "Socio"], as_index=False)
        .agg(
            **{
                "Ventas objetivo": ("Venta objetivo", "sum"),
                "Crosselling": ("Es Crosselling", "sum"),
                "GrossAdd": ("Es GrossAdd", "sum"),
                "Total ventas": ("Código cliente", "count"),
            }
        )
    )
    ranking["EH"] = ranking["EH"].apply(limpiar_eh)
    ranking["Socio"] = ranking.apply(
        lambda r: r["Socio"] if str(r["Socio"]).strip() and str(r["Socio"]).upper() != "SIN NOMBRE" else socio_obj_map.get(r["EH"], "SIN NOMBRE"),
        axis=1,
    )
    ranking["Categoria"] = ranking["EH"].map(categoria_map).fillna("")
    ranking["Objetivo"] = ranking["EH"].map(objetivo_map).fillna(0).astype(int)
    ranking["Cumplimiento"] = ranking.apply(
        lambda r: (float(r["Ventas objetivo"]) / float(r["Objetivo"])) if float(r["Objetivo"] or 0) > 0 else 0,
        axis=1,
    )
    ranking["Cumplimiento %"] = (ranking["Cumplimiento"] * 100).round(1).astype(str) + "%"
    ranking = ranking.sort_values(["Ventas objetivo", "Crosselling", "Total ventas"], ascending=[False, False, False]).reset_index(drop=True)
    return ranking[columnas]


def nombre_corto(nombre: str) -> str:
    partes = str(nombre).split()
    return " ".join(partes[:2]) if len(partes) >= 2 else str(nombre)


def whatsapp_link(texto: str) -> str:
    return "https://wa.me/?text=" + quote(texto)


def formatear_fecha_mensaje(valor: object, con_anio: bool = True) -> str:
    """Devuelve fecha en formato dd/mm/yyyy para mensajes WhatsApp."""
    if valor is None or valor == "" or pd.isna(valor):
        return "S/F"
    try:
        fecha = pd.to_datetime(valor, errors="coerce")
        if pd.isna(fecha):
            return "S/F"
        return fecha.strftime("%d/%m/%Y" if con_anio else "%d/%m")
    except Exception:
        return "S/F"


def limpiar_nombre_socio_mensaje(nombre: object) -> str:
    texto = limpiar_texto(nombre).upper()
    texto = texto.replace("ZAMORA BUEZO,", "ZAMORA")
    texto = texto.replace("VILLAFUERTE GUERRA,", "VILLAFUERTE")
    texto = texto.replace("FERNANDEZ PUENTE,", "FERNANDEZ")
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto or "SIN NOMBRE"



def excel_bytes(hojas: dict[str, pd.DataFrame]) -> BytesIO:
    salida = BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        for nombre_hoja, df in hojas.items():
            safe_name = str(nombre_hoja)[:31] or "Hoja"
            df.to_excel(writer, index=False, sheet_name=safe_name)
        for hoja in writer.book.worksheets:
            hoja.freeze_panes = "A2"
            for col in hoja.columns:
                letra = col[0].column_letter
                ancho = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                hoja.column_dimensions[letra].width = min(max(ancho + 2, 10), 55)
    salida.seek(0)
    return salida


def extraer_nodo_texto(*valores: object) -> str:
    texto = " ".join(limpiar_texto(v).upper() for v in valores if limpiar_texto(v))
    if not texto:
        return ""
    texto = texto.replace("-", " ").replace("_", " ")
    texto = re.sub(r"\s+", " ", texto)
    patron = r"\b(EAL\s*\d{3,4}|LPZ\s*\d{3,4}|EAF\s*\d{3,4}|SCZ\s*\d{3,4}|PTS\s*\d{3,4}|[A-Z]{3}\s*\d{3,4})\b"
    encontrado = re.search(patron, texto)
    return encontrado.group(1).replace(" ", "") if encontrado else ""


# =========================================================
# MÓDULO: AGENDA TÉCNICA
# =========================================================
def preparar_agenda(df_original: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str | None]]:
    df = normalizar_dataframe(df_original)
    col_fecha = buscar_columna(df, [
        "inicio_agendado", "inicio cita", "inicio_cita", "fecha_agendada", "data_agendamiento", "fecha agenda",
    ])
    col_tipo = buscar_columna(df, [
        "tipo_trabajo_op", "tipo_trabajo", "tipo trabajo op", "tipo trabajo", "descripcion_trabajo",
    ])
    col_eh = buscar_columna(df, [
        "ehumano_promotor", "eh_promotor", "eh", "vendedor_eh", "codigo_eh", "promotor", "ehumano",
    ])
    columnas = {"fecha": col_fecha, "tipo_trabajo": col_tipo, "eh": col_eh}
    faltantes = [k for k, v in columnas.items() if v is None]
    if faltantes:
        raise ValueError(
            "No se encontraron columnas requeridas: " + ", ".join(faltantes) +
            ". Columnas del archivo: " + ", ".join(df.columns[:80])
        )

    df["_fecha_agendada"] = parsear_fecha(df[col_fecha])
    df["_tipo_op"] = df[col_tipo].fillna("").astype(str).str.upper().str.strip()
    df["_eh"] = df[col_eh].apply(limpiar_eh)
    return df, columnas


def procesar_agenda(df_original: pd.DataFrame, fecha_consulta: date, socios: dict[str, str], solo_mis_socios: bool):
    df, columnas = preparar_agenda(df_original)
    mask_fecha = df["_fecha_agendada"] == fecha_consulta
    mask_inst = df["_tipo_op"].str.contains("INSTALACION", na=False)
    mask_eh = df["_eh"].isin(socios.keys()) if solo_mis_socios else df["_eh"].ne("")
    filtrado = df[mask_fecha & mask_inst & mask_eh].copy()

    diagnostico = {
        "total_filas": int(len(df)),
        "filas_fecha": int(mask_fecha.sum()),
        "filas_instalacion_fecha": int((mask_fecha & mask_inst).sum()),
        "filas_eh_fecha_instalacion": int((mask_fecha & mask_inst & mask_eh).sum()),
        "fechas": sorted([f for f in df["_fecha_agendada"].dropna().unique()]),
        "columnas_detectadas": columnas,
    }

    columnas_detalle = [
        "EH", "Socio", "Código cliente", "Nodo", "Cliente", "Teléfono", "Ciudad", "Turno", "Hora inicio",
        "Contratista", "Técnico", "Estado", "Confirmación", "Dirección/Dato conexión",
    ]
    if filtrado.empty:
        detalle = pd.DataFrame(columns=columnas_detalle)
    else:
        detalle = pd.DataFrame({
            "EH": filtrado["_eh"],
            "Socio": filtrado["_eh"].map(socios).fillna("SIN NOMBRE CONFIGURADO"),
            "Código cliente": obtener_serie(filtrado, ["cliente_nro", "codigo_cliente", "cod_cliente", "cliente"]).apply(limpiar_codigo),
            "Nodo": filtrado.apply(lambda r: extraer_nodo_texto(
                r.get("dato_onexion", ""), r.get("dato_conexion", ""), r.get("zona_ramal", ""),
                r.get("tap_nap", ""), r.get("descripcion", ""), r.get("comentario", ""), r.get("territorio_servicio", ""),
            ), axis=1),
            "Cliente": obtener_serie(filtrado, ["nombre_contacto", "cliente_nombre", "nombre_cliente"]).apply(limpiar_texto),
            "Teléfono": obtener_serie(filtrado, ["numero_telefono_cliente", "cliente_telefono1", "cliente_telefono2", "telefono"]).apply(limpiar_texto),
            "Ciudad": obtener_serie(filtrado, ["ciudad"]).apply(limpiar_texto),
            "Turno": obtener_serie(filtrado, ["turno_agendamiento", "rango_cita_acordada_con_cliente", "turno"]).apply(limpiar_texto),
            "Hora inicio": obtener_serie(filtrado, ["hora_inicio", "inicio_cita"]).apply(limpiar_texto),
            "Contratista": obtener_serie(filtrado, ["contratista"]).apply(limpiar_texto),
            "Técnico": obtener_serie(filtrado, ["tecnico_nombre", "tecnico"]).apply(limpiar_texto),
            "Estado": obtener_serie(filtrado, ["estado"]).apply(limpiar_texto),
            "Confirmación": obtener_serie(filtrado, ["estado_confirmacion", "confirmacion"]).apply(limpiar_texto),
            "Dirección/Dato conexión": obtener_serie(filtrado, ["dato_onexion", "dato_conexion", "descripcion", "comentario"]).apply(limpiar_texto),
        })
        detalle = detalle.sort_values(["Socio", "Código cliente"], ascending=[True, True]).reset_index(drop=True)

    if detalle.empty:
        resumen = pd.DataFrame(columns=["EH", "Socio", "Instalaciones"])
    else:
        resumen = (
            detalle.groupby(["EH", "Socio"], as_index=False)["Código cliente"]
            .count()
            .rename(columns={"Código cliente": "Instalaciones"})
            .sort_values("Instalaciones", ascending=False)
            .reset_index(drop=True)
        )
    socios_df = pd.DataFrame({"EH": list(socios.keys()), "Socio": list(socios.values())})
    sin_instalacion = socios_df[~socios_df["EH"].isin(resumen["EH"] if not resumen.empty else [])].copy()
    sin_instalacion = sin_instalacion.sort_values("Socio").reset_index(drop=True)
    return resumen, detalle, sin_instalacion, diagnostico


def mensaje_agenda_global(detalle: pd.DataFrame, resumen: pd.DataFrame, fecha_consulta: date) -> str:
    total = len(detalle)
    fecha_txt = fecha_consulta.strftime("%d/%m/%Y")
    lineas = [
        "📌 *AGENDA TÉCNICA DEL DÍA*",
        "",
        f"📅 *Fecha:* {fecha_txt}",
        f"🔧 *Total:* {total} instalaciones",
        f"👥 *Socios con agenda:* {resumen['EH'].nunique() if not resumen.empty else 0}",
        "",
    ]
    if detalle.empty:
        lineas.append("⚠️ No se encontraron instalaciones programadas para los EH configurados.")
        return "\n".join(lineas)

    lineas.append("📋 *Detalle por socio:*")
    for (eh, socio), grupo in detalle.groupby(["EH", "Socio"], sort=False):
        lineas.extend([
            "",
            f"👤 *{socio}*",
            f"🆔 EH: *{eh}* | 🔧 *{len(grupo)}* inst.",
        ])
        for i, (_, r) in enumerate(grupo.iterrows(), start=1):
            codigo = r.get("Código cliente", "") or "S/D"
            nodo = r.get("Nodo", "") or "S/D"
            turno = r.get("Turno", "") or "S/D"
            lineas.append(f"{i}. {codigo} | {nodo} | {turno}")
    lineas.extend([
        "",
        "✅ Favor realizar seguimiento temprano, confirmar contacto con el cliente y reportar cualquier observación durante el día.",
    ])
    return "\n".join(lineas)


def mostrar_agenda_tecnica() -> None:
    st.title("📅 Agenda Técnica / Instalaciones de Hoy")
    st.caption("Carga el archivo BO-CITA SERVICIO NACIONAL. Filtra instalaciones por fecha y EH de tus socios.")

    with st.sidebar:
        st.subheader("Config. Agenda")
        solo_mis_socios = st.checkbox("Filtrar solo mis socios EH", value=True, key="agenda_solo_socios")
        with st.expander("Agregar o corregir EH", expanded=False):
            extra_eh = st.text_area("Formato: 59509 Nombre del socio", height=120, key="agenda_extra_eh")
        socios = construir_socios(extra_eh)

    archivo = st.file_uploader("📤 Subir archivo de agenda técnica", type=["xlsx", "xls", "csv"], key="agenda_uploader")
    if archivo is None:
        st.info("Sube el Excel diario de agenda técnica.")
        return

    try:
        hojas = obtener_hojas_excel(archivo)
        hoja = None
        if hojas and len(hojas) > 1:
            hoja = st.selectbox("Seleccionar hoja", hojas, key="agenda_hoja")
        elif hojas:
            hoja = hojas[0]
        archivo.seek(0)
        df_original = leer_archivo_con_hoja(archivo, hoja)
        df_preparado, _ = preparar_agenda(df_original)
        fechas = sorted([f for f in df_preparado["_fecha_agendada"].dropna().unique()])
    except Exception as exc:
        st.error(f"No se pudo leer/procesar el archivo: {exc}")
        return

    if not fechas:
        st.warning("No se encontraron fechas válidas en el archivo.")
        return

    hoy = hoy_bolivia()
    fecha_default = hoy if hoy in fechas else max(fechas)
    opciones = [f.strftime("%d/%m/%Y") for f in fechas]
    fecha_txt = st.selectbox("Fecha a consultar", opciones, index=fechas.index(fecha_default), key="agenda_fecha")
    fecha_consulta = fechas[opciones.index(fecha_txt)]

    resumen, detalle, sin_instalacion, diagnostico = procesar_agenda(df_original, fecha_consulta, socios, solo_mis_socios)

    k1, k2, k3 = st.columns(3)
    k1.metric("Instalaciones", len(detalle))
    k2.metric("Socios con agenda", resumen["EH"].nunique() if not resumen.empty else 0)
    k3.metric("Socios sin agenda", len(sin_instalacion) if solo_mis_socios else 0)

    with st.expander("Diagnóstico"):
        st.write(diagnostico)

    st.subheader("Resumen por socio")
    st.dataframe(resumen, use_container_width=True, hide_index=True)

    st.subheader("Detalle de instalaciones")
    st.dataframe(detalle, use_container_width=True, hide_index=True)

    st.subheader("📲 Mensaje global WhatsApp para grupo")
    texto = mensaje_agenda_global(detalle, resumen, fecha_consulta)
    st.text_area("Copiar mensaje", texto, height=380, key="agenda_msg_global")
    st.markdown(f"[📲 Enviar por WhatsApp]({whatsapp_link(texto)})")

    excel = excel_bytes({
        "Resumen": resumen,
        "Detalle": detalle,
        "Sin instalacion": sin_instalacion,
        "Diagnostico": pd.DataFrame([diagnostico]).astype(str),
    })
    st.download_button(
        "⬇️ Descargar Excel",
        data=excel,
        file_name=f"agenda_tecnica_{fecha_consulta.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="agenda_excel",
    )


# =========================================================
# MÓDULO: PENDIENTES INSTALACIÓN
# =========================================================
def preparar_pendientes(df_original: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str | None]]:
    df = normalizar_dataframe(df_original)
    col_codigo = buscar_columna(df, ["CLIENTE_NRO", "CODIGO_CLIENTE", "COD_CLIENTE", "CODIGO", "COD", "CLIENTE"])
    col_nodo = buscar_columna(df, ["NODO_NOMBRE", "NODO", "NODO_ACTUAL", "NODO RED", "NODO_RED"])
    col_fecha_ot = buscar_columna(df, ["FECHA_GENERACION_OT", "FECHA GENERACION OT", "FECHA_OT", "FECHA_VENTA", "FECHA_GENERACION"])
    col_fecha_reporte = buscar_columna(df, ["FECHA_REPORTE", "FECHA REPORTE", "FECHA_CARGA", "FECHA"])
    col_eh = buscar_columna(df, ["VENDEDOR_EH", "EH", "CODIGO_EH", "EHUMANO", "EH_PROMOTOR", "EJECUTIVO_EH"])
    col_socio = buscar_columna(df, ["VENDEDOR_NOMBRE", "NOMBRE_VENDEDOR", "SOCIO", "VENDEDOR", "NOMBRE_SOCIO"])
    col_tipo = buscar_columna(df, ["TIPO_VENTA", "TIPO", "TIPO_VTA", "TIPO_OPERACION"])
    col_cliente = buscar_columna(df, ["CLIENTE_NOMBRE", "NOMBRE_CLIENTE", "CLIENTE", "NOMBRE"])
    col_tel1 = buscar_columna(df, ["CLIENTE_TELEFONO1", "TELEFONO1", "TEL1", "TELEFONO", "CELULAR"])
    col_tel2 = buscar_columna(df, ["CLIENTE_TELEFONO2", "TELEFONO2", "TEL2", "REFERENCIA", "REF"])
    col_obs = buscar_columna(df, ["CRM_OBSERVACION", "OBSERVACION", "OBS", "COMENTARIO", "COMENTARIOS"])

    columnas = {
        "codigo": col_codigo, "nodo": col_nodo, "fecha_generacion_ot": col_fecha_ot, "fecha_reporte": col_fecha_reporte,
        "eh": col_eh, "socio": col_socio, "tipo_venta": col_tipo, "cliente": col_cliente,
        "telefono1": col_tel1, "telefono2": col_tel2, "observacion": col_obs,
    }
    faltantes = [k for k in ["codigo", "nodo"] if columnas[k] is None]
    if faltantes:
        raise ValueError("No se encontraron columnas obligatorias: " + ", ".join(faltantes))

    df["_codigo"] = df[col_codigo].apply(limpiar_codigo) if col_codigo else ""
    df["_nodo"] = df[col_nodo].apply(limpiar_texto).str.upper() if col_nodo else ""
    df["_eh"] = df[col_eh].apply(limpiar_eh) if col_eh else ""
    df["_socio"] = df[col_socio].apply(limpiar_texto) if col_socio else ""
    df["_tipo_venta"] = df[col_tipo].apply(limpiar_texto).str.upper() if col_tipo else ""
    df["_cliente"] = df[col_cliente].apply(limpiar_texto) if col_cliente else ""
    df["_telefono1"] = df[col_tel1].apply(limpiar_texto) if col_tel1 else ""
    df["_telefono2"] = df[col_tel2].apply(limpiar_texto) if col_tel2 else ""
    df["_observacion"] = df[col_obs].apply(limpiar_texto) if col_obs else ""

    if col_fecha_ot:
        df["_fecha_base"] = parsear_fecha(df[col_fecha_ot])
        columnas["fecha_base_usada"] = col_fecha_ot
    elif col_fecha_reporte:
        df["_fecha_base"] = parsear_fecha(df[col_fecha_reporte])
        columnas["fecha_base_usada"] = col_fecha_reporte
    else:
        df["_fecha_base"] = pd.NaT
        columnas["fecha_base_usada"] = None
    return df, columnas


def calcular_pendientes_antiguos(df_original: pd.DataFrame, dias_antiguedad: int, fecha_hoy: date, solo_eh_configurados: bool, socios: dict[str, str]):
    df, columnas = preparar_pendientes(df_original)
    fecha_corte = fecha_hoy - pd.Timedelta(days=int(dias_antiguedad))
    mask_codigo = df["_codigo"].ne("")
    if columnas.get("fecha_base_usada"):
        mask_fecha = df["_fecha_base"].notna() & (df["_fecha_base"] <= fecha_corte)
    else:
        mask_fecha = pd.Series([True] * len(df), index=df.index)
    mask_eh = df["_eh"].isin(socios.keys()) if solo_eh_configurados else pd.Series([True] * len(df), index=df.index)
    antiguos = df[mask_codigo & mask_fecha & mask_eh].copy()

    if antiguos.empty:
        detalle = pd.DataFrame(columns=["Código cliente", "Nodo", "Fecha", "Días", "EH", "Socio", "Tipo venta", "Cliente", "Teléfono 1", "Teléfono 2", "Observación"])
    else:
        fecha_base = pd.to_datetime(antiguos["_fecha_base"], errors="coerce")
        dias = (pd.to_datetime(fecha_hoy) - fecha_base).dt.days
        detalle = pd.DataFrame({
            "Código cliente": antiguos["_codigo"],
            "Nodo": antiguos["_nodo"],
            "Fecha": antiguos["_fecha_base"],
            "Días": dias,
            "EH": antiguos["_eh"],
            "Socio": antiguos.apply(lambda r: r["_socio"] or socios.get(r["_eh"], "SIN NOMBRE"), axis=1),
            "Tipo venta": antiguos["_tipo_venta"],
            "Cliente": antiguos["_cliente"],
            "Teléfono 1": antiguos["_telefono1"],
            "Teléfono 2": antiguos["_telefono2"],
            "Observación": antiguos["_observacion"],
        })
        detalle = detalle.sort_values(["Días", "Fecha", "Socio"], ascending=[False, True, True]).reset_index(drop=True)

    resumen = pd.DataFrame(columns=["EH", "Socio", "Pendientes antiguos"]) if detalle.empty else (
        detalle.groupby(["EH", "Socio"], as_index=False)["Código cliente"].count()
        .rename(columns={"Código cliente": "Pendientes antiguos"})
        .sort_values("Pendientes antiguos", ascending=False)
        .reset_index(drop=True)
    )
    diagnostico = {"columnas": columnas, "total_filas": len(df), "fecha_corte": fecha_corte, "fecha_hoy": fecha_hoy, "total_antiguos": len(detalle)}
    return resumen, detalle, diagnostico


def mensaje_pendientes_antiguos(detalle: pd.DataFrame, dias_antiguedad: int, fecha_hoy: date, agrupado: bool = True) -> str:
    """Mensaje de pendientes de instalación: solo código, nodo y fecha."""
    fecha_txt = fecha_hoy.strftime("%d/%m/%Y")
    lineas = [
        f"⏳ *PENDIENTES DE INSTALACIÓN +{dias_antiguedad} DÍAS*",
        f"📅 Fecha revisión: *{fecha_txt}*",
        f"🔢 Total casos: *{len(detalle)}*",
        "",
    ]
    if detalle.empty:
        lineas.append("✅ No se encontraron pendientes antiguos con el filtro seleccionado.")
        return "\n".join(lineas)

    if agrupado:
        lineas.append("📋 *Detalle por socio:*")
        for (eh, socio), grupo in detalle.groupby(["EH", "Socio"], dropna=False, sort=False):
            lineas.append("")
            lineas.append(f"👤 *{limpiar_nombre_socio_mensaje(socio)}*")
            lineas.append(f"🆔 EH: *{eh or 'S/D'}* | 📌 *{len(grupo)}* casos")
            for i, (_, r) in enumerate(grupo.iterrows(), start=1):
                codigo = r.get("Código cliente", "") or "S/D"
                nodo = r.get("Nodo", "") or "S/D"
                fecha = formatear_fecha_mensaje(r.get("Fecha", ""))
                lineas.append(f"{i}. {codigo} | {nodo} | {fecha}")
    else:
        lineas.append("📋 *Código | Nodo | Fecha*")
        for i, (_, r) in enumerate(detalle.iterrows(), start=1):
            codigo = r.get("Código cliente", "") or "S/D"
            nodo = r.get("Nodo", "") or "S/D"
            fecha = formatear_fecha_mensaje(r.get("Fecha", ""))
            lineas.append(f"{i}. {codigo} | {nodo} | {fecha}")

    lineas.append("")
    lineas.append("✅ Favor priorizar estos códigos por antigüedad y reportar avance durante el día.")
    return "\n".join(lineas).strip()


def mostrar_pendientes_inst() -> None:
    st.title("📋 Pendientes de Instalación con/Sin Pago")
    st.caption("Carga PENDIENTE_INST_CON_PAGO o PENDIENTE_INST_SIN_PAGO. Filtra pendientes antiguos por socio y genera WhatsApp con código, nodo y fecha.")

    with st.sidebar:
        st.subheader("Config. Pendientes")
        dias_antiguedad = st.number_input("Antigüedad mínima en días", min_value=1, max_value=30, value=3, step=1, key="pend_dias")
        solo_eh_configurados = st.checkbox("Solo mis socios EH", value=False, key="pend_solo_eh")
        with st.expander("Agregar/corregir EH", expanded=False):
            texto_eh_extra = st.text_area("Formato: 59509 Nombre del socio", height=120, key="pend_extra_eh")
        socios = construir_socios(texto_eh_extra)

    archivo = st.file_uploader("📤 Sube archivo Pendientes de Instalación CON PAGO o SIN PAGO", type=["csv", "xlsx", "xls"], key="pend_uploader")
    if archivo is None:
        st.info("Sube el archivo PENDIENTE_INST_CON_PAGO o PENDIENTE_INST_SIN_PAGO para generar el reporte.")
        return

    try:
        df_original = leer_archivo(archivo)
        st.success(f"Archivo leído correctamente: {len(df_original)} filas y {len(df_original.columns)} columnas.")
        fecha_hoy = hoy_bolivia()
        resumen, detalle, diagnostico = calcular_pendientes_antiguos(df_original, int(dias_antiguedad), fecha_hoy, solo_eh_configurados, socios)
    except Exception as exc:
        st.error(f"No se pudo procesar el archivo: {exc}")
        try:
            archivo.seek(0)
            st.write("Columnas detectadas:", list(normalizar_dataframe(leer_archivo(archivo)).columns))
        except Exception:
            pass
        return

    k1, k2, k3 = st.columns(3)
    k1.metric("Pendientes antiguos", len(detalle))
    k2.metric("Socios/EH", resumen["EH"].nunique() if not resumen.empty else 0)
    k3.metric("Fecha corte", pd.to_datetime(diagnostico["fecha_corte"]).strftime("%d/%m/%Y"))

    with st.expander("Diagnóstico"):
        st.json({k: str(v) for k, v in diagnostico.items()})

    if detalle.empty:
        st.warning("No se encontraron casos antiguos con el filtro actual.")
        return

    st.subheader("📊 Resumen por socio")
    st.dataframe(resumen, use_container_width=True, hide_index=True)

    col_socio, col_nodo = st.columns(2)
    socios_opciones = ["Todos"] + sorted([s for s in detalle["Socio"].dropna().unique() if str(s).strip()])
    nodos_opciones = ["Todos"] + sorted([n for n in detalle["Nodo"].dropna().unique() if str(n).strip()])
    socio_sel = col_socio.selectbox("Filtrar por socio", socios_opciones, key="pend_fil_socio")
    nodo_sel = col_nodo.selectbox("Filtrar por nodo", nodos_opciones, key="pend_fil_nodo")

    detalle_filtrado = detalle.copy()
    if socio_sel != "Todos":
        detalle_filtrado = detalle_filtrado[detalle_filtrado["Socio"] == socio_sel]
    if nodo_sel != "Todos":
        detalle_filtrado = detalle_filtrado[detalle_filtrado["Nodo"] == nodo_sel]

    st.subheader(f"⏳ Casos antiguos +{int(dias_antiguedad)} días")
    columnas_vista = ["Código cliente", "Nodo", "Fecha", "Días", "EH", "Socio"]
    st.dataframe(detalle_filtrado[columnas_vista], use_container_width=True, hide_index=True)

    st.subheader("📲 Mensaje WhatsApp")
    formato = st.radio("Formato", ["Agrupado por socio", "Código + nodo"], horizontal=True, key="pend_formato")
    texto = mensaje_pendientes_antiguos(detalle_filtrado, int(dias_antiguedad), fecha_hoy, agrupado=(formato == "Agrupado por socio"))
    st.text_area("Copiar mensaje", texto, height=360, key="pend_msg")
    st.markdown(f"[📲 Enviar por WhatsApp]({whatsapp_link(texto)})")

    excel = excel_bytes({"Resumen": resumen, "Pendientes +3 dias": detalle_filtrado[columnas_vista], "Diagnostico": pd.DataFrame([diagnostico]).astype(str)})
    st.download_button("⬇️ Descargar Excel", excel, file_name=f"pendientes_inst_antiguos_{fecha_hoy.strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="pend_excel")

# =========================================================
# MÓDULO: SUSPENDIDAS
# =========================================================
def clasificar_observacion(observacion: object) -> str:
    txt = sin_tildes(observacion).upper().strip()
    if not txt:
        return "Sin CRM/observación"
    reglas = [
        ("TAP saturado", ["TAP", "SATUR"]),
        ("Volver a llamar", ["VOLVER A LLAMAR", "VOLVER", "LLAMAR LUEGO", "LLAMAR MAS TARDE"]),
        ("Sin contacto", ["NO CONTESTA", "NO RESPONDE", "SIN CONTACTO", "NO CONTESTO", "NO ATIEND", "APAGADO"]),
        ("Cliente ausente", ["CLIENTE AUSENTE", "NO SE ENCUENTRA", "NO ESTA", "AUSENTE", "NO HABIA"]),
        ("Reagendar", ["REAGEND", "REPROGRAM", "AGENDA", "AGENDAR"]),
        ("Bloqueo/acceso", ["BLOQUEO", "PARO", "TRANSITO", "ACCESO", "NO HAY PASO", "MOVILIDAD"]),
        ("Validar dirección", ["DIRECCION", "DOMICILIO", "REFERENCIA", "UBICACION", "ZONA", "NO UBICA"]),
        ("Cliente desiste", ["DESIST", "ANULA", "ANUL", "NO QUIERE", "CANCELO", "CANCELA"]),
        ("Pago/deuda", ["PAGO", "DEUDA", "MORA", "COBRANZA", "PENDIENTE DE PAGO"]),
        ("Pendiente técnico", ["TECNICO", "CONTRATISTA", "MATERIAL", "SEÑAL", "SENAL", "NAP", "FACTIBILIDAD"]),
    ]
    for categoria, claves in reglas:
        if any(clave in txt for clave in claves):
            return categoria
    return "Revisar observación"


def preparar_suspendidas(df_original: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str | None]]:
    df = normalizar_dataframe(df_original)
    col_codigo = buscar_columna(df, ["CLIENTE_NRO", "CODIGO_CLIENTE", "COD_CLIENTE", "CODIGO", "COD", "CLIENTE"])
    col_nodo = buscar_columna(df, ["NODO_NOMBRE", "NODO", "NODO_ACTUAL", "NODO RED", "NODO_RED"])
    col_fecha_reporte = buscar_columna(df, ["FECHA_REPORTE", "FECHA REPORTE", "FECHA_CARGA", "FECHA", "FECHA_SUSPENDIDA", "FECHA_ESTADO"])
    col_fecha_ot = buscar_columna(df, ["FECHA_GENERACION_OT", "FECHA GENERACION OT", "FECHA_OT", "FECHA_VENTA", "FECHA_GENERACION"])
    col_eh = buscar_columna(df, ["VENDEDOR_EH", "EH", "CODIGO_EH", "EHUMANO", "EH_PROMOTOR", "EJECUTIVO_EH"])
    col_socio = buscar_columna(df, ["VENDEDOR_NOMBRE", "NOMBRE_VENDEDOR", "SOCIO", "VENDEDOR", "NOMBRE_SOCIO"])
    col_tipo = buscar_columna(df, ["TIPO_VENTA", "TIPO", "TIPO_VTA", "TIPO_OPERACION"])
    col_cliente = buscar_columna(df, ["CLIENTE_NOMBRE", "NOMBRE_CLIENTE", "CLIENTE", "NOMBRE"])
    col_tel1 = buscar_columna(df, ["CLIENTE_TELEFONO1", "TELEFONO1", "TEL1", "TELEFONO", "CELULAR"])
    col_tel2 = buscar_columna(df, ["CLIENTE_TELEFONO2", "TELEFONO2", "TEL2", "REFERENCIA", "REF"])
    col_obs = buscar_columna(df, ["CRM_OBSERVACION", "OBSERVACION", "OBS", "COMENTARIO", "COMENTARIOS", "MOTIVO", "MOTIVO_SUSPENSION"])
    col_estado = buscar_columna(df, ["ESTADO", "ESTADO_OT", "ESTADO_ACTUAL", "ESTADO_TRABAJO"])
    columnas = {
        "codigo": col_codigo, "nodo": col_nodo, "fecha_reporte": col_fecha_reporte, "fecha_generacion_ot": col_fecha_ot,
        "eh": col_eh, "socio": col_socio, "tipo_venta": col_tipo, "cliente": col_cliente,
        "telefono1": col_tel1, "telefono2": col_tel2, "observacion": col_obs, "estado": col_estado,
    }
    faltantes = [k for k in ["codigo", "nodo"] if columnas[k] is None]
    if faltantes:
        raise ValueError("No se encontraron columnas obligatorias: " + ", ".join(faltantes))

    df["_codigo"] = df[col_codigo].apply(limpiar_codigo) if col_codigo else ""
    df["_nodo"] = df[col_nodo].apply(limpiar_texto).str.upper() if col_nodo else ""
    df["_eh"] = df[col_eh].apply(limpiar_eh) if col_eh else ""
    df["_socio"] = df[col_socio].apply(limpiar_texto) if col_socio else ""
    df["_tipo_venta"] = df[col_tipo].apply(limpiar_texto).str.upper() if col_tipo else ""
    df["_cliente"] = df[col_cliente].apply(limpiar_texto) if col_cliente else ""
    df["_telefono1"] = df[col_tel1].apply(limpiar_texto) if col_tel1 else ""
    df["_telefono2"] = df[col_tel2].apply(limpiar_texto) if col_tel2 else ""
    df["_observacion"] = df[col_obs].apply(limpiar_texto) if col_obs else ""
    df["_estado"] = df[col_estado].apply(limpiar_texto) if col_estado else ""
    df["_categoria"] = df["_observacion"].apply(clasificar_observacion)

    if col_fecha_reporte:
        df["_fecha_base"] = parsear_fecha(df[col_fecha_reporte])
        columnas["fecha_base_usada"] = col_fecha_reporte
    elif col_fecha_ot:
        df["_fecha_base"] = parsear_fecha(df[col_fecha_ot])
        columnas["fecha_base_usada"] = col_fecha_ot
    else:
        df["_fecha_base"] = pd.NaT
        columnas["fecha_base_usada"] = None
    return df, columnas


def calcular_suspendidas_antiguas(df_original: pd.DataFrame, dias_antiguedad: int, fecha_hoy: date, solo_eh_configurados: bool, socios: dict[str, str], tipo_venta: str, categoria: str):
    df, columnas = preparar_suspendidas(df_original)
    fecha_corte = fecha_hoy - pd.Timedelta(days=int(dias_antiguedad))
    mask_codigo = df["_codigo"].ne("")
    mask_fecha = df["_fecha_base"].notna() & (df["_fecha_base"] <= fecha_corte) if columnas.get("fecha_base_usada") else pd.Series([True] * len(df), index=df.index)
    mask_eh = df["_eh"].isin(socios.keys()) if solo_eh_configurados else pd.Series([True] * len(df), index=df.index)
    mask_tipo = pd.Series([True] * len(df), index=df.index) if tipo_venta == "Todos" else df["_tipo_venta"].str.contains(tipo_venta.upper(), na=False)
    mask_categoria = pd.Series([True] * len(df), index=df.index) if categoria == "Todas" else df["_categoria"].eq(categoria)
    suspendidas = df[mask_codigo & mask_fecha & mask_eh & mask_tipo & mask_categoria].copy()

    if suspendidas.empty:
        detalle = pd.DataFrame(columns=["Código cliente", "Nodo", "Fecha", "Días", "EH", "Socio", "Tipo venta", "Categoría", "Cliente", "Teléfono 1", "Teléfono 2", "Estado", "Observación"])
    else:
        detalle = pd.DataFrame({
            "Código cliente": suspendidas["_codigo"],
            "Nodo": suspendidas["_nodo"],
            "Fecha": suspendidas["_fecha_base"],
            "Días": suspendidas["_fecha_base"].apply(lambda f: (fecha_hoy - f).days if pd.notna(f) else ""),
            "EH": suspendidas["_eh"],
            "Socio": suspendidas.apply(lambda r: r["_socio"] or socios.get(r["_eh"], "SIN NOMBRE CONFIGURADO"), axis=1),
            "Tipo venta": suspendidas["_tipo_venta"],
            "Categoría": suspendidas["_categoria"],
            "Cliente": suspendidas["_cliente"],
            "Teléfono 1": suspendidas["_telefono1"],
            "Teléfono 2": suspendidas["_telefono2"],
            "Estado": suspendidas["_estado"],
            "Observación": suspendidas["_observacion"],
        })
        detalle = detalle.sort_values(["Días", "Socio", "Código cliente"], ascending=[False, True, True]).reset_index(drop=True)

    resumen = pd.DataFrame(columns=["EH", "Socio", "Suspendidas"]) if detalle.empty else (
        detalle.groupby(["EH", "Socio"], as_index=False)["Código cliente"].count()
        .rename(columns={"Código cliente": "Suspendidas"})
        .sort_values("Suspendidas", ascending=False)
        .reset_index(drop=True)
    )
    resumen_categoria = pd.DataFrame(columns=["Categoría", "Cantidad"]) if detalle.empty else (
        detalle.groupby("Categoría", as_index=False)["Código cliente"].count()
        .rename(columns={"Código cliente": "Cantidad"})
        .sort_values("Cantidad", ascending=False)
        .reset_index(drop=True)
    )
    diagnostico = {
        "total_filas": len(df), "total_suspendidas_antiguas": len(detalle), "fecha_corte": fecha_corte,
        "columnas_detectadas": columnas, "categorias": sorted(df["_categoria"].dropna().unique().tolist()),
    }
    return resumen, resumen_categoria, detalle, diagnostico


def mensaje_suspendidas_global(detalle: pd.DataFrame, fecha_hoy: date, dias_antiguedad: int, incluir_motivo: bool = False) -> str:
    """Mensaje de suspendidas agrupado por socio y reducido."""
    fecha_txt = fecha_hoy.strftime("%d/%m/%Y")
    lineas = [
        f"🚨 *SUSPENDIDAS +{dias_antiguedad} DÍAS*",
        f"📅 Corte: *{fecha_txt}*",
        f"📌 Total casos: *{len(detalle)}*",
        "",
    ]
    if detalle.empty:
        lineas.append("✅ No se encontraron suspendidas antiguas con los filtros seleccionados.")
        return "\n".join(lineas)
    lineas.append("📋 *Detalle por socio:*")
    for (eh, socio), grupo in detalle.groupby(["EH", "Socio"], dropna=False, sort=False):
        lineas.append("")
        lineas.append(f"👤 *{limpiar_nombre_socio_mensaje(socio)}*")
        lineas.append(f"🆔 EH: *{eh or 'S/D'}* | 🚨 *{len(grupo)}* casos")
        for i, (_, r) in enumerate(grupo.iterrows(), start=1):
            base = f"{i}. {r.get('Código cliente','S/D')} | {r.get('Nodo','S/D')} | {formatear_fecha_mensaje(r.get('Fecha',''))}"
            if incluir_motivo:
                base += f" | {r.get('Categoría','Revisar')}"
            lineas.append(base)
    lineas.append("")
    lineas.append("✅ Favor revisar, contactar al cliente y reportar avance.")
    return "\n".join(lineas)


def mostrar_suspendidas() -> None:
    st.title("🚨 Suspendidas")
    st.caption("Muestra suspendidas antiguas por socio. Mensaje WhatsApp reducido: código, nodo y fecha.")

    with st.sidebar:
        st.subheader("Config. Suspendidas")
        dias_antiguedad = st.number_input("Antigüedad mínima en días", min_value=1, max_value=30, value=3, step=1, key="susp_dias")
        fecha_hoy = st.date_input("Fecha de corte", value=hoy_bolivia(), format="DD/MM/YYYY", key="susp_fecha")
        solo_eh_configurados = st.checkbox("Solo mis socios EH", value=True, key="susp_solo_eh")
        tipo_venta = st.selectbox("Tipo de venta", ["Todos", "GROSSADD", "CROSS_SELLING"], key="susp_tipo")
        with st.expander("Agregar o corregir EH", expanded=False):
            extra_eh = st.text_area("Formato: 59509 Nombre del socio", height=120, key="susp_extra_eh")
        socios = construir_socios(extra_eh)

    archivo = st.file_uploader("📤 Sube archivo Suspendidas", type=["csv", "xlsx", "xls"], key="susp_uploader")
    if archivo is None:
        st.info("Sube el archivo de suspendidas para generar el reporte.")
        return

    try:
        df_original = leer_archivo(archivo)
        df_preparado, _ = preparar_suspendidas(df_original)
        categorias = ["Todas"] + sorted(df_preparado["_categoria"].dropna().unique().tolist())
        categoria = st.selectbox("Filtrar por categoría CRM", categorias, key="susp_categoria")
        resumen, resumen_categoria, detalle, diagnostico = calcular_suspendidas_antiguas(
            df_original, int(dias_antiguedad), fecha_hoy, solo_eh_configurados, socios, tipo_venta, categoria
        )
    except Exception as exc:
        st.error(f"No se pudo procesar el archivo: {exc}")
        return

    k1, k2, k3 = st.columns(3)
    k1.metric("Suspendidas antiguas", len(detalle))
    k2.metric("Socios con casos", resumen["EH"].nunique() if not resumen.empty else 0)
    k3.metric("Categorías", resumen_categoria["Categoría"].nunique() if not resumen_categoria.empty else 0)

    with st.expander("Diagnóstico"):
        st.json({k: str(v) for k, v in diagnostico.items()})

    st.subheader("📊 Resumen por socio")
    st.dataframe(resumen, use_container_width=True, hide_index=True)

    st.subheader("📊 Resumen por categoría CRM")
    st.dataframe(resumen_categoria, use_container_width=True, hide_index=True)

    if detalle.empty:
        st.warning("No se encontraron suspendidas antiguas con el filtro actual.")
        return

    col_socio, col_nodo = st.columns(2)
    socios_opciones = ["Todos"] + sorted([s for s in detalle["Socio"].dropna().unique() if str(s).strip()])
    nodos_opciones = ["Todos"] + sorted([n for n in detalle["Nodo"].dropna().unique() if str(n).strip()])
    socio_sel = col_socio.selectbox("Filtrar por socio", socios_opciones, key="susp_fil_socio")
    nodo_sel = col_nodo.selectbox("Filtrar por nodo", nodos_opciones, key="susp_fil_nodo")

    detalle_filtrado = detalle.copy()
    if socio_sel != "Todos":
        detalle_filtrado = detalle_filtrado[detalle_filtrado["Socio"] == socio_sel]
    if nodo_sel != "Todos":
        detalle_filtrado = detalle_filtrado[detalle_filtrado["Nodo"] == nodo_sel]

    st.subheader("🚨 Detalle suspendidas antiguas")
    columnas_vista = ["Código cliente", "Nodo", "Fecha", "Días", "EH", "Socio", "Categoría"]
    st.dataframe(detalle_filtrado[columnas_vista], use_container_width=True, hide_index=True)

    st.subheader("📲 Mensaje global WhatsApp")
    incluir_motivo = st.checkbox("Incluir motivo/categoría en el mensaje", value=False, key="susp_incluir_motivo")
    texto = mensaje_suspendidas_global(detalle_filtrado, fecha_hoy, int(dias_antiguedad), incluir_motivo=incluir_motivo)
    st.text_area("Copiar mensaje global", texto, height=360, key="susp_msg")
    st.markdown(f"[📲 Enviar por WhatsApp]({whatsapp_link(texto)})")

    excel = excel_bytes({"Resumen por socio": resumen, "Resumen categoria": resumen_categoria, "Detalle": detalle_filtrado[columnas_vista], "Diagnostico": pd.DataFrame([diagnostico]).astype(str)})
    st.download_button("⬇️ Descargar Excel", excel, file_name=f"suspendidas_antiguas_{fecha_hoy.strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="susp_excel")

# =========================================================
# MÓDULO: DASHBOARD / RANKING
# =========================================================
def _preparar_base_comercial(df_original: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str | None]]:
    df = normalizar_dataframe(df_original)

    col_codigo = buscar_columna(df, ["CLIENTE_NRO", "CODIGO_CLIENTE", "COD_CLIENTE", "CODIGO", "COD", "CLIENTE"])
    col_fecha_inst = buscar_columna(df, ["FECHA_INSTALACION", "FECHA INSTALACION", "FECHA_COMPLETADO", "FECHA_CIERRE"])
    col_fecha_gen = buscar_columna(df, ["FECHA_GENERACION_OT", "FECHA_GENERACION", "FECHA_REPORTE", "FECHA", "FECHA_VENTA"])
    col_eh = buscar_columna(df, ["VENDEDOR_EH", "EH", "CODIGO_EH", "EHUMANO", "EH_PROMOTOR", "EJECUTIVO_EH"])
    col_socio = buscar_columna(df, ["VENDEDOR_NOMBRE", "NOMBRE_VENDEDOR", "SOCIO", "VENDEDOR", "NOMBRE_SOCIO"])
    col_tipo = buscar_columna(df, ["TIPO_VENTA", "TIPO", "TIPO_VTA", "TIPO_OPERACION"])
    col_nodo = buscar_columna(df, ["NODO_NOMBRE", "NODO", "NODO_ACTUAL", "NODO_RED"])
    col_cliente = buscar_columna(df, ["CLIENTE_NOMBRE", "NOMBRE_CLIENTE", "CLIENTE", "NOMBRE"])
    col_tel1 = buscar_columna(df, ["CLIENTE_TELEFONO1", "TELEFONO1", "TEL1", "TELEFONO", "CELULAR"])
    col_tel2 = buscar_columna(df, ["CLIENTE_TELEFONO2", "TELEFONO2", "TEL2", "REFERENCIA", "REF"])

    columnas = {
        "codigo": col_codigo,
        "fecha_instalacion": col_fecha_inst,
        "fecha_base": col_fecha_inst or col_fecha_gen,
        "fecha_generacion": col_fecha_gen,
        "eh": col_eh,
        "socio": col_socio,
        "tipo_venta": col_tipo,
        "nodo": col_nodo,
        "cliente": col_cliente,
        "telefono1": col_tel1,
        "telefono2": col_tel2,
    }

    df["_codigo"] = df[col_codigo].apply(limpiar_codigo) if col_codigo else ""
    df["_eh"] = df[col_eh].apply(limpiar_eh) if col_eh else ""
    df["_socio"] = df[col_socio].apply(limpiar_texto) if col_socio else ""
    df["_tipo_venta"] = df[col_tipo].apply(limpiar_texto).str.upper() if col_tipo else ""
    df["_nodo"] = df[col_nodo].apply(limpiar_texto).str.upper() if col_nodo else ""
    df["_cliente"] = df[col_cliente].apply(limpiar_texto) if col_cliente else ""
    df["_telefono1"] = df[col_tel1].apply(limpiar_texto) if col_tel1 else ""
    df["_telefono2"] = df[col_tel2].apply(limpiar_texto) if col_tel2 else ""

    if columnas["fecha_base"]:
        df["_fecha_base"] = parsear_fecha(df[columnas["fecha_base"]])
    else:
        df["_fecha_base"] = pd.NaT

    if col_fecha_inst:
        df["_fecha_instalacion"] = parsear_fecha(df[col_fecha_inst])
    else:
        df["_fecha_instalacion"] = pd.NaT

    return df, columnas


def mostrar_dashboard() -> None:
    st.title("📊 Dashboard de Ventas")
    st.caption("Ranking por socio con regla: Crosselling no suma al objetivo.")

    cargar_objetivos_widget("dash")

    archivo = st.file_uploader("📤 Subir archivo de ventas / GrossAdd", type=["csv", "xlsx", "xls"], key="dash_uploader")
    if archivo is None:
        st.info("Sube un archivo de ventas para generar el dashboard.")
        return

    try:
        df_original = leer_archivo(archivo)
        df, columnas = _preparar_base_comercial(df_original)
    except Exception as exc:
        st.error(f"No se pudo procesar el archivo: {exc}")
        return

    st.success(f"Archivo leído correctamente: {len(df_original)} filas y {len(df_original.columns)} columnas.")
    with st.expander("Diagnóstico de columnas", expanded=False):
        st.json(columnas)

    df_filtro = df.copy()
    fechas_validas = sorted([f for f in df_filtro["_fecha_base"].dropna().unique()])
    if fechas_validas:
        c1, c2 = st.columns(2)
        desde = c1.date_input("Desde", value=min(fechas_validas), format="DD/MM/YYYY", key="dash_desde")
        hasta = c2.date_input("Hasta", value=max(fechas_validas), format="DD/MM/YYYY", key="dash_hasta")
        df_filtro = df_filtro[df_filtro["_fecha_base"].between(desde, hasta)]

    if "_eh" in df_filtro.columns and df_filtro["_eh"].ne("").any():
        opciones_eh = ["Todos"] + sorted(df_filtro["_eh"].dropna().astype(str).unique().tolist())
        eh_sel = st.selectbox("Filtrar EH", opciones_eh, key="dash_eh")
        if eh_sel != "Todos":
            df_filtro = df_filtro[df_filtro["_eh"].astype(str) == eh_sel]

    detalle = construir_detalle_comercial(df_filtro)
    ranking = construir_ranking_objetivos(detalle)

    total = len(detalle)
    ventas_obj = int(detalle["Venta objetivo"].sum()) if not detalle.empty else 0
    cross = int(detalle["Es Crosselling"].sum()) if not detalle.empty else 0
    gross = int(detalle["Es GrossAdd"].sum()) if not detalle.empty else 0
    objetivo_total = int(ranking["Objetivo"].sum()) if not ranking.empty else 0
    cumplimiento = (ventas_obj / objetivo_total * 100) if objetivo_total > 0 else 0

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total ventas", total)
    k2.metric("Ventas objetivo", ventas_obj)
    k3.metric("Crosselling", cross)
    k4.metric("Objetivo", objetivo_total)
    k5.metric("Cumplimiento", f"{cumplimiento:.1f}%")

    if detalle.empty:
        st.warning("No hay datos con los filtros seleccionados.")
        return

    st.subheader("🏆 Ranking por socio")
    st.caption("Ventas objetivo = total de ventas sin Crosselling. Crosselling se muestra separado y no suma al objetivo.")
    st.dataframe(ranking, use_container_width=True, hide_index=True)

    st.subheader("📋 Detalle")
    st.dataframe(detalle.drop(columns=["Venta objetivo", "Es Crosselling", "Es GrossAdd"], errors="ignore"), use_container_width=True, hide_index=True)

    lineas = [
        "📊 *RESUMEN DE VENTAS*",
        f"🔢 Total ventas: *{total}*",
        f"🎯 Ventas nuevas objetivo: *{ventas_obj}*",
        f"🔁 Crosselling: *{cross}* _(no suma al objetivo)_",
        f"📌 Objetivo ventas nuevas: *{objetivo_total}*",
        f"📈 Cumplimiento: *{cumplimiento:.1f}%*",
        "",
        "🏆 *Ranking por socio:*",
    ]
    for i, (_, r) in enumerate(ranking.iterrows(), start=1):
        icono = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "🔹"
        lineas.append(
            f"{icono} {r['EH']} - {nombre_corto(r['Socio'])}: "
            f"*{int(r['Ventas objetivo'])}/{int(r['Objetivo'])}* ({r['Cumplimiento %']}) | Cross: *{int(r['Crosselling'])}*"
        )
    lineas += ["", "💪 *Mensaje:* vamos equipo, prioricemos ventas nuevas para subir el cumplimiento. Crosselling se informa separado."]
    texto = "\n".join(lineas)

    st.subheader("📲 Mensaje WhatsApp")
    st.text_area("Copiar mensaje", texto, height=330, key="dash_msg")
    st.markdown(f"[📲 Enviar por WhatsApp]({whatsapp_link(texto)})")

    excel = excel_bytes({"Ranking": ranking, "Detalle": detalle, "Objetivos": get_objetivos_df()})
    st.download_button(
        "⬇️ Descargar Excel",
        data=excel,
        file_name=f"dashboard_ventas_{hoy_bolivia().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dash_excel",
    )


# =========================================================
# MÓDULO: WHATSAPP
# =========================================================
def _fecha_guion(valor: object) -> str:
    """Fecha en formato dd-mm-aaaa para el mensaje comercial."""
    if valor is None or valor == "" or pd.isna(valor):
        return "S/F"
    try:
        fecha = pd.to_datetime(valor, errors="coerce")
        if pd.isna(fecha):
            return "S/F"
        return fecha.strftime("%d-%m-%Y")
    except Exception:
        return "S/F"


def _mensaje_motivacional(ventas_obj: int, objetivo: int) -> str:
    if objetivo <= 0:
        return "💪 Revisemos el objetivo asignado y sigamos avanzando con el seguimiento."
    cumplimiento = (ventas_obj / objetivo) * 100
    if cumplimiento >= 100:
        return "🏆 Excelente trabajo, objetivo cumplido. Sigamos sumando ventas nuevas."
    if cumplimiento >= 80:
        return "🔥 Estamos cerca del objetivo. Prioricemos los cierres pendientes."
    if cumplimiento >= 50:
        return "💪 Buen avance. Sigamos empujando las ventas nuevas para llegar al objetivo."
    return "💪 Sigamos avanzando hacia el objetivo del mes."


def _filas_codigos_objetivo(detalle_socio: pd.DataFrame) -> list[str]:
    codigos_obj = detalle_socio[detalle_socio["Venta objetivo"] == True].copy()
    codigos_obj = codigos_obj.sort_values(["Fecha", "Código cliente"], ascending=[True, True])
    lineas: list[str] = []
    for _, r in codigos_obj.iterrows():
        codigo = r.get("Código cliente", "") or "S/D"
        cliente = limpiar_texto(r.get("Cliente", "")) or "SIN NOMBRE"
        fecha = _fecha_guion(r.get("Fecha", ""))
        lineas.append(f"🔹 {codigo} | {cliente} | {fecha}")
    return lineas


def _filas_codigos_cross(detalle_socio: pd.DataFrame) -> list[str]:
    codigos_cross = detalle_socio[detalle_socio["Es Crosselling"] == True].copy()
    codigos_cross = codigos_cross.sort_values(["Fecha", "Código cliente"], ascending=[True, True])
    lineas: list[str] = []
    for _, r in codigos_cross.iterrows():
        codigo = r.get("Código cliente", "") or "S/D"
        cliente = limpiar_texto(r.get("Cliente", "")) or "SIN NOMBRE"
        fecha = _fecha_guion(r.get("Fecha", ""))
        lineas.append(f"🔸 {codigo} | {cliente} | {fecha}")
    return lineas


def mensaje_avance_socio(detalle_socio: pd.DataFrame, ranking_socio: pd.Series | None = None) -> str:
    if detalle_socio.empty:
        return "⚠️ No hay ventas para el socio seleccionado."

    eh = limpiar_eh(detalle_socio.iloc[0].get("EH", ""))
    socio = limpiar_texto(detalle_socio.iloc[0].get("Socio", "")) or "SIN NOMBRE"

    ventas_obj = int(detalle_socio["Venta objetivo"].sum())
    cross = int(detalle_socio["Es Crosselling"].sum())
    total = int(len(detalle_socio))

    if ranking_socio is not None and not ranking_socio.empty:
        objetivo = int(ranking_socio.get("Objetivo", 0) or 0)
    else:
        objetivo_map, _, _ = objetivos_maps()
        objetivo = int(objetivo_map.get(eh, 0) or 0)

    cumplimiento = (ventas_obj / objetivo * 100) if objetivo > 0 else 0
    faltan = max(objetivo - ventas_obj, 0)

    lineas = [
        "📊 *AVANCE DE VENTAS*",
        "",
        f"👤 *{socio}*",
        f"EH: *{eh or 'S/D'}*",
        "",
        f"✅ Ventas objetivo: *{ventas_obj}*",
        f"🔄 Crosselling: *{cross}*",
        f"📊 Total ventas: *{total}*",
        "",
        f"🎯 Objetivo: *{objetivo}*",
        f"📈 Cumplimiento: *{cumplimiento:.1f}%*",
        f"⏳ Faltan: *{faltan}*",
        "",
        _mensaje_motivacional(ventas_obj, objetivo),
        "",
        "✅ *CÓDIGOS QUE CUENTAN AL OBJETIVO:*",
        "",
    ]

    codigos_obj = _filas_codigos_objetivo(detalle_socio)
    if codigos_obj:
        lineas.extend(codigos_obj)
    else:
        lineas.append("Sin códigos que sumen al objetivo.")

    codigos_cross = _filas_codigos_cross(detalle_socio)
    if codigos_cross:
        lineas.extend([
            "",
            "🔄 *CÓDIGOS CROSSSELLING:*",
            "_(No suman al objetivo de ventas nuevas)_",
            "",
        ])
        lineas.extend(codigos_cross)

    return "\n".join(lineas)


def mensaje_ranking_general(detalle: pd.DataFrame, ranking: pd.DataFrame) -> str:
    total = int(len(detalle))
    ventas_obj = int(detalle["Venta objetivo"].sum()) if not detalle.empty else 0
    cross = int(detalle["Es Crosselling"].sum()) if not detalle.empty else 0
    objetivo_total = int(ranking["Objetivo"].sum()) if not ranking.empty else 0
    cumplimiento = (ventas_obj / objetivo_total * 100) if objetivo_total > 0 else 0

    lineas = [
        "📊 *AVANCE GENERAL DE VENTAS*",
        "",
        f"✅ Ventas objetivo: *{ventas_obj}*",
        f"🔄 Crosselling: *{cross}*",
        f"📊 Total ventas: *{total}*",
        "",
        f"🎯 Objetivo total: *{objetivo_total}*",
        f"📈 Cumplimiento general: *{cumplimiento:.1f}%*",
        "",
        "🏆 *Ranking por socio:*",
    ]

    for i, (_, r) in enumerate(ranking.iterrows(), start=1):
        icono = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "🔹"
        ventas = int(r.get("Ventas objetivo", 0) or 0)
        objetivo = int(r.get("Objetivo", 0) or 0)
        cross_socio = int(r.get("Crosselling", 0) or 0)
        cumplimiento_socio = r.get("Cumplimiento %", "0%")
        faltan = max(objetivo - ventas, 0)
        lineas.append(
            f"{icono} *{r['EH']}* - {limpiar_nombre_socio_mensaje(r['Socio'])}: "
            f"*{ventas}/{objetivo}* ({cumplimiento_socio}) | Falta: *{faltan}* | Cross: *{cross_socio}*"
        )

    lineas += [
        "",
        "💪 Equipo, enfoquemos el seguimiento en ventas nuevas para llegar al objetivo. Crosselling se reporta separado.",
    ]
    return "\n".join(lineas)


def mostrar_whatsapp() -> None:
    st.title("📲 WhatsApp / Reporte por Socio")
    st.caption("Avance de ventas contra objetivo. Crosselling queda separado y no suma al objetivo.")

    cargar_objetivos_widget("wa")

    archivo = st.file_uploader("📤 Subir archivo base", type=["csv", "xlsx", "xls"], key="wa_uploader")
    if archivo is None:
        st.info("Sube el archivo GrossAdd/ventas para generar mensajes automáticos.")
        return

    try:
        df_original = leer_archivo(archivo)
        df, columnas = _preparar_base_comercial(df_original)
    except Exception as exc:
        st.error(f"No se pudo procesar el archivo: {exc}")
        return

    detalle = construir_detalle_comercial(df)
    if detalle.empty:
        st.warning("No se encontraron códigos de cliente en el archivo.")
        return

    ranking = construir_ranking_objetivos(detalle)

    st.subheader("🏆 Ranking de ventas a objetivo")
    st.caption("Ventas objetivo = ventas nuevas. Crosselling se muestra separado y no suma al objetivo.")
    st.dataframe(ranking, use_container_width=True, hide_index=True)

    tipo = st.selectbox(
        "Tipo de mensaje",
        ["Avance por socio", "Ranking general", "Todos los socios separados", "Solo códigos objetivo"],
        key="wa_tipo_mensaje_v2",
    )

    if tipo == "Ranking general":
        texto = mensaje_ranking_general(detalle, ranking)
        st.subheader("📲 Mensaje general para grupo")
        st.text_area("Copiar mensaje", texto, height=420, key="wa_msg_general_v2")
        st.markdown(f"[📲 Enviar por WhatsApp]({whatsapp_link(texto)})")
        return

    socios_opciones = [f"{r['EH']} - {r['Socio']}" for _, r in ranking.iterrows()]
    if not socios_opciones:
        st.warning("No hay socios para mostrar.")
        return

    if tipo in ["Avance por socio", "Solo códigos objetivo"]:
        seleccion = st.selectbox("Seleccionar socio", socios_opciones, key="wa_socio_v2")
        eh_sel = seleccion.split(" - ")[0]
        detalle_socio = detalle[detalle["EH"].astype(str) == eh_sel].copy()
        ranking_socio = ranking[ranking["EH"].astype(str) == eh_sel]
        ranking_row = ranking_socio.iloc[0] if not ranking_socio.empty else None

        st.subheader("📋 Detalle del socio")
        st.dataframe(
            detalle_socio[["Código cliente", "EH", "Socio", "Tipo venta", "Nodo", "Fecha", "Cliente"]],
            use_container_width=True,
            hide_index=True,
        )

        if tipo == "Avance por socio":
            texto = mensaje_avance_socio(detalle_socio, ranking_row)
        else:
            lineas = [
                "✅ *CÓDIGOS QUE CUENTAN AL OBJETIVO*",
                "",
                f"👤 *{detalle_socio.iloc[0]['Socio']}*",
                f"EH: *{eh_sel}*",
                "",
            ]
            codigos = _filas_codigos_objetivo(detalle_socio)
            lineas.extend(codigos if codigos else ["Sin códigos que sumen al objetivo."])
            texto = "\n".join(lineas)

        st.subheader("📲 Mensaje WhatsApp")
        st.text_area("Copiar mensaje", texto, height=520, key="wa_msg_socio_v2")
        st.markdown(f"[📲 Enviar por WhatsApp]({whatsapp_link(texto)})")
        return

    # Todos los socios separados: genera un mensaje individual por cada socio.
    st.subheader("📲 Mensajes separados por socio")
    st.info("Cada socio tiene su propio mensaje y su propio enlace de WhatsApp. Usa esto para enviar los códigos por separado.")

    for _, r in ranking.iterrows():
        eh = str(r["EH"])
        detalle_socio = detalle[detalle["EH"].astype(str) == eh].copy()
        if detalle_socio.empty:
            continue
        texto = mensaje_avance_socio(detalle_socio, r)
        titulo = f"{r['EH']} - {limpiar_nombre_socio_mensaje(r['Socio'])} | {int(r['Ventas objetivo'])}/{int(r['Objetivo'])} | Cross {int(r['Crosselling'])}"
        with st.expander(titulo, expanded=False):
            st.text_area("Mensaje", texto, height=420, key=f"wa_msg_todos_{eh}")
            st.markdown(f"[📲 Enviar por WhatsApp]({whatsapp_link(texto)})")

# =========================================================
# MÓDULO: PENDIENTES DE PAGO
# =========================================================
def mostrar_pendientes_pago() -> None:
    st.title("💳 Pendientes de Pago")
    st.caption("Agrupa pendientes de pago por socio. Para WhatsApp se usa código, nodo y fecha; no se incluye teléfono.")

    with st.sidebar:
        st.subheader("Config. Pago")
        solo_eh_configurados = st.checkbox("Solo mis socios EH", value=False, key="pago_solo_eh")
        with st.expander("Agregar/corregir EH", expanded=False):
            texto_eh_extra = st.text_area("Formato: 59509 Nombre del socio", height=120, key="pago_extra_eh")
        socios = construir_socios(texto_eh_extra)

    archivo = st.file_uploader("📤 Sube archivo Pendientes de Pago", type=["csv", "xlsx", "xls"], key="pago_uploader")
    if archivo is None:
        st.info("Sube el archivo de pendientes de pago para generar el reporte.")
        return

    try:
        df_original = leer_archivo(archivo)
        df, columnas = preparar_pendientes(df_original)
    except Exception as exc:
        st.error(f"No se pudo procesar el archivo: {exc}")
        return

    if solo_eh_configurados:
        df = df[df["_eh"].isin(socios.keys())].copy()

    detalle = pd.DataFrame({
        "Código cliente": df["_codigo"],
        "Nodo": df["_nodo"],
        "Fecha": df["_fecha_base"],
        "EH": df["_eh"],
        "Socio": df.apply(lambda r: r["_socio"] or socios.get(r["_eh"], "SIN NOMBRE"), axis=1),
        "Tipo venta": df["_tipo_venta"],
        "Cliente": df["_cliente"],
        "Observación": df["_observacion"],
    })
    detalle = detalle[detalle["Código cliente"].astype(str).str.strip() != ""].reset_index(drop=True)

    if detalle.empty:
        st.warning("No se encontraron pendientes de pago con el filtro actual.")
        return

    resumen = (
        detalle.groupby(["EH", "Socio"], as_index=False)["Código cliente"]
        .count()
        .rename(columns={"Código cliente": "Pendientes"})
        .sort_values("Pendientes", ascending=False)
        .reset_index(drop=True)
    )

    k1, k2 = st.columns(2)
    k1.metric("Pendientes de pago", len(detalle))
    k2.metric("Socios con pendientes", resumen["EH"].nunique() if not resumen.empty else 0)

    st.subheader("📊 Resumen por socio")
    st.dataframe(resumen, use_container_width=True, hide_index=True)

    col_socio, col_nodo = st.columns(2)
    socios_opciones = ["Todos"] + sorted([s for s in detalle["Socio"].dropna().unique() if str(s).strip()])
    nodos_opciones = ["Todos"] + sorted([n for n in detalle["Nodo"].dropna().unique() if str(n).strip()])
    socio_sel = col_socio.selectbox("Filtrar por socio", socios_opciones, key="pago_fil_socio")
    nodo_sel = col_nodo.selectbox("Filtrar por nodo", nodos_opciones, key="pago_fil_nodo")

    detalle_filtrado = detalle.copy()
    if socio_sel != "Todos":
        detalle_filtrado = detalle_filtrado[detalle_filtrado["Socio"] == socio_sel]
    if nodo_sel != "Todos":
        detalle_filtrado = detalle_filtrado[detalle_filtrado["Nodo"] == nodo_sel]

    st.subheader("📋 Detalle")
    st.dataframe(detalle_filtrado[["Código cliente", "Nodo", "Fecha", "EH", "Socio", "Tipo venta", "Cliente"]], use_container_width=True, hide_index=True)

    st.subheader("📲 Mensaje WhatsApp")
    formato = st.radio("Formato", ["Agrupado por socio", "Código + nodo"], horizontal=True, key="pago_formato")
    lineas = [
        "💳 *PENDIENTES DE PAGO*",
        f"📅 Revisión: *{hoy_bolivia().strftime('%d/%m/%Y')}*",
        f"🔢 Total casos: *{len(detalle_filtrado)}*",
        "",
    ]
    if formato == "Agrupado por socio":
        lineas.append("📋 *Detalle por socio:*")
        for (eh, socio), grupo in detalle_filtrado.groupby(["EH", "Socio"], dropna=False, sort=False):
            lineas.append("")
            lineas.append(f"👤 *{limpiar_nombre_socio_mensaje(socio)}*")
            lineas.append(f"🆔 EH: *{eh or 'S/D'}* | 💳 *{len(grupo)}* casos")
            for i, (_, r) in enumerate(grupo.iterrows(), start=1):
                lineas.append(f"{i}. {r['Código cliente']} | {r['Nodo'] or 'S/D'} | {formatear_fecha_mensaje(r.get('Fecha',''))}")
    else:
        lineas.append("📋 *Código | Nodo | Fecha*")
        for i, (_, r) in enumerate(detalle_filtrado.iterrows(), start=1):
            lineas.append(f"{i}. {r['Código cliente']} | {r['Nodo'] or 'S/D'} | {formatear_fecha_mensaje(r.get('Fecha',''))}")
    lineas.append("")
    lineas.append("✅ Favor contactar al cliente y reportar avance del pago.")
    texto = "\n".join(lineas)

    st.text_area("Copiar mensaje", texto, height=360, key="pago_msg")
    st.markdown(f"[📲 Enviar por WhatsApp]({whatsapp_link(texto)})")

    excel = excel_bytes({"Resumen": resumen, "Detalle": detalle_filtrado})
    st.download_button("⬇️ Descargar Excel", data=excel, file_name=f"pendientes_pago_{hoy_bolivia().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="pago_excel")

# =========================================================
# MÓDULO: OBJETIVOS
# =========================================================
def mostrar_objetivos_configuracion() -> None:
    st.title("🎯 Objetivos Configuración")
    st.caption("Carga el archivo Objetivo.xlsx. Formato soportado: POS_CODE, POS_OWNER, CATEGORIA y BU JUNIO.")

    archivo = st.file_uploader("📤 Subir archivo de objetivos", type=["csv", "xlsx", "xls"], key="obj_uploader")
    if archivo is not None:
        try:
            df_obj_original = leer_archivo(archivo)
            df_obj = normalizar_objetivos_archivo(df_obj_original)
            set_objetivos_df(df_obj)
            st.success(f"Objetivos cargados correctamente: {len(df_obj)} socios.")
        except Exception as exc:
            st.error(f"No se pudo leer objetivos: {exc}")
            try:
                st.write("Columnas detectadas:", list(normalizar_dataframe(leer_archivo(archivo)).columns))
            except Exception:
                pass

    default_df = get_objetivos_df()
    editado = st.data_editor(
        default_df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="objetivos_editor",
    )
    if editado is not None:
        set_objetivos_df(editado)

    objetivos = get_objetivos_df()
    st.subheader("📋 Objetivos configurados")
    st.dataframe(objetivos, use_container_width=True, hide_index=True)

    total_obj = int(pd.to_numeric(objetivos.get("Objetivo", 0), errors="coerce").fillna(0).sum()) if not objetivos.empty else 0
    st.metric("Objetivo total", total_obj)

    lineas = [
        "🎯 *OBJETIVOS POR SOCIO*",
        f"📅 Actualización: *{hoy_bolivia().strftime('%d/%m/%Y')}*",
        f"🔢 Objetivo total: *{total_obj}*",
        "",
    ]
    for _, r in objetivos.iterrows():
        try:
            obj = int(r.get("Objetivo", 0))
        except Exception:
            obj = 0
        if obj > 0:
            lineas.append(f"🔹 {r.get('EH','')} - {nombre_corto(r.get('Socio',''))}: *{obj}*")
    texto = "\n".join(lineas)

    st.subheader("📲 Mensaje WhatsApp")
    st.text_area("Copiar mensaje", texto, height=270, key="obj_msg")
    st.markdown(f"[📲 Enviar por WhatsApp]({whatsapp_link(texto)})")

    excel = excel_bytes({"Objetivos": objetivos})
    st.download_button(
        "⬇️ Descargar objetivos",
        data=excel,
        file_name="objetivos_socios.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="obj_excel",
    )


# =========================================================
# APP PRINCIPAL ÚNICA
# =========================================================
def mostrar_inicio() -> None:
    st.markdown(
        """
        <div style="padding: 30px; border-radius: 18px; background: linear-gradient(135deg, #0033A0, #00AEEF); color: white;">
            <h1>🛠️ Operaciones Tigo Hogar</h1>
            <p style="font-size: 18px;">Agenda técnica, pendientes de instalación, pendientes de pago y suspendidas.</p>
            <hr style="border: 1px solid rgba(255,255,255,.25);">
            <p>👷 💻 Desarrollado por Vladimir Cuenca López</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("### 🏠 Inicio")
    st.write("Selecciona un módulo en el menú lateral.")
    st.info("Esta versión está en un solo archivo app.py y no usa carpeta pages ni imports de módulos externos.")


def main() -> None:
    st.set_page_config(page_title="Autogestión Tigo", page_icon="📊", layout="wide")

    st.sidebar.title("🛠️ Operaciones Tigo")
    modulo = st.sidebar.radio(
        "Selecciona módulo",
        [
            "Inicio",
            "Agenda Técnica",
            "Pendientes de Instalación",
            "Pendientes Pago",
            "Suspendidas",
        ],
        key="menu_principal_unico",
    )

    if modulo == "Inicio":
        mostrar_inicio()
    elif modulo == "Agenda Técnica":
        mostrar_agenda_tecnica()
    elif modulo == "Pendientes de Instalación":
        mostrar_pendientes_inst()
    elif modulo == "Pendientes Pago":
        mostrar_pendientes_pago()
    elif modulo == "Suspendidas":
        mostrar_suspendidas()


if __name__ == "__main__":
    main()
