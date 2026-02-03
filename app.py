import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import zipfile
import html
import re
from io import BytesIO
import xmlrpc.client
from functools import lru_cache

# ----------------------------
# ODOO (via st.secrets)
# ----------------------------
# En Streamlit Cloud: Settings -> Secrets
# [odoo]
# url = "https://xxxx"
# db = "xxxx"
# user = "xxxx"
# password = "xxxx"

ODOO_URL = st.secrets["odoo"]["url"]
DB = st.secrets["odoo"]["db"]
USER = st.secrets["odoo"]["user"]
PASS = st.secrets["odoo"]["password"]


@lru_cache(maxsize=1)
def odoo_clients():
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(DB, USER, PASS, {})
    if not uid:
        raise RuntimeError("No se pudo autenticar contra Odoo. Revisá st.secrets.")
    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
    return uid, models


@lru_cache(maxsize=2048)
def get_company_tipo_contabilidad_by_vat(vat: str) -> str:
    vat = (vat or "").strip()
    if not vat:
        return ""

    uid, models = odoo_clients()

    ids = models.execute_kw(
        DB, uid, PASS,
        "res.company", "search",
        [[("vat", "=", vat)]],
        {"limit": 1}
    )
    if not ids:
        return ""

    rec = models.execute_kw(
        DB, uid, PASS,
        "res.company", "read",
        [ids],
        {"fields": ["x_studio_tipo_contabilidad"]}
    )
    val = rec[0].get("x_studio_tipo_contabilidad")

    # Por las dudas (si fuera many2one)
    if isinstance(val, list) and len(val) == 2:
        return val[1] or ""
    return val or ""


# ----------------------------
# FUNCIONES DE LÓGICA
# ----------------------------

def traducir_iva(codigo):
    dict_iva = {
        "1": "Exento", "2": "Tasa Mínima (10%)", "3": "Tasa Básica (22%)",
        "4": "Exportación", "10": "Exportación Servicios"
    }
    return dict_iva.get(str(codigo), "Otros/No Grav.")


def to_num(x):
    if x is None:
        return 0.0
    s = str(x).strip()
    if s == "":
        return 0.0

    # Si tiene coma, asumimos coma decimal (LATAM/Europa): 1.234,56
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    # Si no tiene coma, dejamos el punto como decimal si existe: 22.50
    try:
        return float(s)
    except:
        return 0.0


def limpiar_adenda(texto_sucio):
    if not texto_sucio:
        return ""
    texto_claro = html.unescape(texto_sucio)
    texto_limpio = re.sub(r"<[^>]+>", " ", texto_claro)
    return " ".join(texto_limpio.split())


def buscar_dato(nodo, nombre_tag):
    for elem in nodo.iter():
        tag_name = elem.tag.split("}")[-1]
        if tag_name == nombre_tag:
            return elem.text.strip() if elem.text else ""
    return ""


def extraer_items(item_nodo):
    """Extrae sub-elementos de un Item a dict; se queda con la primera ocurrencia de cada tag."""
    d = {}
    for sub in item_nodo.iter():
        k = sub.tag.split("}")[-1]
        if k not in d:
            d[k] = (sub.text or "").strip()
    return d


def detectar_tipo_por_ruta(nombre_archivo: str) -> str:
    """
    Detecta si el XML es de recibidos o emitidos usando la ruta dentro del ZIP.
    Ajustá keywords si tus carpetas tienen otros nombres.
    """
    p = (nombre_archivo or "").lower()
    if "recib" in p:
        return "recibido"
    if "emit" in p:
        return "emitido"
    return "desconocido"


def procesar_contenido_xml(contenido, nombre_archivo, tipo_doc):
    try:
        root = ET.fromstring(contenido)

        # Cabecera
        rut_e = buscar_dato(root, "RUCEmisor")
        rzn_e = buscar_dato(root, "RznSoc")
        rut_r = buscar_dato(root, "DocRecep")
        serie = buscar_dato(root, "Serie")
        nro = buscar_dato(root, "Nro")
        fch_e = buscar_dato(root, "FchEmis")
        fch_v = buscar_dato(root, "FchVenc")
        moneda = buscar_dato(root, "TpoMoneda")
        tipo_cfe = buscar_dato(root, "TipoCFE")
        adenda_raw = buscar_dato(root, "Adenda")
        adenda_final = limpiar_adenda(adenda_raw)

        # Según carpeta: quién es "la empresa"
        rut_company = ""
        if tipo_doc == "recibido":
            rut_company = rut_r
        elif tipo_doc == "emitido":
            rut_company = rut_e

        # Lookup en Odoo: res.company.vat -> x_studio_tipo_contabilidad
        tipo_contab = ""
        if rut_company:
            tipo_contab = get_company_tipo_contabilidad_by_vat(rut_company)

        # Items
        items_nodos = [e for e in root.iter() if e.tag.split("}")[-1] == "Item"]
        if not items_nodos:
            return []

        lineas_archivo = []
        for nodo in items_nodos:
            it = extraer_items(nodo)

            cod_iva = it.get("IndFact", "")
            neto = to_num(it.get("MontoItem"))
            iva_monto = to_num(it.get("IVAMonto"))
            cant = to_num(it.get("Cantidad"))
            precio = to_num(it.get("PrecioUnitario"))
            desc = it.get("NomItem", "")
            nro_lin = it.get("NroLinDet", "")

            lineas_archivo.append({
                "Archivo": nombre_archivo,
                "Tipo Doc (Carpeta)": tipo_doc,
                "RUT Company (según carpeta)": rut_company,
                "Tipo Contabilidad Company": tipo_contab,

                "RUT Emisor": rut_e,
                "Razón Social": rzn_e,
                "RUT Receptor": rut_r,
                "Serie-Nro": f"{serie}-{nro}",
                "Fch Emisión": fch_e,
                "Fch Vencimiento": fch_v,
                "Moneda": moneda,
                "Línea": nro_lin,
                "Descripción": desc,
                "Cant.": cant,
                "Precio Unit.": precio,
                "Cod. IVA": cod_iva,
                "Tasa IVA": traducir_iva(cod_iva),
                "Neto": neto,
                "Monto IVA": iva_monto,
                "Total Línea": neto + iva_monto,
                "Tipo CFE": tipo_cfe,
                "Adenda": adenda_final
            })

        return lineas_archivo

    except Exception:
        # Para esta etapa, sin errores detallados: si algo falla, no aporta líneas.
        return []


# ----------------------------
# UI STREAMLIT
# ----------------------------

st.title("Helios XML Extractor")
st.write("Subí un archivo ZIP para consolidar tus CFE en Excel, y agregar Tipo de Contabilidad desde Odoo.")

archivo_zip = st.file_uploader("Seleccioná el archivo .ZIP", type=["zip"])

if archivo_zip:
    total_data = []
    ok_count = 0
    vacios_o_fallidos = 0

    # Opcional: mostrar estado de conexión
    try:
        _ = odoo_clients()
        st.info("Conexión a Odoo: OK")
    except Exception as e:
        st.error(f"Conexión a Odoo: FAIL ({e})")
        st.stop()

    with zipfile.ZipFile(archivo_zip, "r") as z:
        archivos_xml = [f for f in z.namelist() if f.lower().endswith(".xml")]

        if not archivos_xml:
            st.error("No se encontraron archivos XML dentro del ZIP.")
            st.stop()

        for nombre_arc in archivos_xml:
            with z.open(nombre_arc) as f:
                contenido = f.read()

            tipo_doc = detectar_tipo_por_ruta(nombre_arc)
            res = procesar_contenido_xml(contenido, nombre_arc, tipo_doc)

            if res:
                total_data.extend(res)
                ok_count += 1
            else:
                vacios_o_fallidos += 1

    if total_data:
        df = pd.DataFrame(total_data)

        # Nombre dinámico
        rut_receptor = df["RUT Receptor"].dropna().unique()
        rut_str = str(rut_receptor[0]) if len(rut_receptor) > 0 else "SIN_RUT"

        df["Fch_DT"] = pd.to_datetime(df["Fch Emisión"], errors="coerce")
        fmin = df["Fch_DT"].min().strftime("%m%Y") if pd.notnull(df["Fch_DT"].min()) else "ini"
        fmax = df["Fch_DT"].max().strftime("%m%Y") if pd.notnull(df["Fch_DT"].max()) else "fin"
        nombre_sugerido = f"ReporteXML_{rut_str}_{fmin}_{fmax}.xlsx"
        df = df.drop(columns=["Fch_DT"])

        st.success(f"XML con líneas: {ok_count} / {len(archivos_xml)}")
        if vacios_o_fallidos:
            st.warning(f"{vacios_o_fallidos} XML no generaron líneas (sin <Item> o formato inesperado).")

        # Preview útil de la nueva columna
        st.write(
            df[["Archivo", "Tipo Doc (Carpeta)", "RUT Company (según carpeta)", "Tipo Contabilidad Company"]]
            .drop_duplicates()
            .head(20)
        )

        st.dataframe(df.head())

        # Export Excel
        output = BytesIO()
        try:
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False)
        except Exception:
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)

        st.download_button(
            label="Descargar Reporte Excel",
            data=output.getvalue(),
            file_name=nombre_sugerido,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("No se pudo extraer información válida de los XML.")
