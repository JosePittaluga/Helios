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
# CONFIGURACIÓN DE PÁGINA
# ----------------------------
st.set_page_config(page_title="Helios XML Extractor", layout="wide")

# ----------------------------
# ODOO (via st.secrets)
# ----------------------------
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
def get_company_id_by_vat(vat: str) -> int | None:
    vat = (vat or "").strip()
    if not vat: return None
    uid, models = odoo_clients()
    ids = models.execute_kw(DB, uid, PASS, "res.company", "search", [[("vat", "=", vat)]], {"limit": 1})
    return ids[0] if ids else None

@lru_cache(maxsize=128)
def get_odoo_partners_vat(company_id: int) -> set:
    """Trae todos los RUTs (vat) de contactos de esa compañía."""
    if not company_id: return set()
    uid, models = odoo_clients()
    domain = ['|', ('company_id', '=', company_id), ('company_id', '=', False)]
    partners = models.execute_kw(DB, uid, PASS, "res.partner", "search_read", [domain], {"fields": ["vat"]})
    # Limpiamos los ruts para asegurar match (quitar espacios, etc)
    return {str(p["vat"]).strip() for p in partners if p.get("vat")}

# ----------------------------
# FUNCIONES DE LÓGICA XML
# ----------------------------
def traducir_iva(codigo):
    dict_iva = {"1": "Exento", "2": "Tasa Mínima (10%)", "3": "Tasa Básica (22%)", "4": "Exportación", "10": "Exportación Servicios"}
    return dict_iva.get(str(codigo), "Otros/No Grav.")

def to_num(x):
    if x is None: return 0.0
    s = str(x).strip()
    if not s: return 0.0
    if "," in s: s = s.replace(".", "").replace(",", ".")
    try: return float(s)
    except: return 0.0

def limpiar_adenda(texto_sucio):
    if not texto_sucio: return ""
    texto_claro = html.unescape(texto_sucio)
    return " ".join(re.sub(r"<[^>]+>", " ", texto_claro).split())

def buscar_dato(nodo, nombre_tag):
    for elem in nodo.iter():
        if elem.tag.split("}")[-1] == nombre_tag: return elem.text.strip() if elem.text else ""
    return ""

def extraer_items(item_nodo):
    d = {}
    for sub in item_nodo.iter():
        k = sub.tag.split("}")[-1]
        if k not in d: d[k] = (sub.text or "").strip()
    return d

def detectar_tipo_por_ruta(nombre_archivo: str) -> str:
    p = (nombre_archivo or "").lower()
    if "recib" in p: return "recibido"
    if "emit" in p: return "emitido"
    return "desconocido"

def procesar_contenido_xml(contenido, nombre_archivo, tipo_doc):
    try:
        root = ET.fromstring(contenido)
        rut_e, rzn_e = buscar_dato(root, "RUCEmisor"), buscar_dato(root, "RznSoc")
        rut_r, rzn_r = buscar_dato(root, "DocRecep"), buscar_dato(root, "RznSocRecep")
        serie, nro = buscar_dato(root, "Serie"), buscar_dato(root, "Nro")
        fch_e, fch_v, moneda = buscar_dato(root, "FchEmis"), buscar_dato(root, "FchVenc"), buscar_dato(root, "TpoMoneda")
        tipo_cfe, adenda_final = buscar_dato(root, "TipoCFE"), limpiar_adenda(buscar_dato(root, "Adenda"))

        # Identificamos quién es el 'tercero' según si es emitido o recibido
        if tipo_doc == "recibido":
            tercero_rut = rut_e
            tercero_nombre = rzn_e
            rut_company = rut_r
        elif tipo_doc == "emitido":
            tercero_rut = rut_r
            tercero_nombre = rzn_r
            rut_company = rut_e
        else:
            tercero_rut, tercero_nombre, rut_company = "", "", ""

        items_nodos = [e for e in root.iter() if e.tag.split("}")[-1] == "Item"]
        lineas = []
        for nodo in items_nodos:
            it = extraer_items(nodo)
            neto, iva_monto = to_num(it.get("MontoItem")), to_num(it.get("IVAMonto"))
            lineas.append({
                "Archivo": nombre_archivo, "Tipo Doc": tipo_doc, "RUT Company": rut_company,
                "RUT Tercero": tercero_rut, "Nombre Tercero": tercero_nombre,
                "Serie-Nro": f"{serie}-{nro}", "Fch Emisión": fch_e, "Moneda": moneda,
                "Descripción": it.get("NomItem", ""), "Neto": neto, "Monto IVA": iva_monto, "Total Línea": neto + iva_monto
            })
        return lineas
    except: return []

# ----------------------------
# UI STREAMLIT
# ----------------------------
st.title("Helios XML Extractor & Odoo Audit")
archivo_zip = st.file_uploader("Subí el ZIP de Helios", type=["zip"])

if archivo_zip:
    try:
        _ = odoo_clients()
        st.sidebar.success("Conexión Odoo: OK")
    except Exception as e:
        st.error(f"Error Odoo: {e}"); st.stop()

    total_data = []
    with zipfile.ZipFile(archivo_zip, "r") as z:
        xmls = [f for f in z.namelist() if f.lower().endswith(".xml")]
        for arc in xmls:
            with z.open(arc) as f:
                total_data.extend(procesar_contenido_xml(f.read(), arc, detectar_tipo_por_ruta(arc)))

    if total_data:
        df = pd.DataFrame(total_data)
        
        # --- SECCIÓN DE CRUCE POR RUT ---
        st.header("🔍 Control de Contactos por RUT")
        ruts_propios = [c for c in df["RUT Company"].dropna().unique() if c]
        
        if ruts_propios:
            rut_sel = st.selectbox("Seleccioná el RUT de la empresa para auditar", ruts_propios)
            comp_id = get_company_id_by_vat(rut_sel)
            
            if comp_id:
                with st.spinner("Comparando RUTs con Odoo..."):
                    ruts_en_odoo = get_odoo_partners_vat(comp_id)
                    
                    # Agrupamos por RUT de terceros para no repetir en la lista de faltantes
                    terceros = df[df["RUT Company"] == rut_sel][["RUT Tercero", "Nombre Tercero", "Tipo Doc"]].drop_duplicates()
                    
                    # Filtramos los que no están en Odoo
                    faltantes = terceros[~terceros["RUT Tercero"].isin(ruts_en_odoo)]

                if not faltantes.empty:
                    st.warning(f"Se encontraron {len(faltantes)} RUTs en los XML que NO existen en Odoo:")
                    st.dataframe(faltantes, use_container_width=True)
                    
                    # Excel de faltantes
                    buf = BytesIO()
                    faltantes.to_excel(buf, index=False)
                    st.download_button("Descargar RUTs Faltantes (Excel)", buf.getvalue(), "ruts_no_en_odoo.xlsx")
                else:
                    st.success("✅ Todos los emisores/receptores de los XML existen en Odoo.")
            else:
                st.error("No se encontró la empresa seleccionada en Odoo.")

        # --- REPORTE GENERAL ---
        st.header("📊 Reporte Detallado")
        st.dataframe(df)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False)
        st.download_button("Descargar Reporte Completo", output.getvalue(), "Reporte_Helios.xlsx")
    else:
        st.warning("No se encontraron datos procesables.")
