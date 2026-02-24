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
    if not company_id: return set()
    uid, models = odoo_clients()
    domain = ['|', ('company_id', '=', company_id), ('company_id', '=', False)]
    partners = models.execute_kw(DB, uid, PASS, "res.partner", "search_read", [domain], {"fields": ["vat"]})
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
        # Cabecera
        rut_e = buscar_dato(root, "RUCEmisor")
        rzn_e = buscar_dato(root, "RznSoc")
        rut_r = buscar_dato(root, "DocRecep")
        rzn_r = buscar_dato(root, "RznSocRecep")
        serie = buscar_dato(root, "Serie")
        nro = buscar_dato(root, "Nro")
        fch_e = buscar_dato(root, "FchEmis")
        fch_v = buscar_dato(root, "FchVenc")
        moneda = buscar_dato(root, "TpoMoneda")
        tipo_cfe = buscar_dato(root, "TipoCFE")
        adenda_final = limpiar_adenda(buscar_dato(root, "Adenda"))

        # Identificamos la empresa propia para la auditoría
        rut_company = rut_r if tipo_doc == "recibido" else rut_e

        items_nodos = [e for e in root.iter() if e.tag.split("}")[-1] == "Item"]
        lineas = []
        for nodo in items_nodos:
            it = extraer_items(nodo)
            neto = to_num(it.get("MontoItem"))
            iva_monto = to_num(it.get("IVAMonto"))
            cod_iva = it.get("IndFact", "")
            
            lineas.append({
                "Archivo": nombre_archivo,
                "RUT Emisor": rut_e,
                "Razón Social": rzn_e,
                "RUT Receptor": rut_r,
                "Serie-Nro": f"{serie}-{nro}",
                "Fch Emisión": fch_e,
                "Fch Vencimiento": fch_v,
                "Moneda": moneda,
                "Línea": it.get("NroLinDR", ""),
                "Descripción": it.get("NomItem", ""),
                "Cant.": to_num(it.get("Cantidad")),
                "Precio Unit.": to_num(it.get("PrecioUnitario")),
                "Cod. IVA": cod_iva,
                "Tasa IVA": traducir_iva(cod_iva),
                "Neto": neto,
                "Monto IVA": iva_monto,
                "Total Línea": neto + iva_monto,
                "Tipo CFE": tipo_cfe,
                "Adenda": adenda_final,
                "RUT Company": rut_company, # Oculto para lógica
                "Nombre Receptor": rzn_r # Oculto para lógica
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
        
        # --- AUDITORÍA DE RUTS ---
        st.header("🔍 Control de Contactos por RUT")
        ruts_propios = [c for c in df["RUT Company"].dropna().unique() if c]
        
        if ruts_propios:
            rut_sel = st.selectbox("Seleccioná el RUT de tu empresa para auditar", ruts_propios)
            comp_id = get_company_id_by_vat(rut_sel)
            
            if comp_id:
                with st.spinner("Comparando con Odoo..."):
                    ruts_en_odoo = get_odoo_partners_vat(comp_id)
                    
                    # Consolidar terceros (emisores si soy receptor, receptores si soy emisor)
                    emisores = df[df["RUT Emisor"] != rut_sel][["RUT Emisor", "Razón Social"]].rename(columns={"RUT Emisor": "RUT", "Razón Social": "Nombre"})
                    receptores = df[df["RUT Receptor"] != rut_sel][["RUT Receptor", "Nombre Receptor"]].rename(columns={"RUT Receptor": "RUT", "Nombre Receptor": "Nombre"})
                    
                    terceros = pd.concat([emisores, receptores]).drop_duplicates(subset=["RUT"])
                    faltantes = terceros[(~terceros["RUT"].isin(ruts_en_odoo)) & (terceros["RUT"] != "")]

                if not faltantes.empty:
                    st.warning(f"Se encontraron {len(faltantes)} RUTs que NO existen en Odoo:")
                    st.dataframe(faltantes, use_container_width=True)
                    buf = BytesIO()
                    faltantes.to_excel(buf, index=False)
                    st.download_button("Descargar RUTs Faltantes", buf.getvalue(), "faltantes_odoo.xlsx")
                else:
                    st.success("✅ Todos los contactos ya existen en Odoo.")

        # --- REPORTE DETALLADO ---
        st.header("📊 Reporte Detallado")
        # Columnas finales según tu imagen
        cols_finales = ["Archivo", "RUT Emisor", "Razón Social", "RUT Receptor", "Serie-Nro", 
                        "Fch Emisión", "Fch Vencimiento", "Moneda", "Línea", "Descripción", 
                        "Cant.", "Precio Unit.", "Cod. IVA", "Tasa IVA", "Neto", "Monto IVA", 
                        "Total Línea", "Tipo CFE", "Adenda"]
        
        st.dataframe(df[cols_finales])
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df[cols_finales].to_excel(writer, index=False)
        st.download_button("Descargar Reporte Completo", output.getvalue(), "Reporte_Helios.xlsx")
    else:
        st.warning("No se encontraron datos.")
