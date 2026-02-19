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
def get_company_tipo_contabilidad_by_vat(vat: str) -> str:
    vat = (vat or "").strip()
    if not vat: return ""
    uid, models = odoo_clients()
    ids = models.execute_kw(DB, uid, PASS, "res.company", "search", [[("vat", "=", vat)]], {"limit": 1})
    if not ids: return ""
    rec = models.execute_kw(DB, uid, PASS, "res.company", "read", [ids], {"fields": ["x_studio_tipo_contabilidad"]})
    val = rec[0].get("x_studio_tipo_contabilidad")
    if isinstance(val, list) and len(val) == 2: return val[1] or ""
    return val or ""

@lru_cache(maxsize=2048)
def get_company_id_by_vat(vat: str) -> int | None:
    vat = (vat or "").strip()
    if not vat: return None
    uid, models = odoo_clients()
    ids = models.execute_kw(DB, uid, PASS, "res.company", "search", [[("vat", "=", vat)]], {"limit": 1})
    return ids[0] if ids else None

@lru_cache(maxsize=256)
def get_chart_of_accounts(company_id: int) -> list[tuple[str, str]]:
    if not company_id: return []
    uid, models = odoo_clients()
    domain = [[("company_id", "=", company_id)]]
    recs = models.execute_kw(DB, uid, PASS, "account.account", "search_read", domain, {"fields": ["code", "name", "deprecated"], "limit": 5000})
    out = []
    for r in recs:
        if r.get("deprecated"): continue
        code, name = (r.get("code") or "").strip(), (r.get("name") or "").strip()
        if code or name: out.append((f"{code} - {name}".strip(" -"), code))
    out.sort(key=lambda x: x[1] or "")
    return out

@lru_cache(maxsize=128)
def get_odoo_partners(company_id: int) -> set:
    """Trae todos los nombres de contactos (proveedores/clientes) de esa compañía o globales."""
    if not company_id: return set()
    uid, models = odoo_clients()
    # Buscamos partners de la compañía o compartidos (company_id = False)
    domain = ['|', ('company_id', '=', company_id), ('company_id', '=', False)]
    partners = models.execute_kw(DB, uid, PASS, "res.partner", "search_read", [domain], {"fields": ["name"]})
    return {str(p["name"]).strip().lower() for p in partners if p.get("name")}

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
        rut_r, serie, nro = buscar_dato(root, "DocRecep"), buscar_dato(root, "Serie"), buscar_dato(root, "Nro")
        fch_e, fch_v, moneda = buscar_dato(root, "FchEmis"), buscar_dato(root, "FchVenc"), buscar_dato(root, "TpoMoneda")
        tipo_cfe, adenda_final = buscar_dato(root, "TipoCFE"), limpiar_adenda(buscar_dato(root, "Adenda"))

        rut_company = rut_r if tipo_doc == "recibido" else (rut_e if tipo_doc == "emitido" else "")
        tipo_contab = get_company_tipo_contabilidad_by_vat(rut_company) if rut_company else ""

        items_nodos = [e for e in root.iter() if e.tag.split("}")[-1] == "Item"]
        lineas = []
        for nodo in items_nodos:
            it = extraer_items(nodo)
            cod_iva = it.get("IndFact", "")
            neto, iva_monto = to_num(it.get("MontoItem")), to_num(it.get("IVAMonto"))
            lineas.append({
                "Archivo": nombre_archivo, "Tipo Doc (Carpeta)": tipo_doc, "RUT Company (según carpeta)": rut_company,
                "Tipo Contabilidad Company": tipo_contab, "RUT Emisor": rut_e, "Razón Social": rzn_e, "RUT Receptor": rut_r,
                "Serie-Nro": f"{serie}-{nro}", "Fch Emisión": fch_e, "Fch Vencimiento": fch_v, "Moneda": moneda,
                "Línea": it.get("NroLinDet", ""), "Descripción": it.get("NomItem", ""), "Cant.": to_num(it.get("Cantidad")),
                "Precio Unit.": to_num(it.get("PrecioUnitario")), "Cod. IVA": cod_iva, "Tasa IVA": traducir_iva(cod_iva),
                "Neto": neto, "Monto IVA": iva_monto, "Total Línea": neto + iva_monto, "Tipo CFE": tipo_cfe, "Adenda": adenda_final
            })
        return lineas
    except: return []

# ----------------------------
# UI STREAMLIT
# ----------------------------
st.title("Helios XML Extractor & Odoo Sync")
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
        
        # --- SECCIÓN DE CRUCE DE PROVEEDORES ---
        st.header("🔍 Control de Proveedores / Clientes")
        ruts_detectados = [c for c in df["RUT Company (según carpeta)"].dropna().unique() if c]
        
        if ruts_detectados:
            rut_sel = st.selectbox("Seleccioná la Empresa para validar contra Odoo", ruts_detectados)
            comp_id = get_company_id_by_vat(rut_sel)
            
            if comp_id:
                with st.spinner("Validando nombres contra Odoo..."):
                    nombres_odoo = get_odoo_partners(comp_id)
                    # Tomamos Razon Social de lo que NO sea la propia empresa (si es recibido -> emisor, si emitido -> receptor)
                    # Simplificado: cruzamos todos los nombres únicos encontrados en el XML
                    nombres_xml = set(df["Razón Social"].dropna().unique())
                    faltantes = [n for n in nombres_xml if n.strip().lower() not in nombres_odoo]

                if faltantes:
                    st.warning(f"Se detectaron {len(faltantes)} proveedores/clientes que no existen en Odoo (por nombre):")
                    df_f = pd.DataFrame(faltantes, columns=["Razón Social No Encontrada"])
                    st.dataframe(df_f, use_container_width=True)
                    st.download_button("Descargar Faltantes (CSV)", df_f.to_csv(index=False), "faltantes.csv", "text/csv")
                else:
                    st.success("✅ Todos los contactos de los XML ya existen en Odoo.")
            else:
                st.error("No se encontró el ID de la empresa en Odoo para este RUT.")

        # --- REPORTE Y EXPORTACIÓN ---
        st.header("📊 Reporte Consolidado")
        st.dataframe(df.head())
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Descargar Reporte Completo (Excel)",
            data=output.getvalue(),
            file_name="Reporte_Helios_Odoo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No se procesaron datos válidos del ZIP.")
