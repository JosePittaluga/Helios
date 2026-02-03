import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import zipfile
import html
import re
from io import BytesIO

# --- FUNCIONES DE LÓGICA (Tus funciones originales adaptadas) ---

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
    s = s.replace(".", "").replace(",", ".")  # 1.234,56 -> 1234.56
    try:
        return float(s)
    except:
        return 0.0


def limpiar_adenda(texto_sucio):
    if not texto_sucio: return ""
    texto_claro = html.unescape(texto_sucio)
    texto_limpio = re.sub(r'<[^>]+>', ' ', texto_claro)
    return ' '.join(texto_limpio.split())

def buscar_dato(nodo, nombre_tag):
    for elem in nodo.iter():
        tag_name = elem.tag.split('}')[-1]
        if tag_name == nombre_tag:
            return elem.text.strip() if elem.text else ""
    return ""
def item_dict(item):
    d = {}
    for sub in item.iter():
        k = sub.tag.split('}')[-1]
        if k not in d:
            d[k] = (sub.text or "").strip()
    return d


def procesar_contenido_xml(contenido, nombre_archivo):
    try:
        root = ET.fromstring(contenido)
        rut_e = buscar_dato(root, "RUCEmisor")
        rzn_e = buscar_dato(root, "RznSoc")
        rut_r = buscar_dato(root, "DocRecep")
        serie = buscar_dato(root, "Serie")
        nro   = buscar_dato(root, "Nro")
        fch_e = buscar_dato(root, "FchEmis")
        fch_v = buscar_dato(root, "FchVenc")
        moneda = buscar_dato(root, "TpoMoneda")
        tipo_cfe = buscar_dato(root, "TipoCFE")
        adenda_raw = buscar_dato(root, "Adenda")
        adenda_final = limpiar_adenda(adenda_raw)
        

        items = [e for e in root.iter() if e.tag.split('}')[-1] == "Item"]
        lineas_archivo = []
        
        for item in items:
            it = item_dict(item)
            cod_iva = it.get("IndFact", "")
            neto = to_num(it.get("MontoItem"))
            iva_monto = to_num(it.get("IVAMonto"))
            cant = to_num(it.get("Cantidad"))
            precio = to_num(it.get("PrecioUnitario"))

            doc_key = f"{rut_e}|{rut_r}|{serie}|{nro}|{fch_e}|{tipo_cfe}"

            lineas_archivo.append({
                "DocKey": doc_key,
                "Archivo": nombre_archivo,
                "RUT Emisor": rut_e,
                "Razón Social": rzn_e,
                "RUT Receptor": rut_r,
                "Serie-Nro": f"{serie}-{nro}",
                "Fch Emisión": fch_e,
                "Fch Vencimiento": fch_v,
                "Moneda": moneda,
                "Línea": val_i("NroLinDet"),
                "Descripción": val_i("NomItem"),
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
    except Exception as e:
        return []

# --- INTERFAZ STREAMLIT ---

st.title("🛡️ Helios XML Extractor")
st.write("Subí un archivo ZIP con XMLs para generar el Excel consolidado.")

archivo_zip = st.file_uploader("Seleccioná el archivo .ZIP", type=["zip"])

if archivo_zip:
    total_data = []
    ok = 0
    vacios_o_fallidos = 0
    
    with zipfile.ZipFile(archivo_zip, 'r') as z:
        archivos_xml = [f for f in z.namelist() if f.lower().endswith(".xml")]
        
        for nombre_arc in archivos_xml:
            with z.open(nombre_arc) as f:
                contenido = f.read()
                res = procesar_contenido_xml(contenido, nombre_arc)

            if res:
                ok += 1
                total_data.extend(res)
            else:
                vacios_o_fallidos += 1

    if total_data:
        df = pd.DataFrame(total_data)
        
        # Lógica de nombre dinámico (RUT y Fechas)
        rut_receptor = df["RUT Receptor"].dropna().unique()
        rut_str = str(rut_receptor[0]) if len(rut_receptor) > 0 else "SIN_RUT"
        
        df['Fch_DT'] = pd.to_datetime(df['Fch Emisión'], errors='coerce')
        fecha_min = df['Fch_DT'].min()
        fecha_max = df['Fch_DT'].max()
        fmin_str = fecha_min.strftime('%m%Y') if pd.notnull(fecha_min) else "XXXX"
        fmax_str = fecha_max.strftime('%m%Y') if pd.notnull(fecha_max) else "XXXX"
        
        nombre_sugerido = f"ReporteXML_{rut_str}_{fmin_str}_{fmax_str}.xlsx"
        df = df.drop(columns=['Fch_DT'])

        # Mostrar vista previa
        st.success(f"Archivos con líneas: {ok} / {len(archivos_xml)}")
        if vacios_o_fallidos:
            st.warning(f"{vacios_o_fallidos} XML no generaron líneas (vacíos o formato inesperado).")

        st.dataframe(df.head())

        # Botón de descarga
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Descargar Reporte Excel",
            data=output.getvalue(),
            file_name=nombre_sugerido,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:

        st.error("No se encontraron datos válidos dentro de los XML.")

