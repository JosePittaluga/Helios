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
            def val_i(t):
                for sub in item.iter():
                    if sub.tag.split('}')[-1] == t:
                        return sub.text.strip() if sub.text else ""
                return ""

            cod_iva = val_i("IndFact")
            neto = float(val_i("MontoItem") or 0)
            iva_monto = float(val_i("IVAMonto") or 0)

            lineas_archivo.append({
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
                "Cant.": float(val_i("Cantidad") or 0),
                "Precio Unit.": float(val_i("PrecioUnitario") or 0),
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
    
    with zipfile.ZipFile(archivo_zip, 'r') as z:
        archivos_xml = [f for f in z.namelist() if f.lower().endswith(".xml")]
        
        for nombre_arc in archivos_xml:
            with z.open(nombre_arc) as f:
                contenido = f.read()
                res = procesar_contenido_xml(contenido, nombre_arc)
                total_data.extend(res)

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
        st.success(f"Se procesaron {len(archivos_xml)} archivos correctamente.")
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