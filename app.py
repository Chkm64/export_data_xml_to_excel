from utils.xlsx_writer import Excel
import xml.etree.ElementTree as ET
import os

CARPETA = "xmls"

def obtener_nodos_xml(root):
    emisor = False
    impuestos = False
    complemento = False
    for child in root:
        if child.tag.split("}")[1] == "Emisor":
            emisor = child.attrib
        elif child.tag.split("}")[1] == "Impuestos":
            impuestos = child
        elif child.tag.split("}")[1] == "Complemento":
            complemento = child
    return emisor, impuestos, complemento

def convertir_a_float(diccionario):
    for clave, valor in diccionario.items():
        try:
            diccionario[clave] = float(valor)
        except ValueError:
            pass  # Si no se puede convertir a float, deja el valor como está
    return diccionario

def extraer_datos(data_cfdi, data_impuestos, archivo_xml):
    tree = ET.parse(f"{CARPETA}\\{archivo_xml}")
    root = tree.getroot()
    emisor, impuestos, complemento = obtener_nodos_xml(root)
    timbre_fiscal = False
    for timbre in complemento:
        timbre_fiscal = timbre
    uuid = timbre_fiscal.attrib.get("UUID", "")
    tipocomprobante = root.attrib.get("TipoDeComprobante", 0)
    if tipocomprobante == "I":
        data_cfdi.append({
            "NombreXML": archivo_xml,
            "UUID": uuid,
            "Fecha": root.attrib.get("Fecha", "T").split("T")[0] + " "+ root.attrib.get("Fecha", "T").split("T")[1],
            "Serie": root.attrib.get("Serie", ""),
            "Folio": root.attrib.get("Folio", ""),
            "Moneda": root.attrib.get("Moneda", "") ,
            "Total": float(root.attrib.get("Total", "0")),
            "Subtotal": float(root.attrib.get("SubTotal", "0")),
            "Descuento": float(root.attrib.get("Descuento", "0")),
            "Version": float(root.attrib.get("Version", "0")),
            "TipoDeComprobante": root.attrib.get("TipoDeComprobante", ""),
            "Nombre": emisor.get("Nombre", ""),
            "Rfc": emisor.get("Rfc", ""),
            "RegimenFiscal": emisor.get("RegimenFiscal", ""),
            "FechaTimbrado": timbre_fiscal.attrib.get("FechaTimbrado", "T").split("T")[0] + " "+ timbre_fiscal.attrib.get("FechaTimbrado", "T").split("T")[1],
            })
        if impuestos:
            for child in impuestos:
                if child.tag.split("}")[1] == "Retenciones":
                    for retencion in child:
                        dicc_retencion = retencion.attrib
                        dicc_retencion_base = {
                            "NombreXML": archivo_xml,
                            "TipoImpuesto": "RETENCIÓN",
                            "UUID": uuid,
                            "Base": "",
                            "TasaOCuota": "",
                            "TipoFactor": ""
                            }
                        dicc_retencion_base.update(dicc_retencion)
                        if len(dicc_retencion_base) > 6:
                            dicc_retencion_base = convertir_a_float(dicc_retencion_base)
                            data_impuestos.append(dicc_retencion_base)
                if child.tag.split("}")[1] == "Traslados":
                    for traslado in child:
                        dicc_traslado_base = {
                            "NombreXML": archivo_xml,
                            "TipoImpuesto": "TRASLADO",
                            "UUID": uuid
                            }
                        dicc_traslado = traslado.attrib
                        dicc_traslado_base.update(dicc_traslado)
                        if len(dicc_traslado_base) > 3:
                            dicc_traslado_base = convertir_a_float(dicc_traslado_base)
                            data_impuestos.append(dicc_traslado_base)
    return data_cfdi, data_impuestos

def main():
    data_cfdi = []
    data_impuestos = []
    if os.path.exists(CARPETA) and os.path.isdir(CARPETA):
        archivos = os.listdir(CARPETA)
        archivos_xml = [archivo for archivo in archivos if archivo.endswith(".xml")]
        for archivo_xml in archivos_xml:
            # print(archivo_xml)
            data_cfdi, data_impuestos = extraer_datos(data_cfdi, data_impuestos, archivo_xml)
        if len(data_cfdi) > 0:
            document_data_cdfi = Excel(f"Extracción de datos - cfdi", data_cfdi)
            document_data_cdfi.generate()
        if len(data_impuestos) > 0:
            document_data_impuestos = Excel(f"Extracción de datos - impuestos", data_impuestos)
            document_data_impuestos.generate()
    else:
        print(f"La carpeta {CARPETA} no existe o no es un directorio.")

if __name__ == '__main__':
    main()
