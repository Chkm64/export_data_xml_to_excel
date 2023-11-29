from utils.xlsx_writer import Excel
import xml.etree.ElementTree as ET
import os

CARPETA = "xmls"
NAMESPACE_4 = {'cfdi': 'http://www.sat.gob.mx/cfd/4', 'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'}
NAMESPACE_33 = {'cfdi': 'http://www.sat.gob.mx/cfd/3', 'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'}
DICC_CONCEPTO_BASE = {"NombreXML": "",
    "TipoLinea": "",
    "CuentaOImpuesto": "",
    "NombreProveedor": "",
    "UUID": "",
    "ID_PRODUCTO": "",
    "ClaveProdServ": "","ClaveUnidad": "","Unidad": "","Descripcion": "",
    "NoIdentificacion": "",
    "ObjetoImp": "","Cantidad": "","Descuento": "","ValorUnitario":  "",
    "Importe": "",
    "Descuento": "",
    "Base": "",
    "Impuesto": "",
    "TasaOCuota": "",
    "TipoFactor": "",
    "Concatenar": "",
    }

def obtener_nodos_xml(root):
    emisor = False
    # impuestos = False
    conceptos = False
    for child in root:
        if child.tag.split("}")[1] == "Emisor":
            emisor = child.attrib
        # elif child.tag.split("}")[1] == "Impuestos":
        #     impuestos = child
        elif child.tag.split("}")[1] == "Conceptos":
            conceptos = child
    return emisor, conceptos

def convertir_a_float(diccionario):
    for clave, valor in diccionario.items():
        try:
            diccionario[clave] = float(valor)
        except ValueError:
            pass  # Si no se puede convertir a float, deja el valor como está
    return diccionario

def extraer_datos(data_cfdi, data_conceptos, archivo_xml):
    tree = ET.parse(f"{CARPETA}\\{archivo_xml}")
    root = tree.getroot()
    emisor, conceptos = obtener_nodos_xml(root)
    version_cfdi = float(root.attrib.get("Version", "0"))
    timbre_fiscal = root.find('.//cfdi:Complemento/tfd:TimbreFiscalDigital', NAMESPACE_4)
    if version_cfdi == 3.3:
        timbre_fiscal = root.find('.//cfdi:Complemento/tfd:TimbreFiscalDigital', NAMESPACE_33)
    uuid = timbre_fiscal.attrib.get("UUID", "")
    nombre_proveedor = emisor.get("Nombre", "")
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
            "Version": version_cfdi,
            "TipoDeComprobante": root.attrib.get("TipoDeComprobante", ""),
            "Nombre": nombre_proveedor,
            "Rfc": emisor.get("Rfc", ""),
            "RegimenFiscal": emisor.get("RegimenFiscal", ""),
            "CuentaProveedor": "",
            "CuentaProveedorUSD": "",
            "FechaTimbrado": timbre_fiscal.attrib.get("FechaTimbrado", "T").split("T")[0] + " "+ timbre_fiscal.attrib.get("FechaTimbrado", "T").split("T")[1],
            "Concatenar": "",
            })
        cont_productos = 1
        if conceptos:
            for concepto in conceptos:
                data_concepto_base = DICC_CONCEPTO_BASE.copy()
                descripcion = concepto.get("Descripcion", "").replace('\n', '')
                for caractares in [("\n", " "), ("'", "")]:
                    descripcion = descripcion.replace(caractares[0], caractares[1])
                data_concepto_base.update({
                    "NombreXML": archivo_xml,
                    "TipoLinea": "PRODUCTO",
                    "NombreProveedor": nombre_proveedor,
                    "UUID": uuid,
                    "ID_PRODUCTO": cont_productos,
                    "ClaveProdServ": concepto.get("ClaveProdServ", ""),
                    "ClaveUnidad": concepto.get("ClaveUnidad", ""),
                    "Unidad": concepto.get("Unidad", ""),
                    "Descripcion": descripcion,
                    "NoIdentificacion": concepto.get("NoIdentificacion", ""),
                    "ObjetoImp": concepto.get("ObjetoImp", ""),
                    "Cantidad": float(concepto.get("Cantidad", 0)),
                    "Descuento": float(concepto.get("Descuento", 0)),
                    "ValorUnitario":  float(concepto.get("ValorUnitario", 0)),
                    "Importe": float(concepto.get("Importe", 0))
                })
                data_conceptos.append(data_concepto_base)
                for child in concepto:
                    if child.tag.split("}")[1] == "Impuestos":
                        impuestos = child
                        for impuesto in impuestos:
                            if impuesto.tag.split("}")[1] == "Retenciones":
                                for retencion in impuesto:
                                    dicc_retencion_base = DICC_CONCEPTO_BASE.copy()
                                    dicc_retencion_base.update({
                                        "NombreXML": archivo_xml,
                                        "TipoLinea": "RETENCIÓN",
                                        "NombreProveedor": nombre_proveedor,
                                        "UUID": uuid,
                                        "ID_PRODUCTO": cont_productos,
                                        })
                                    dicc_retencion = retencion.attrib
                                    dicc_retencion_base.update(dicc_retencion)
                                    if len(dicc_retencion_base) > 6:
                                        dicc_retencion_base = convertir_a_float(dicc_retencion_base)
                                        data_conceptos.append(dicc_retencion_base)
                            if impuesto.tag.split("}")[1] == "Traslados":
                                for traslado in impuesto:
                                    dicc_traslado_base = DICC_CONCEPTO_BASE.copy()
                                    dicc_traslado_base.update({
                                        "NombreXML": archivo_xml,
                                        "TipoLinea": "TRASLADO",
                                        "NombreProveedor": nombre_proveedor,
                                        "UUID": uuid,
                                        "ID_PRODUCTO": cont_productos,
                                        })
                                    dicc_traslado = traslado.attrib
                                    dicc_traslado_base.update(dicc_traslado)
                                    if len(dicc_traslado_base) > 3:
                                        dicc_traslado_base = convertir_a_float(dicc_traslado_base)
                                        data_conceptos.append(dicc_traslado_base)
                cont_productos += 1
    else:
        print(f"Tipo de Comprobante {tipocomprobante} >> {archivo_xml}")                
    return data_cfdi, data_conceptos

def main():
    data_cfdi = []
    data_conceptos = []
    if os.path.exists(CARPETA) and os.path.isdir(CARPETA):
        archivos = os.listdir(CARPETA)
        archivos_xml = [archivo for archivo in archivos if archivo.endswith(".xml") or archivo.endswith(".XML")]
        for archivo_xml in archivos_xml:
            data_cfdi, data_conceptos = extraer_datos(data_cfdi, data_conceptos, archivo_xml)
        if len(data_cfdi) > 0:
            document_data_cdfi = Excel(f"Extracción de datos - cfdi", data_cfdi)
            document_data_cdfi.generate()
        if len(data_conceptos) > 0:
            document_data_conceptos = Excel(f"Extracción de datos - conceptos", data_conceptos)
            document_data_conceptos.generate()
    else:
        print(f"La carpeta {CARPETA} no existe o no es un directorio.")

if __name__ == '__main__':
    main()
