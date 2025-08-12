import os
import sys
import xml.etree.ElementTree as ET
import pandas as pd
from tkinter import messagebox, Tk
import traceback

# ==== Utilidades de namespaces y búsqueda ====
NS33 = "http://www.sat.gob.mx/cfd/3"
NS40 = "http://www.sat.gob.mx/cfd/4"
NS_TFD = "http://www.sat.gob.mx/TimbreFiscalDigital"
NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"

def find_one(elem, local, default_ns_list=(NS33, NS40)):
    """Busca el primer elemento con nombre local 'local' en cualquiera de los namespaces dados."""
    for ns in default_ns_list:
        node = elem.find(f'.//{{{ns}}}{local}')
        if node is not None:
            return node
    return None

def findall_any(elem, local, default_ns_list=(NS33, NS40)):
    """Busca todos los elementos con nombre local 'local' en cualquiera de los namespaces."""
    out = []
    for ns in default_ns_list:
        out.extend(elem.findall(f'.//{{{ns}}}{local}'))
    return out

# === Carpeta de trabajo (igual que tu lógica) ===
if getattr(sys, 'frozen', False):
    carpeta_xml = os.path.dirname(sys.executable)
else:
    carpeta_xml = os.path.dirname(os.path.abspath(__file__))

root_tk = Tk()
root_tk.withdraw()
messagebox.showinfo("Buscando XML", f"Buscando archivos XML en:\n\n{carpeta_xml}")

registros = []

try:
    archivos_xml = [f for f in os.listdir(carpeta_xml) if f.lower().endswith('.xml')]
    if not archivos_xml:
        raise FileNotFoundError("No se encontraron archivos XML en esta carpeta.")

    for archivo in archivos_xml:
        ruta = os.path.join(carpeta_xml, archivo)
        tree = ET.parse(ruta)
        root_node = tree.getroot()
        data = {'Archivo': archivo}

        # ===== Atributos del Comprobante (3.3 y 4.0) =====
        # Lista ampliada con campos comunes en ambas versiones
        attrs_comprobante = [
            'Version','Serie','Folio','Fecha','Moneda','TipoCambio','SubTotal','Descuento','Total',
            'TipoDeComprobante','FormaPago','MetodoPago','LugarExpedicion','Confirmacion','Exportacion',
            'Sello','NoCertificado','Certificado'
        ]
        for a in attrs_comprobante:
            data[f'@{a}'] = root_node.attrib.get(a, '')

        # schemaLocation (xsi)
        data['@xsi:schemaLocation'] = root_node.attrib.get(f'{{{NS_XSI}}}schemaLocation', '')

        # ===== Emisor =====
        emisor = find_one(root_node, 'Emisor')
        if emisor is not None:
            # 3.3: Rfc, Nombre, RegimenFiscal
            # 4.0: Rfc, Nombre (opcional), RegimenFiscal
            for a in ['Rfc', 'Nombre', 'RegimenFiscal']:
                data[f'/cfdi:Emisor/@{a}'] = emisor.attrib.get(a, '')

        # ===== Receptor =====
        receptor = find_one(root_node, 'Receptor')
        if receptor is not None:
            # 3.3: Rfc, Nombre(opc), ResidenciaFiscal, NumRegIdTrib, UsoCFDI
            # 4.0: Rfc, Nombre(opc), DomicilioFiscalReceptor, RegimenFiscalReceptor, UsoCFDI
            for a in [
                'Rfc','Nombre','UsoCFDI',
                'ResidenciaFiscal','NumRegIdTrib',                 # 3.3
                'DomicilioFiscalReceptor','RegimenFiscalReceptor'  # 4.0
            ]:
                data[f'/cfdi:Receptor/@{a}'] = receptor.attrib.get(a, '')

        # ===== CfdiRelacionados (opcional) =====
        relacionados = find_one(root_node, 'CfdiRelacionados')
        if relacionados is not None:
            # Tipo de relación (3.3 y 4.0)
            data['/cfdi:CfdiRelacionados/@TipoRelacion'] = relacionados.attrib.get('TipoRelacion', '')
            rels = findall_any(relacionados, 'CfdiRelacionado')
            for i, rel in enumerate(rels, 1):
                data[f'/cfdi:CfdiRelacionados/cfdi:CfdiRelacionado[{i}]/@UUID'] = rel.attrib.get('UUID', '')

        # ===== Complemento / Timbre Fiscal Digital =====
        # Recorremos todos los Complementos en 3.3 y 4.0
        complementos = findall_any(root_node, 'Complemento')
        for complemento in complementos:
            # Timbre Fiscal Digital
            # Está en namespace propio
            tfd = complemento.find(f'.//{{{NS_TFD}}}TimbreFiscalDigital')
            if tfd is not None:
                for a in ['UUID', 'FechaTimbrado', 'NoCertificadoSAT', 'RfcProvCertif', 'SelloCFD', 'SelloSAT', 'Version']:
                    data[f'/cfdi:Complemento/tfd:TimbreFiscalDigital/@{a}'] = tfd.attrib.get(a, '')
                data['/cfdi:Complemento/tfd:TimbreFiscalDigital/@xsi:schemaLocation'] = tfd.attrib.get(
                    f'{{{NS_XSI}}}schemaLocation', ''
                )

        # ===== Impuestos (global) =====
        impuestos = find_one(root_node, 'Impuestos')
        if impuestos is not None:
            data['/cfdi:Impuestos/@TotalImpuestosTrasladados'] = impuestos.attrib.get('TotalImpuestosTrasladados', '')
            data['/cfdi:Impuestos/@TotalImpuestosRetenidos'] = impuestos.attrib.get('TotalImpuestosRetenidos', '')

            # Traslados globales (pueden ser varios)
            traslados_padre = findall_any(impuestos, 'Traslados')
            idx_t = 0
            for tp in traslados_padre:
                for t in findall_any(tp, 'Traslado'):
                    idx_t += 1
                    for a in ['Base','Importe','Impuesto','TasaOCuota','TipoFactor']:
                        data[f'/cfdi:Impuestos/cfdi:Traslados/cfdi:Traslado[{idx_t}]/@{a}'] = t.attrib.get(a, '')

            # Retenciones globales (pueden ser varias)
            ret_padre = findall_any(impuestos, 'Retenciones')
            idx_r = 0
            for rp in ret_padre:
                for r in findall_any(rp, 'Retencion'):
                    idx_r += 1
                    for a in ['Base','Importe','Impuesto','TasaOCuota','TipoFactor']:
                        data[f'/cfdi:Impuestos/cfdi:Retenciones/cfdi:Retencion[{idx_r}]/@{a}'] = r.attrib.get(a, '')

        # ===== Conceptos =====
        conceptos = findall_any(root_node, 'Concepto')
        for i, c in enumerate(conceptos, 1):
            base = f'/cfdi:Conceptos/cfdi:Concepto[{i}]'
            # Atributos comunes 3.3/4.0
            for a in [
                'Cantidad','ClaveProdServ','ClaveUnidad','Unidad','Descripcion',
                'NoIdentificacion','ValorUnitario','Importe','Descuento',
                'ObjetoImp'  # 4.0
            ]:
                data[f'{base}/@{a}'] = c.attrib.get(a, '')

            # Impuestos por concepto (varios)
            imp_c = findall_any(c, 'Impuestos')
            # Por si hay varios nodos Impuestos bajo el concepto
            idx_ct = 0
            idx_cr = 0
            for ic in imp_c:
                # Traslados por concepto
                for tp in findall_any(ic, 'Traslados'):
                    for t in findall_any(tp, 'Traslado'):
                        idx_ct += 1
                        for a in ['Base','Importe','Impuesto','TasaOCuota','TipoFactor']:
                            data[f'{base}/cfdi:Impuestos/cfdi:Traslados/cfdi:Traslado[{idx_ct}]/@{a}'] = t.attrib.get(a, '')
                # Retenciones por concepto
                for rp in findall_any(ic, 'Retenciones'):
                    for r in findall_any(rp, 'Retencion'):
                        idx_cr += 1
                        for a in ['Base','Importe','Impuesto','TasaOCuota','TipoFactor']:
                            data[f'{base}/cfdi:Impuestos/cfdi:Retenciones/cfdi:Retencion[{idx_cr}]/@{a}'] = r.attrib.get(a, '')

        registros.append(data)

    # Exportar a Excel
    df = pd.DataFrame(registros)
    archivo_salida = os.path.join(carpeta_xml, 'resultado_cfdi.xlsx')
    df.to_excel(archivo_salida, index=False)

    messagebox.showinfo("✅ Éxito", f"Se generó el archivo:\n{archivo_salida}")

except Exception as e:
    messagebox.showerror("❌ Error", f"{str(e)}\n\n{traceback.format_exc()}")
