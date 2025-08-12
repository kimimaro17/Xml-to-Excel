import os
import sys
import xml.etree.ElementTree as ET
import pandas as pd
from tkinter import messagebox, Tk
import traceback

# Obtener la carpeta donde está el EXE (o el .py)
if getattr(sys, 'frozen', False):
    carpeta_xml = os.path.dirname(sys.executable)
else:
    carpeta_xml = os.path.dirname(os.path.abspath(__file__))

# Mostrar al usuario desde dónde está buscando
root = Tk()
root.withdraw()
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

        atributos_raiz = [
            'Certificado','Fecha','Folio','FormaPago','LugarExpedicion','MetodoPago','Moneda',
            'NoCertificado','Sello','Serie','SubTotal','TipoCambio','TipoDeComprobante','Total','Version'
        ]
        for attr in atributos_raiz:
            data[f'@{attr}'] = root_node.attrib.get(attr, '')

        data['@xsi:schemaLocation'] = root_node.attrib.get('{http://www.w3.org/2001/XMLSchema-instance}schemaLocation', '')

        emisor = root_node.find('.//{http://www.sat.gob.mx/cfd/3}Emisor')
        if emisor is None:
            emisor = root_node.find('.//{http://www.sat.gob.mx/cfd/4}Emisor')
        if emisor is not None:
            for attr in ['Nombre', 'RegimenFiscal', 'Rfc']:
                data[f'/cfdi:Emisor/@{attr}'] = emisor.attrib.get(attr, '')

        receptor = root_node.find('.//{http://www.sat.gob.mx/cfd/3}Receptor')
        if receptor is None:
            receptor = root_node.find('.//{http://www.sat.gob.mx/cfd/4}Receptor')
        if receptor is not None:
            for attr in ['Rfc', 'UsoCFDI']:
                data[f'/cfdi:Receptor/@{attr}'] = receptor.attrib.get(attr, '')

        complementos = root_node.findall('.//{http://www.sat.gob.mx/cfd/3}Complemento')
        if not complementos:
            complementos = root_node.findall('.//{http://www.sat.gob.mx/cfd/4}Complemento')
        for complemento in complementos:
            for nodo in complemento:
                if 'TimbreFiscalDigital' in nodo.tag:
                    for attr in ['UUID', 'FechaTimbrado', 'NoCertificadoSAT', 'RfcProvCertif', 'SelloCFD', 'SelloSAT', 'Version']:
                        data[f'/cfdi:Complemento/tfd:TimbreFiscalDigital/@{attr}'] = nodo.attrib.get(attr, '')
                    data['/cfdi:Complemento/tfd:TimbreFiscalDigital/@xsi:schemaLocation'] = nodo.attrib.get('{http://www.w3.org/2001/XMLSchema-instance}schemaLocation', '')

        impuestos = root_node.find('.//{http://www.sat.gob.mx/cfd/3}Impuestos')
        if impuestos is None:
            impuestos = root_node.find('.//{http://www.sat.gob.mx/cfd/4}Impuestos')
        if impuestos is not None:
            data['/cfdi:Impuestos/@TotalImpuestosTrasladados'] = impuestos.attrib.get('TotalImpuestosTrasladados', '')
            traslado = impuestos.find('.//{http://www.sat.gob.mx/cfd/3}Traslado')
            if traslado is None:
                traslado = impuestos.find('.//{http://www.sat.gob.mx/cfd/4}Traslado')
            if traslado is not None:
                for attr in ['Importe', 'Impuesto', 'TasaOCuota', 'TipoFactor']:
                    data[f'/cfdi:Impuestos/cfdi:Traslados/cfdi:Traslado/@{attr}'] = traslado.attrib.get(attr, '')

        conceptos = root_node.findall('.//{http://www.sat.gob.mx/cfd/3}Concepto')
        if not conceptos:
            conceptos = root_node.findall('.//{http://www.sat.gob.mx/cfd/4}Concepto')
        for i, concepto in enumerate(conceptos):
            base = f'/cfdi:Conceptos/cfdi:Concepto[{i+1}]'
            for attr in ['Cantidad', 'ClaveProdServ', 'ClaveUnidad', 'Descripcion',
                         'Importe', 'NoIdentificacion', 'ValorUnitario']:
                data[f'{base}/@{attr}'] = concepto.attrib.get(attr, '')
            traslado = concepto.find('.//{http://www.sat.gob.mx/cfd/3}Traslado')
            if traslado is None:
                traslado = concepto.find('.//{http://www.sat.gob.mx/cfd/4}Traslado')
            if traslado is not None:
                for attr in ['Base', 'Importe', 'Impuesto', 'TasaOCuota', 'TipoFactor']:
                    data[f'{base}/cfdi:Impuestos/cfdi:Traslados/cfdi:Traslado/@{attr}'] = traslado.attrib.get(attr, '')

        registros.append(data)

    # Exportar a Excel
    df = pd.DataFrame(registros)
    archivo_salida = os.path.join(carpeta_xml, 'resultado_cfdi.xlsx')
    df.to_excel(archivo_salida, index=False)

    messagebox.showinfo("✅ Éxito", f"Se generó el archivo:\n{archivo_salida}")

except Exception as e:
    messagebox.showerror("❌ Error", f"{str(e)}\n\n{traceback.format_exc()}")
