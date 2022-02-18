import os
from threading import Thread
from tkinter import Button, Frame, Label, Tk
from tkinter.filedialog import askdirectory
from tkinter.messagebox import askokcancel, showinfo
from tkinter.ttk import Progressbar

from lxml import etree
from openpyxl import Workbook

NSMAP = {
    'cfdi': 'http://www.sat.gob.mx/cfd/3',
    'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital',
    'implocal': 'http://www.sat.gob.mx/implocal',
    'pago10': 'http://www.sat.gob.mx/Pagos',
}

NSMAP_V4 = {
    'cfdi': 'http://www.sat.gob.mx/cfd/4',
    'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital',
    'implocal': 'http://www.sat.gob.mx/implocal',
    'pago20': 'http://www.sat.gob.mx/Pagos20',
}


class MainApplication(Frame):
    def __init__(self, master, *args, **kwargs):
        Frame.__init__(self, master, *args, **kwargs)

        self.master.title('CFDI 2 Excel')

        self.master.resizable(0, 0)

        self.lbl_titulo = Label(
            self, text='CFDI 2 Excel', font=('Arial Bold', 20)
        )
        self.lbl_titulo.grid(
            column=0, row=0, columnspan=4,
            sticky='ew', padx=10, pady=10
        )

        self.btn_fuente = Button(
            self, text='...', command=self.btn_fuente_click
        )
        self.btn_fuente.grid(row=1, column=0, sticky='ew', padx=10, pady=10)

        self.lbl_fuente = Label(self, text='Folder fuente:')
        self.lbl_fuente.grid(row=1, column=1, sticky='ew', padx=10, pady=10)

        self.lbl_path_fuente = Label(self, text=os.getcwd())
        self.lbl_path_fuente.grid(
            row=1, column=2, columnspan=2,
            sticky='ew', padx=10, pady=10
        )

        self.btn_destino = Button(
            self, text='...', command=self.btn_destino_click
        )
        self.btn_destino.grid(row=2, column=0, sticky='ew', padx=10, pady=10)

        self.lbl_destino = Label(self, text='Folder destino:')
        self.lbl_destino.grid(row=2, column=1, sticky='ew', padx=10, pady=10)

        self.lbl_path_destino = Label(self, text=os.getcwd())
        self.lbl_path_destino.grid(
            row=2, column=2, columnspan=2,
            sticky='ew', padx=10, pady=10
        )

        self.pgb_status = Progressbar(self)
        self.pgb_status.grid(
            row=3, column=0, columnspan=3,
            sticky='ew', padx=10, pady=10
        )

        self.btn_procesar = Button(
            self, text='Procesar', command=self.btn_procesar_click
        )
        self.btn_procesar.grid(row=3, column=3, sticky='ew', padx=10, pady=10)

        self.lbl_estado = Label(self, text='Listo')
        self.lbl_estado.grid(
            row=4, column=0, columnspan=4,
            sticky='ew', padx=10, pady=10
        )

    def btn_fuente_click(self):
        path = self.lbl_path_fuente['text']
        path = askdirectory(initialdir=path)
        if path:
            self.lbl_path_fuente['text'] = path

    def btn_destino_click(self):
        path = self.lbl_path_destino['text']
        path = askdirectory(initialdir=path)
        if path:
            self.lbl_path_destino['text'] = path

    def btn_procesar_click(self):
        res = askokcancel('Confirmar', '¿Seguro que quiere procesar?')

        if not res:
            return

        self.btn_fuente['state'] = 'disabled'
        self.btn_destino['state'] = 'disabled'
        self.btn_procesar['state'] = 'disabled'

        Thread(target=self.task_xml_to_excel).start()

    def task_xml_to_excel(self):
        self.pgb_status.start()

        fuente_path = self.lbl_path_fuente['text']
        destino_path = self.lbl_path_destino['text']

        files = [os.path.join(dp, f) for dp, dn, filenames in os.walk(
            fuente_path) for f in filenames if os.path.splitext(f)[1] == '.xml']

        text = 'XML encontrados: {}'.format(len(files))
        self.lbl_estado['text'] = text

        wb = Workbook()

        sheet = wb.active

        sheet.append(
            (
                'Versión CFDI',
                'UUID',
                'UUIDs relacionados',
                'Tipo relacion',
                'RFC emisor',
                'Razon emisor',
                'RFC receptor',
                'Razon receptor',
                'Fecha emision',
                'Fecha certificacion',
                'Fecha cancelacion',
                'Estado',
                'Relacionados',
                'NomBancoExt',
                'RfcEmisorCtaOrd',
                'CtaOrdenante',
                'RfcEmisorCtaBen',
                'CtaBeneficiario',
                'TipoCadPago',
                'CadPago',
                'Número operación',
                'MonedaP',
                'TipoCambioP',
                'FormaDePagoP',
                'Fecha pago',
                'Monto',
                'Id documento',
                'Estatus (Almacen)',
                'Fecha emision (Doc)',
                'Fecha certificacion (Doc)',
                'Fecha cancelacion (Doc)',
                'Serie',
                'Folio',
                'MonedaDR',
                'TipoCambioDR',
                'MetodoDePagoDR',
                'NumParcialidad',
                'Saldo anterior',
                'Importe pagado',
                'Saldo actual'
            )
        )

        for f in files:
            # try:
            if True:
                self.lbl_estado['text'] = 'Procesando: {}'.format(f)

                root = etree.parse(f, parser=etree.XMLParser(
                    huge_tree=True, recover=True)).getroot()

                version = root.get('Version')

                if version == '4.0':
                    nsmap = NSMAP_V4
                    pagonsmap = 'pago20'
                else:
                    nsmap = NSMAP
                    pagonsmap = 'pago10'

                uuid = root.find(
                    'cfdi:Complemento/tfd:TimbreFiscalDigital',
                    namespaces=nsmap
                ).get('UUID')

                serie = root.get('Serie')

                folio = root.get('Folio')

                tipo = root.get('TipoDeComprobante')

                fecha = root.get('Fecha')

                fecha_timbrado = root.find(
                    'cfdi:Complemento/tfd:TimbreFiscalDigital',
                    namespaces=nsmap
                ).get('FechaTimbrado')

                pac = root.find(
                    'cfdi:Complemento/tfd:TimbreFiscalDigital',
                    namespaces=nsmap
                ).get('RfcProvCertif')

                rfc_emisor = root.find(
                    'cfdi:Emisor',
                    namespaces=nsmap
                ).get('Rfc')

                nombre_emisor = root.find(
                    'cfdi:Emisor',
                    namespaces=nsmap
                ).get('Nombre')

                rfc_receptor = root.find(
                    'cfdi:Receptor',
                    namespaces=nsmap
                ).get('Rfc')

                nombre_receptor = root.find(
                    'cfdi:Receptor',
                    namespaces=nsmap
                ).get('Nombre')

                conceptos = ''

                for i, c in enumerate(root.findall('cfdi:Conceptos/cfdi:Concepto', namespaces=nsmap)):
                    conceptos += '|-{}-|: {}: {} '.format(
                        i + 1,
                        c.get('Descripcion'),
                        c.get('Importe')
                    )

                uso = root.find(
                    'cfdi:Receptor',
                    namespaces=nsmap
                ).get('UsoCFDI')

                moneda = root.get('Moneda')

                tipo_cambio = root.get('TipoCambio')

                metodo_pago = root.get('MetodoPago')

                forma_pago = root.get('FormaPago')

                subtotal = root.get('SubTotal')

                descuento = root.get('Descuento')

                iva = 0.0
                isr = 0.0
                ieps = 0.0
                for t in root.findall('cfdi:Impuestos/cfdi:Traslados/cfdi:Traslado', namespaces=nsmap):
                    if t.get('Impuesto') == '002':
                        iva += float(t.get('Importe'))
                    if t.get('Impuesto') == '001':
                        isr += float(t.get('Importe'))
                    if t.get('Impuesto') == '003':
                        ieps += float(t.get('Importe'))

                iva_ret = 0
                isr_ret = 0
                for t in root.findall('cfdi:Impuestos/cfdi:Retenciones/cfdi:Retencion', namespaces=nsmap):
                    if t.get('Impuesto') == '002':
                        iva_ret += float(t.get('Importe'))
                    if t.get('Impuesto') == '001':
                        isr_ret += float(t.get('Importe'))

                total = root.get('Total')

                tipo_relacion = ''
                relaciones = ''

                cfdi_relacionados = root.find(
                    'cfdi:CfdiRelacionados', namespaces=nsmap)

                if cfdi_relacionados is not None:

                    tipo_relacion = cfdi_relacionados.get('TipoRelacion')

                    for r in cfdi_relacionados.findall('cfdi:CfdiRelacionado', namespaces=nsmap):
                        relaciones += '{}, '.format(
                            r.get('UUID')
                        )

                implocal = 0

                for t in root.findall('cfdi:Complemento/implocal:ImpuestosLocales/implocal:TrasladosLocales', namespaces=nsmap):
                    implocal += float(t.get('Importe'))

                pago_documento_relaciones = ''
                pago_documento_relacionados = root.find(
                    f'cfdi:Complemento/{pagonsmap}:Pagos/{pagonsmap}:Pago', namespaces=nsmap
                )
                if pago_documento_relacionados is not None:
                    cuenta_ordenante = pago_documento_relacionados.get(
                        'CtaOrdenante')
                    rfc_emisor_cuenta_ben = pago_documento_relacionados.get(
                        'RfcEmisorCtaBen')
                    cuenta_beneficiario = pago_documento_relacionados.get(
                        'CtaBeneficiario')
                    numero_operacion = pago_documento_relacionados.get(
                        'NumOperacion')
                    moneda_pago = pago_documento_relacionados.get('MonedaP')
                    tipo_cambio_pago = pago_documento_relacionados.get(
                        'TipoCambioP')
                    forma_de_pago_pago = pago_documento_relacionados.get(
                        'FormaDePagoP')
                    fecha_pago = pago_documento_relacionados.get('FechaPago')
                    monto = pago_documento_relacionados.get('Monto')
                    for r in pago_documento_relacionados.findall(f'{pagonsmap}:DoctoRelacionado', namespaces=nsmap):
                        pago_documento_relaciones += '{}, '.format(
                            r.get('IdDocumento')
                        )
                        id_dr = r.get('IdDocumento', '')
                        serie_dr = r.get('Serie', '')
                        folio_dr = r.get('Folio', '')
                        moneda_dr = r.get('MonedaDR', '')
                        tipo_cambio_dr = r.get('TipoCambioDR', '')
                        metodo_de_pago_dr = r.get('MetodoDePagoDR', '')
                        numero_parcialidad = r.get('NumParcialidad', '')
                        saldo_anterior = r.get('ImpSaldoAnt', '')
                        importe_pagado = r.get('ImpPagado', '')
                        saldo_actual = r.get('ImpSaldoInsoluto', '')

                sheet.append(
                    (
                        version,  # 'Versión CFDI',
                        uuid,  # 'UUID',
                        relaciones,  # 'UUIDs relacionados',
                        tipo_relacion,  # 'Tipo relacion',
                        rfc_emisor,  # 'RFC emisor',
                        nombre_emisor,  # 'Razon emisor',
                        rfc_receptor,  # 'RFC receptor',
                        nombre_receptor,  # 'Razon receptor',
                        fecha,  # 'Fecha emision',
                        fecha_timbrado,  # 'Fecha certificacion',
                        'Fecha cancelacion',
                        'Estado',
                        pago_documento_relaciones,  # 'Relacionados',
                        'NomBancoExt',
                        'RfcEmisorCtaOrd',
                        cuenta_ordenante if 'cuenta_ordenante' in locals() else "",  # 'CtaOrdenante',
                        rfc_emisor_cuenta_ben if 'rfc_emisor_cuenta_ben' in locals(
                        ) else "",  # 'RfcEmisorCtaBen',
                        cuenta_beneficiario if 'cuenta_beneficiario' in locals() else '',  # 'CtaBeneficiario',
                        'TipoCadPago',
                        'CadPago',
                        numero_operacion if 'numero_operacion' in locals() else '',  # 'Número operación',
                        moneda_pago if 'moneda_pago' in locals() else '',  # 'MonedaP',
                        tipo_cambio_pago if 'tipo_cambio_pago' in locals() else '',  # 'TipoCambioP',
                        forma_de_pago_pago if 'forma_de_pago_pago' in locals() else '',  # 'FormaDePagoP',
                        fecha_pago if 'fecha_pago' in locals()else'',  # 'Fecha pago',
                        monto if 'monto' in locals() else'',  # 'Monto',
                        id_dr if 'id_dr' in locals() else "",  # 'Id documento',
                        'Estatus (Almacen)',
                        'Fecha emision (Doc)',
                        'Fecha certificacion (Doc)',
                        'Fecha cancelacion (Doc)',
                        serie_dr if 'serie_dr' in locals() else '',  # 'Serie',
                        folio_dr if 'folio_dr' in locals() else '',  # 'Folio',
                        moneda_dr if 'moneda_dr' in locals() else '',  # 'MonedaDR',
                        tipo_cambio_dr if 'tipo_cambio_dr' in locals() else '',  # 'TipoCambioDR',
                        metodo_de_pago_dr if 'metodo_de_pago_dr' in locals() else '',  # 'MetodoDePagoDR',
                        numero_parcialidad if 'numero_parcialidad' in locals() else '',  # 'NumParcialidad',
                        saldo_anterior if 'saldo_anterior' in locals() else '',  # 'Saldo anterior',
                        importe_pagado if 'importe_pagado' in locals() else '',  # 'Importe pagado',
                        saldo_actual if 'saldo_actual' in locals() else '',  # 'Saldo actual'
                    )
                )
            # except Exception as e:
            #     sheet.append((str(e), ))

        file_path = os.path.join(destino_path, 'cfdis.xlsx')

        wb.save(file_path)

        self.pgb_status.stop()

        self.btn_fuente['state'] = 'normal'
        self.btn_destino['state'] = 'normal'
        self.btn_procesar['state'] = 'normal'

        showinfo(
            'Completado',
            'Proceso completado\nArchivo guardado en: {}'.format(file_path)
        )

        os.system('start excel "{}"'.format(file_path))


if __name__ == '__main__':
    root = Tk()

    MainApplication(root).grid(row=0, column=0)

    root.mainloop()
