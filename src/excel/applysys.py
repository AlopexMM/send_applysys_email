from .mail.Mail import SendEmail
import xlrd
import datetime

def process_excel(file_path):
    # Procesa el excel leyendo los datos y generando el objeto para 
    # el envio del email
    
    mail_username = # Tomar datos de variable de entorno
    mail_password = # Tomar datos de variable de entorno
    email_server = SendEmail(
        username = mail_username,
        sender = mail_username,
        password = mail_password,
        mail_server = # Tomar datos de variable de entorno,
        mail_port = 465
    )
    excel = xlrd.open_workbook(filename=file_path)
    sheet = excel.sheet_by_name('Cliente_Comprobantes')

    for row in range(5,sheet.nrows):
        data = {}
        for col in range(sheet.ncols):
            if col == 1:
                data['tipo_doc'] = sheet.cell_value(row, col)
            elif col == 2:
                data['serie'] = sheet.cell_value(row, col)
            elif col == 3:
                data['numero'] = sheet.cell_value(row,col)
            elif col == 4:
                data['importe'] = sheet.cell_value(row,col)
            elif col == 6:
                extract_date_tuple = xlrd.xldate_as_tuple(sheet.cell_value(row,col), excel.datemode)
                convert_to_date = datetime.datetime(*extract_date_tuple).strftime('%d/%m/%Y')
                data['vencimiento'] = convert_to_date
            elif col == 26:
                mails = sheet.cell_value(row,col).split(';')
                data['emails'] = mails
            elif col == 27:
                data['archivo'] = sheet.cell_value(row,col)


        text = f"""Estimado Cliente,

Por medio de la presente le hacemos llegar su factura por los servicios contratados con IPOSS SA(Scanntech):
Comprobante: {data['tipo_doc']} {data['serie']}
Número: {data['numero']}
Importe: ${data['importe']}
Vencimiento: {data['vencimiento']}

Para su mayor comodidad le recordamos que nuestro medio de pago es:
- Pago fácil
- Depósito en cuenta bancaria
- Adhesión débito automático mediante CBU

Por consultas relativas a la factura asociada en este correo puede contactarse por email a mrava@scanntech.com o ldanni@scanntech.com o por teléfono al 011-4849-4436/37.
Recuerde que Usted puede consultar sus facturas pendientes en www.scanntech.com ingresando con su Usuario y contraseña.
Descargue el programa Acrobat Reader para Visualizar su factura: http://get.adobe.com/es/reader/

Atentamente.

Departamento de Facturación IPOSS SA Migueletes 1231 – 3A 1426 - Ciudad de Buenos Aires"""

        for email in data['emails']:
            if email == '':
                continue
            else:
                email_server.send_email(
                    body = text,
                    recipent = email,
                    subject = "Factura",
                    attach = data['archivo'],
                )

if __name__ == '__main__':

    process_excel('./testing.xls')