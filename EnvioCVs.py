import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from optparse import OptionParser
import configparser
import pandas as pd
import numpy as np
import time
from datetime import datetime
import re

MAX_EMAILS_TO_SEND = 25
EXCLUDED_COMPANIES = ['outlook', 'gmail', 'hotmail', 'live', 'yahoo', 'terra', 'telefonica', 'notariado', 'correonotarial']

def send_emails(config, destinatarios, options):

    emails_send = 0
    s = smtplib.SMTP(config['EMAIL_CONF']['EMAIL_SERVER'], config['EMAIL_CONF']['EMAIL_PORT'])
    s.ehlo() # Hostname to send for this command defaults to the fully qualified domain name of the local host.
    s.starttls() #Puts connection to SMTP server in TLS mode
    s.ehlo()
    s.login(config['EMAIL_CONF']['EMAIL_FROM'], config['EMAIL_CONF']['EMAIL_PASSWD'])

    f = open(config['EMAIL_MSG']['EMAIL_TEXT'],'r', encoding="utf-8")
    text_email = f.read()

    rows, colums = destinatarios.shape
    enviados = []
    for i in range(rows):
        destinatario = destinatarios['email'][i].strip()
        enviado = destinatarios['enviado'][i]

        if(len(enviado) == 0):
            email_valid = check_email(destinatario)
            duplicated, info_duplicated = search_duplicated(enviados, destinatario)
            
            if(not email_valid):
                print("INFO: '%s' is invalid" % (destinatario))
                destinatarios['enviado'][i] = 'INVALID'
                destinatarios['fecha'][i] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            elif(duplicated):
                print("INFO: '%s' is duplicated: %s" % (destinatario, info_duplicated))
                destinatarios['enviado'][i] = 'DUPLICATED'
                destinatarios['fecha'][i] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                destinatarios['notas'][i] = info_duplicated
            elif(not options.duplicates):
                msg = MIMEMultipart()
                msg['From'] = config['EMAIL_CONF']['EMAIL_FROM']
                msg['To'] = destinatario
                msg['Subject'] = config['EMAIL_MSG']['EMAIL_SUBJECT']

                # Add body to email
                msg.attach(MIMEText(text_email, "plain"))

                # Add attachment to message and convert message to string
                part = adjuntar_archivo(config)
                if(part):
                    msg.attach(part)

                s.sendmail(config['EMAIL_CONF']['EMAIL_FROM'], destinatario, msg.as_string().encode('utf-8'))

                print("[%d] Email enviado a '%s'" % (i, destinatario))
                time.sleep(5)

                destinatarios['enviado'][i] = 'SI'
                destinatarios['fecha'][i] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

                emails_send += 1
            else:
                print("WARNING! Email to '%s' is not send. Only Check duplicates" % (destinatario))
        
        # Add to enviados list
        enviados.append(destinatario)

        if(emails_send >= MAX_EMAILS_TO_SEND):
            break

    s.quit()

    print(destinatarios)

    guardar_destinatarios(options, destinatarios)

def adjuntar_archivo(config):
    part = None
    filename = config['EMAIL_MSG']['EMAIL_ATTACHMENT']

    if(len(filename) > 0):
        # Open PDF file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email    
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

    return part

def search_duplicated(enviados, destinatario):
    duplicated = False
    info_duplicated = ''
    #company_destinatario = destinatario.split('@')[1].split('.')[0]
    #if not company_destinatario in EXCLUDED_COMPANIES:
    #    for enviado in enviados:
    #        company = enviado.split('@')[1].split('.')[0]
    #        if(company_destinatario == company):
    #            duplicated = True
    #            info_duplicated = enviado
    #            break
    #else:
    if(destinatario in enviados):
        duplicated = True
        info_duplicated = destinatario

    return duplicated, info_duplicated


def obtener_destinatarios(options):
    #destinatarios = ['pedroemisario@hotmail.com', 'pedroemisario@gmail.com']
    destinatarios = None

    # Assign spreadsheet filename to `file` and load spreadsheet
    xl = pd.ExcelFile(options.destinatarios)

    # Print the sheet names
    print(xl.sheet_names)

    # Load a sheet into a DataFrame by name: df1
    destinatarios = xl.parse('destinatarios')
    destinatarios = destinatarios.replace(np.nan, '', regex=True)

    print(destinatarios)

    return destinatarios

def guardar_destinatarios(options, destinatarios):
    # Specify a writer
    writer = pd.ExcelWriter(options.destinatarios, engine='xlsxwriter')

    # Write your DataFrame to a file     
    destinatarios.to_excel(writer, 'destinatarios')

    # Save the result 
    writer.save()

def check_email(email):
    regex = '^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$'
    if(re.search(regex,email)):
        return True
    else:
        return False

if __name__ == '__main__':

    # Check arguments
    parser = OptionParser(usage="%prog: [options]")
    parser.add_option("-f", "--forzarEnvio", dest="forzarEnvio", default=False, action="store_true", help='Forzar envio Emails')
    parser.add_option("-c", "--config", dest="config", default="data.ini", type="string", help='Configuration File')
    parser.add_option("-d", "--destinatarios", dest="destinatarios", default="", type="string", help='Fichero con destinatarios')
    parser.add_option("-p", "--duplicates", dest="duplicates", default=False, action="store_true", help='Comprobar solo duplicados')

    (options, args) = parser.parse_args()

    config = configparser.ConfigParser()
    config.read(options.config)

    destinatarios = obtener_destinatarios(options)
    send_emails(config, destinatarios, options)