import pandas as pd
import smtplib


# LOS DATOS DE LA CUENTA DE LA QUE SE ENVIAN LOS EMAILS

name_account = "Nombre de cuenta"
email_account = "email@gmail.com"   
password_account = "password"       # tiene que ser la contraseña creada para app 


# CREAMOS UN SERVIDOR SEGURO CON SSL
# EN ESTE CASO ES POR QUE LA CUENTA ES GMAIL, PARA OTROS COMO OUTLOOK, LOS ARGUMENTOS SON DISTINTOS

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()                                       # INICIA EL SERVIDOR
server.login(email_account, password_account)       # SE LOGUEA LA CUENTA

# LEEMOS EL EXCEL USANDO PANDAS

read_data = pd.read_excel("./emails.xlsx")

# OBTENEMOS LOS DATOS DE LA TABLA DE EXCEL, VA A DEPENDER DE LOS DATOS QUE NECESITES PARA ENVIAR LOS EMAILS.
# EN EL EXCEL SE PUEDE ESPECIFICAR EL ASUNTO Y EL MENSAJE PARA CADA PERSONA, EN ESTE CASO VAMOS A MANDAR UN MENSAJE Y ASUNTO GENERALIZADO

all_names = read_data['Nombre']
all_emails = read_data['Email']
#all_subjets = read_data['Asunto']
#all_messages = read_data['Mensajes']

# ITERAMOS POR CADA FILA DENTRO DEL EXCEL

for i in range(len(read_data)):
    name = all_names[i]
    email = all_emails[i]


    # personalizamos el asunto y el mensaje

    subject = ("Hola, " + name)

    message = ('Hola ' + name + '!\n\n' + 
               'Este es un mensaje masivo automatizado con Python!\n\n' + 
               'Te saluda atte. \n' + name_account )
    
# DANDOLE FORMATO PARA ENVIAR EL EMAIL    

    sent_email = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}"
                  .format(name_account, email_account, name, email, subject, message))
    

# ENVIAMOS EL MENSAJE  Y SI NO SE PUEDE ENVIAR, VA A TIRAR POR CONSOLA EL ERROR

    try: 
        server.sendmail(email_account, [email], sent_email)
    except Exception:
        print("no se pudo enviar el mail a {}. Error: {}\n".format(email, str(Exception)))

server.close()            