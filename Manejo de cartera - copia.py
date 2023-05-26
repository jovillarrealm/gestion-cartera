import pandas as pd
import ssl
import smtplib
import ctypes
import os
import datetime
import pymsgbox
from twilio.rest import Client


def read_excel_file(report_path, client):
    # Read the Excel file
    df = pd.read_excel(report_path)

    # Check if the required columns exist in the dataframe
    required_columns = ["Dias Car", "Poliza", "Asegurado", "Total Pendiente", "email"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        print(f"Columns {missing_columns} not found in the Excel file.")
        return


    # Create a list of lists to store the values
    data = []
    pendientes = []
    devoluciones = []

    # Iterate over each row and extract the values
    for index, row in df.iterrows():
        values = [row[column] for column in required_columns]
        data.append(values)

    for i in data:
        if i[3] > 0:
            if i[0] < 21:
                # Send emails to clients
                # Define email to be sent
                message = (
                    "Estimado/a  "
                    + str(i[2])
                    + "\n"
                    + "\n"
                    + "Esperamos que se encuentre bien. Queremos recordarle "
                    "amablemente que el pago de su poliza numero "
                    + str(i[1])
                    + " esta pendiente por un valor de $"
                    + str(i[3])
                    + ". Para garantizar que la cobertura siga vigente"
                    " y sin interrupciones, por favor realice el pago lo antes posible."
                    + "\n"
                    + "\n"
                    + "Si necesita cualquier tipo de asesoria o tiene alguna inquietud, no dude en contactarnos."
                    " Estaremos encantados de brindarle toda la informacion necesaria."
                    + " Apreciamos sinceramente su atencion a este recordatorio y agradecemos su confianza "
                    "en nuestra compania. Si ya ha realizado el pago, por favor,"
                    " ignore este mensaje."
                    + "\n"
                    + "\n"
                    + "Gracias"
                    + "\n"
                    + "\n"
                    + "Cordialmente,"
                    + "\n"
                    + "\n"
                    + "Orlando de Jesus Gonzalez R"
                )

                context = ssl.create_default_context()
                subject = "Recordatorio de pago"
                message = "Subject: {}\n\n{}".format(subject, message)
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls(context=context)
                server.login("[Correo emisor]", "[contraseña]")
                server.sendmail("[Correo emisor]", i[4], message)
                server.quit()

                # Send WhatsApp messages via Twilio
                send_text(i, client)
                

            else:
                pendientes.append(i[2])
        else:
            devoluciones.append(i[2])



    # Send email to the commertial area
    context = ssl.create_default_context()
    message = (
        "Las personas con mas de 21 en cartera son: "
        + str(pendientes)
        + "\n"
        + "Las personas con devoluciones pendientes son: "
        + str(devoluciones)
    )
    subject = "Urgentes y devoluciones"
    message = "Subject: {}\n\n{}".format(subject, message)
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls(context=context)
    server.login("[Correo emisor]", "[Contrasenia]")
    server.sendmail(
        "[Correo emisor]", "[Correo administrador]", message
    )
    server.quit()


def validate_report(report_path: str):
    MessageBox = ctypes.windll.user32.MessageBoxW
    try:
        # Un float para saber el tiempo de ultima modificacion
        modify_time_f = os.path.getmtime(report_path)
        # Datetime para saber cuando se modifica este reporte
        modify_time = datetime.datetime.fromtimestamp(modify_time_f)
        # Un float para saber el tiempo de creacion
        created_time_f = os.path.getctime(report_path)
        # Datetime para saber cuando se crea este reporte
        created_time = datetime.datetime.fromtimestamp(created_time_f)
    except FileNotFoundError:
        MessageBox(None, "Reporte no encontrado", "Alerta", 0)
        return False

    now = datetime.datetime.now()
    days_since_mod, days_since_created = (now - modify_time).days, (
        now - created_time
    ).days
    if days_since_mod > 4:
        msg = f"El reporte tiene {days_since_mod} dias desde la ultima fecha de modificacion. Desea continuar?"
        if days_since_mod > 365:
            msg = "El reporte tiene más de un año desde ultima fecha de modificacion. Desea continuar?"
        confirm = pymsgbox.confirm(msg, "Reporte posiblemente viejo", ["Si", "No"])
        if confirm == "No":
            return False
    if days_since_created > 4:
        msg = f"El reporte tiene {days_since_created} dias desde su creacion. Desea continuar?"
        if days_since_created > 365:
            msg = "El reporte tiene más de un año de viejo. Desea continuar?"
        confirm = pymsgbox.confirm(msg, "Reporte posiblemente viejo", ["Si", "No"])
        if confirm == "No":
            return False
    return True

# Relate to Twilio API
def send_text(i: list, client: Client):
    # Define WhatsApp text to be sent via Twilio
    w_text = (
        f"Hola, {i[2]}. "
        + f"Esperamos que se encuentre bien. Queremos recordarle amablemente "
        + f"que el pago de su poliza numero {str(i[1])} esta pendiente por un valor de ${str(i[3])}. "
        + "Para garantizar que la cobertura siga vigente "
        + "y sin interrupciones, por favor realice el pago lo antes posible."
    )
    msg = client.messages.create(
        from_="whatsapp:+14155238886",
        body=w_text,
        to=f"whatsapp:+57{3208055735}",
    )


if __name__ == "__main__":
    name = pymsgbox.prompt(
        "Ingrese nombre del excel del reporte sin formato (sin .xlsx)",
        default="cartera",
    )
    
    if name is not None:
        repath = f"{name}.xlsx"
        if validate_report(repath):
            
            # Related to Twilio
            account_sid = "[Token]"
            auth_token = "[Auth Token]"
            client = Client(account_sid, auth_token)
            
            read_excel_file(repath, client)
            MessageBox = ctypes.windll.user32.MessageBoxW
            MessageBox(
                None,
                "Revision de cartera completada",
                "Alerta",
                0,
            )
