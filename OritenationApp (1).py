import json
import qrcode
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
import tempfile
import csv

def create_smtp_server(smtp_server,port,sender_email,sender_password):
    context = ssl.create_default_context()
    server = smtplib.SMTP_SSL(smtp_server, port, context=context)
    server.login(sender_email,sender_password)
    return server

def read_csv_file():
    with open('test.csv', mode='r') as file:
        csv_reader = csv.DictReader(file)
        csv_data = list(csv_reader)
    file.close()
    return csv_data
    
def generate_qr_code_from_dict(data):
    # Convert the dictionary to a JSON string
    data_str = json.dumps(data)

    # Create a QRCode object
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )

    # Add data to the QRCode object
    qr.add_data(data_str)
    qr.make(fit=True)

    # Create an image from the QRCode object
    img = qr.make_image(fill='black', back_color='white')

    return img

def create_pdf_in_memory(qr_image,name):
    # Create a BytesIO object to hold the PDF content
    pdf_buffer = BytesIO()

    # Create a PDF canvas
    c = canvas.Canvas(pdf_buffer, pagesize=letter)
    width, height = letter

    # Draw logo
    logo_path = "Gandhinagar University Logo - Final.png"  # Replace with the path to your image
    logo_width = 250   # Desired width of the logo
    logo_height = 125  # Desired height of the logo
    logo_x = (width - logo_width) / 2   # X-coordinate to place the logo
    logo_y = height - logo_height - 50

    #name
    name_text = name
    name_font_size = 50  # Font size for the name text
    name_x = (width - c.stringWidth(name_text, "Times-Roman", name_font_size)) / 2.5
    name_y = 120 # Positioning below the QR code

    #TEXT
    text = "ORIENTATION 2024"
    font_size = 35  # Desired font size for the text
    text_x = (width - c.stringWidth(text, "Times-Roman", font_size))/2
    text_y = 25  # Position the text 30 points below the QR code

    text2 = "03-AUGUST"
    text2_font_size =32
    text2_x = (width - c.stringWidth(text2,"Times-Roman",text2_font_size))/2
    text2_y = 75

    # Save the QR code image to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as temp_img_file:
        qr_image.save(temp_img_file, format='PNG')
        temp_img_path = temp_img_file.name

    # Draw the QR code image on the PDF canvas
    c.drawImage(temp_img_path, (width - qr_image.size[0]) / 2, (height - qr_image.size[1]) / 2, width=qr_image.size[0], height=qr_image.size[1])
    c.drawImage(logo_path, logo_x, logo_y, width=logo_width, height=logo_height) #Drawing the logo
    c.setFont("Times-Roman", font_size)#text
    c.drawString(text_x, text_y, text)
    c.setFont("Helvetica-BoldOblique", name_font_size)#for users name
    c.drawString(name_x, name_y, name_text)
    c.setFont("Times-Roman", text2_font_size)
    c.drawString(text2_x, text2_y, text2)

    # Save the PDF content to the BytesIO object
    c.save()

    # Get the PDF content from the BytesIO object
    pdf_buffer.seek(0)

    # Clean up the temporary file
    import os
    os.remove(temp_img_path)

    return pdf_buffer

def send_email_with_pdf(sender_email, recipient_email, subject, body, pdf_buffer, pdf_filename,server):
    # Create the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Attach the email body
    msg.attach(MIMEText(body, 'plain'))

    # Attach the PDF file
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(pdf_buffer.getvalue())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={pdf_filename}')
    msg.attach(part)

    # Create a secure SSL context
    
    server.send_message(msg, from_addr=sender_email, to_addrs=recipient_email)
    
def main():
    sender_email = "nimesh.gajjar@gandhinagaruni.ac.in"  # Enter your address
    sender_password = 'yapg ypjv lhzz lpnb' # Add Your Password
    subject = 'QR Code For GU Orientation 2024'
    body = 'Dear Students,\n \n \nI am sharing registeration QR Code to attend Orientation 2024.\n \nIt is compulsory to have downloaded QR Code PDF in your device at the time of registeration while coming on 3rd-Aug-24 at our university campus.\n\n\nThanks & Regards,\nDr.Nimesh Gajjar'
    
    smtp_server = 'smtp.gmail.com'
    port = 465
    
    server = create_smtp_server(
        smtp_server=smtp_server,
        port=port,
        sender_email=sender_email,
        sender_password=sender_password
    )
    print("SMTP SERVER CREATED")
    
    csv_data = read_csv_file()
    for row in csv_data:

        recipient_email = row['email']
        pdf_filename = row['name']+".pdf"


        print("Email:",recipient_email)
        print("File Name:",pdf_filename)

        qr_image = generate_qr_code_from_dict(row)
        pdf_buffer = create_pdf_in_memory(qr_image,row['name'])
        send_email_with_pdf(
            sender_email=sender_email,
            recipient_email=recipient_email, 
            subject=subject, 
            body=body, 
            pdf_buffer=pdf_buffer, 
            pdf_filename=pdf_filename, 
            server=server
        )
        print("MAIL SENT...")

    server.quit()
    print("SMTP SERVER CLOSED")


if __name__ == '__main__':
    main()