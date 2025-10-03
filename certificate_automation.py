import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Step 1: Create a Certificate with Background Image and Text
def create_certificate(name, background_image_path):
    # Load the certificate background image
    certificate = Image.open(background_image_path)
    draw = ImageDraw.Draw(certificate)

    # Define font and size for the text (this is the participant's name)
    font_path = r"C:\Windows\Fonts\arial.ttf"  # Path to Arial font on Windows
    font_medium = ImageFont.truetype(font_path, 40)  # Medium font for name

    # Add text to the certificate (participant's name)
    participant_text = name

    # Get the bounding box of the text (instead of using textsize())
    bbox = draw.textbbox((0, 0), participant_text, font=font_medium)
    text_width, text_height = bbox[2] - bbox[0], bbox[3] - bbox[1]

    # Calculate the position to center the text
    position = ((certificate.width - text_width) // 2, (certificate.height - text_height) // 2)

    # Draw the text (centered)
    draw.text(position, participant_text, font=font_medium, fill="black")

    # Save the certificate as an image
    output_image_path = f"certificates/{name}_certificate.png"
    certificate.save(output_image_path)
    return output_image_path

# Step 2: Send Email with Attachment
def send_email(recipient_email, certificate_path):
    sender_email = "hishamsidhic@gmail.com"  # Your email
    password = "gect ytib hdhh maam"  # Your app-specific password for Google

    # Create message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = "Your Certificate"

    # Attach the certificate image
    part = MIMEBase('application', "octet-stream")
    with open(certificate_path, "rb") as attachment:
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(certificate_path)}"')
    msg.attach(part)

    # Send the email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)  # Use your app-specific password here
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.close()
        print(f"Email sent successfully to {recipient_email}!")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")

# Step 3: Track Delivery Status
def track_delivery_status(email_status):
    # Save delivery status in an Excel sheet
    df = pd.DataFrame(email_status, columns=["Name", "Email", "Status"])
    df.to_excel("certificate_delivery_status.xlsx", index=False)

# Main Function to automate the certificate process
def automate_certificates(participant_list, background_image_path):
    email_status = []
    os.makedirs("certificates", exist_ok=True)  # Create directory for certificates if it doesn't exist

    for participant in participant_list:
        # Step 1: Use the participant's name as is (no spelling correction)
        name = participant['name']

        # Step 2: Generate the certificate with background image and participant's name
        certificate_path = create_certificate(name, background_image_path)

        # Step 3: Send the certificate via email
        try:
            send_email(participant['email'], certificate_path)
            email_status.append([participant['name'], participant['email'], 'Sent'])
        except:
            email_status.append([participant['name'], participant['email'], 'Failed'])

    # Step 4: Track the delivery status
    track_delivery_status(email_status)

# Sample participant data
participants = [
    {"name": "Aseem", "email": "aseemmuhammedct@gmail.com"},
    {"name": "Nadirsha", "email": "mohamednadirsha10@gmail.com"}
]

# Path to your certificate background image
background_image_path = r"C:\Users\OMEN\Downloads\WhatsApp Image 2025-10-03 at 7.29.22 PM.jpeg"

# Automate the certificate process
automate_certificates(participants, background_image_path)