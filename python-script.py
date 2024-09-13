import serial
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Configure the serial port
ser = serial.Serial('/dev/cu.usbserial-1420', 9600)  # Replace with your actual port

# Initialize an empty DataFrame
df = pd.DataFrame(columns=['Timestamp', 'TagID', 'Name', 'Status'])

# Email configuration
SMTP_SERVER = 'smtp.gmail.com'  # Replace with your SMTP server
SMTP_PORT = 587  # Replace with your SMTP port
EMAIL_ADDRESS = 'kiderxkindergarten@gmail.com'  # Replace with your email address
EMAIL_PASSWORD = 'iiqn wtii mvev jjma'  # Replace with your email password
RECIPIENT_EMAIL = 'ardiannsutaj@gmail.com,argjendazizi36@gmail.com,albinsh53@gmail.com,dren.azemi444@gmail.com,'  # Replace with the recipient's email address

def send_email(timestamp, tag_id, name, status):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = RECIPIENT_EMAIL
    msg['Subject'] = 'New RFID Tag Entry'

    # Create the email body
    body = f"New RFID Tag Entry:\n\nTimestamp: {timestamp}\nTagID: {tag_id}\nName: {name}\nStatus: {status}"
    msg.attach(MIMEText(body, 'plain'))

    # Attach the Excel file
    filename = 'rfid_tags.xlsx'
    attachment = open(filename, 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={filename}')

    msg.attach(part)

    # Send the email
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        recipient_list = RECIPIENT_EMAIL.split(',')  # Split recipient emails by comma

        server.sendmail(EMAIL_ADDRESS, recipient_list, msg.as_string())
        server.quit()
        print("Email sent successfully")
    except Exception as e:
        print(f"Failed to send email: {e}")

def update_status(tag_id):
    global df  # Declare df as a global variable

    # Check if the tag ID is already in the DataFrame
    if tag_id in df['TagID'].values:
        # Find the last status of the user
        last_status = df[df['TagID'] == tag_id]['Status'].iloc[-1]
        new_status = 'OUT' if last_status == 'IN' else 'IN'
    else:
        # If the tag ID is not in the DataFrame, set the status to 'IN'
        new_status = 'IN'

    # Append the new row to the DataFrame
    df = df._append({'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 'TagID': tag_id, 'Name': name, 'Status': new_status}, ignore_index=True)

    return new_status




try:
    while True:
        if ser.in_waiting > 0:
            line = ser.readline().decode('utf-8').strip()
            if line.startswith("Tag: "):
                parts = line.split(", ")
                tag_id = parts[0].split(": ")[1]
                name = parts[1].split(": ")[1]
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                # Update status
                status = update_status(tag_id)

                # Save DataFrame to Excel
                df.to_excel('rfid_tags.xlsx', index=False)

                print(f"Saved: {timestamp}, {tag_id}, {name}, {status}")

                # Send email notification
                send_email(timestamp, tag_id, name, status)

except KeyboardInterrupt:
    print("Program interrupted")

finally:
    ser.close()
