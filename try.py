import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Function to send email
def send_email(subject, body_template, recipient_info, cc_list=[]):
    # Email account credentials
    sender_email = "email id"  # Authentication email
    sender_password = "app password"

    # Display email to recipients
    display_sender_email = "same mail or alias mail if any"

    # Set up the SMTP server
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    # Connect to the SMTP server
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)

    for recipient_email, recipient_name in recipient_info.items():
        # Create the email object
        message = MIMEMultipart()
        message['From'] = f"display_name <{display_sender_email}>"
        message['To'] = recipient_email
        message['Subject'] = subject
        if cc_list:
            message['Cc'] = ", ".join(cc_list)

        # Personalize the body by replacing the placeholder with the recipient's name
        personalized_body = body_template.replace("{{name}}", recipient_name)
        message.attach(MIMEText(personalized_body, 'html'))

        # Send the email
        all_recipients = [recipient_email] + cc_list
        server.sendmail(sender_email, all_recipients, message.as_string())

    # Disconnect from the server
    server.quit()

# Function to read recipient information from a CSV or Excel file
def read_recipients_from_file(file_path):
    # Load the data into a DataFrame
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    elif file_path.endswith(('.xls', '.xlsx')):
        df = pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file format. Please provide a CSV or Excel file.")

    # Convert the DataFrame into a dictionary
    recipient_info = pd.Series(df['Name'].values, index=df['Email']).to_dict()
    return recipient_info

# Define the subject, body template, recipient information, and CC list
subject = "Subject For the Mail"
body_template = """\
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; }
        h2 { font-weight: bold; }
        h3 { font-weight: 600; margin-bottom: 5px; }
        .content { margin-left: 20px; }
        p { margin: 0 0 10px 0; }
    </style>
</head>
<body>
    <p>Greeting {{name}},</p>

    <p>wishing line</p>

    <p>your message</p>

    <h2>bold heading if needed</h2>

    <h3>1. bulletin 1</h3>
    <div class="content">
        <p>Bulletin Content 1</p>
        <p><a href="Link">Hyperlink Word</a></p>
    </div>

    <h3>2. Bulletin 2</h3>
    <div class="content">
        <p>Bulletin Content 2</p>
        <p><a href="Link">Hyperlink Word</a></p>
    </div>

    <p><strong>Any Small Heading</strong></p>
    <p>Final Words</p>

    <p>Thanking Note</p>


    <p>Ending Greeting,<br>
    Signature</p>
</body>
</html>
"""

recipient_info = {"Recipent mail 1": "recipent name 1",
    "Recipent mail 2": "recipent name 2",
}

cc_list = ["Email list for CC"]

# Path to your CSV or Excel file
# file_path = "Csv file path"  # or "recipients.xlsx"

# Read the recipient information from the file
# recipient_info = read_recipients_from_file(file_path)

# Send the email
send_email(subject, body_template, recipient_info, cc_list)
