import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from dotenv import load_dotenv
import time
from datetime import datetime

# Load environment variables from .env file
load_dotenv()

def read_client_data(file_path):
    """
    Read client data from Excel or CSV file
    Returns a pandas DataFrame with client information
    """
    if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
        return pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    else:
        raise ValueError("Unsupported file format. Please use Excel (.xlsx, .xls) or CSV (.csv)")

def send_personalized_email(recipient_email, cc_email, recipient_name, subject, template):
    """
    Send personalized email to a recipient with CC
    """
    # Hardcoded sender email as requested
    sender_email = os.getenv('EMAIL_SENDER')
    sender_password = os.getenv('EMAIL_PASSWORD')
    smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')  # Default to Gmail SMTP
    smtp_port = int(os.getenv('SMTP_PORT', 587))  # Default to TLS port

    if not sender_password:
        raise ValueError("Email password not found. Please set EMAIL_PASSWORD in .env file")

    # Create message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email

    # Add CC if provided
    if cc_email and not pd.isna(cc_email) and cc_email.strip():
        msg['Cc'] = cc_email

    msg['Subject'] = subject

    # Personalize the email content by replacing placeholders
    personalized_content = template.replace('{{NAME}}', recipient_name)

    # Attach the email body
    msg.attach(MIMEText(personalized_content, 'html'))

    # Determine all recipients (for sending)
    recipients = [recipient_email]
    if cc_email and not pd.isna(cc_email) and cc_email.strip():
        recipients.append(cc_email)

    # Send the email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure the connection
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()

        # Get current timestamp for status update
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status_message = f"Sent on {timestamp}"

        print(f"Email sent successfully to {recipient_name} ({recipient_email})")
        return True, status_message
    except Exception as e:
        error_message = f"Failed: {str(e)}"
        print(f"Failed to send email to {recipient_email}: {str(e)}")
        return False, error_message

def main():
    # Configuration
    file_path = 'aarushi_main.xlsx'  # Change to your actual file name
    # email_subject = "sending you test message again, please ignore it"

    # Email template with HTML formatting and placeholders
    email_subject = "Invitation: Q2 Vertex AI Search for Media Product Roadmap Quarterly"

    email_template = """
    <html>
    <body>
        <p>Dear {{NAME}},</p>
        
        <p>You're invited to our upcoming Q2 Vertex AI Search for Media Product Roadmap Quarterly! RSVP today to join us for a firsthand look at the latest features and AI upgrades to Vertex AI Search for Media, along with our plans for the rest of the year.</p>
        
        <p>We'll be sharing some exciting new features and updates, including:</p>
        <ul>
            <li>Personalized search</li>
            <li>Multi-modal video search</li>
            <li>Conversational discovery</li>
        </ul>
        
        <p>This is a great opportunity to learn directly from Google experts, ask questions, and share your feedback with our product team.</p>
        
        <p><strong>Event Details:</strong><br>
        🗓️ Date: May 14th, 11:00am PT<br>
        📅 Event Name: VAIS: Media Q2 Roadmap Quarterly<br>
        🔗 <a href="https://cloudonair.withgoogle.com/events/vais-q2-event">Register here</a></p>
        
        <p>Looking forward to having you with us.</p>
        
        <p>Best,<br>
        Aarushi</p>
    </body>
    </html>
    """

    # Read client data
    try:
        df = read_client_data(file_path)

        # Check if required columns exist
        required_columns = ['First Name', 'Recipient', 'CC', 'Email Sent status']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            print(f"Error: Missing required columns in spreadsheet: {', '.join(missing_columns)}")
            print(f"Available columns are: {', '.join(df.columns)}")
            return

        # Count of emails
        total_rows = len(df)
        successful_emails = 0

        # Process each row and send personalized emails
        for index, row in df.iterrows():
            name = row['First Name']
            recipient_email = row['Recipient']
            cc_email = row['CC']
            current_status = row['Email Sent status']

            # Skip rows with missing recipient email
            if pd.isna(recipient_email) or not str(recipient_email).strip():
                print(f"Skipping row {index+2}: Missing recipient email address")
                continue

            # Skip rows that already have a status (already sent)
            if not pd.isna(current_status) and str(current_status).strip():
                print(f"Skipping row {index+2}: Email already sent previously")
                continue

            # Send the email
            success, status_message = send_personalized_email(
                recipient_email,
                cc_email,
                name,
                email_subject,
                email_template
            )

            # Update the status in the DataFrame
            df.at[index, 'Email Sent status'] = status_message

            if success:
                successful_emails += 1

            # Add a small delay between emails to avoid being flagged as spam
            time.sleep(2)

        # Save the updated DataFrame back to the Excel file
        df.to_excel(file_path, index=False)
        print(f"Updated status information saved to {file_path}")

        # Print summary
        print(f"\nEmail Campaign Summary:")
        print(f"Total rows processed: {total_rows}")
        print(f"Emails sent: {successful_emails}")
        print(f"Rows skipped or failed: {total_rows - successful_emails}")

    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()