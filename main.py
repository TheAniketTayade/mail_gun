import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from dotenv import load_dotenv
import time
from datetime import datetime
import json
import re
from pathlib import Path

# Load environment variables from .env file
load_dotenv()

class EmailConfig:
    """Configuration class to make settings easier to manage"""
    def __init__(self):
        # Email settings - users can modify these
        self.EXCEL_FILE = 'aarushi_main.xlsx'  # Your client data file
        self.EMAIL_SUBJECT = "Invitation: Q2 Vertex AI Search for Media Product Roadmap Quarterly"
        self.TEMPLATE_FILE = 'email_template.html'  # Optional: external template file
        self.ATTACHMENTS_FOLDER = 'attachments'  # Folder containing files to attach
        self.DELAY_BETWEEN_EMAILS = 2  # Seconds between each email (to avoid spam filters)
        self.MAX_EMAILS_PER_RUN = 100  # Maximum emails to send in one run
        self.TEST_MODE = False  # Set to True to preview emails without sending
        self.TEST_EMAIL = "test@example.com"  # Email to use in test mode

        # Column names in Excel (can be customized)
        self.COLUMNS = {
            'name': 'First Name',
            'to': 'To',  # Changed from 'Recipient'
            'cc': 'CC',
            'bcc': 'BCC',  # New column for BCC
            'status': 'Email Status',
            'timestamp': 'Sent Timestamp',
            'attachments': 'Attachments',  # Column for specific attachments per recipient
            'custom_subject': 'Custom Subject'  # Optional custom subject per recipient
        }

def read_config_file():
    """Read configuration from a simple config file if it exists"""
    config_file = 'email_config.json'
    if os.path.exists(config_file):
        with open(config_file, 'r') as f:
            return json.load(f)
    return {}

def validate_email(email):
    """Validate email format"""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email.strip()) is not None

def parse_email_list(email_string):
    """Parse comma-separated email list and validate each email"""
    if pd.isna(email_string) or not str(email_string).strip():
        return []

    emails = [email.strip() for email in str(email_string).split(',')]
    valid_emails = []
    invalid_emails = []

    for email in emails:
        if email and validate_email(email):
            valid_emails.append(email)
        elif email:
            invalid_emails.append(email)

    if invalid_emails:
        print(f"Warning: Invalid email addresses found and skipped: {', '.join(invalid_emails)}")

    return valid_emails

def read_email_template(template_path='email_template.html'):
    """Read email template from HTML file"""
    if os.path.exists(template_path):
        print(f"✓ Loading email template from: {template_path}")
        with open(template_path, 'r', encoding='utf-8') as f:
            return f.read()
    else:
        print(f"⚠️  Email template file '{template_path}' not found!")
        print("Creating a default template file for you...")

        # Create a default template file
        default_template = """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
        }
        .container {
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
        }
        .header {
            background-color: #f8f9fa;
            padding: 20px;
            text-align: center;
            border-radius: 5px;
        }
        .content {
            padding: 20px 0;
        }
        .footer {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #eee;
            font-size: 14px;
            color: #666;
        }
        .button {
            display: inline-block;
            padding: 10px 20px;
            background-color: #007bff;
            color: white;
            text-decoration: none;
            border-radius: 5px;
            margin: 10px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>Welcome {{NAME}}!</h2>
        </div>
        
        <div class="content">
            <p>Dear {{NAME}},</p>
            
            <p>This is your email template. You can edit this HTML file to change the email content.</p>
            
            <p><strong>Available placeholders you can use:</strong></p>
            <ul>
                <li><code>{{NAME}}</code> - Recipient's first name</li>
                <li><code>{{DATE}}</code> - Current date</li>
                <li><code>{{TO}}</code> - Recipient's email</li>
                <li><code>{{COMPANY}}</code> - Company name (if column exists)</li>
                <li><code>{{ANY_COLUMN_NAME}}</code> - Any column from your Excel file (in UPPERCASE)</li>
            </ul>
            
            <p>You can add images, links, buttons, and any HTML content:</p>
            
            <a href="https://example.com" class="button">Click Here</a>
            
            <p>Edit this file with any text editor or HTML editor!</p>
        </div>
        
        <div class="footer">
            <p>Best regards,<br>
            Your Name</p>
            
            <p><small>This email was sent to {{TO}} on {{DATE}}</small></p>
        </div>
    </div>
</body>
</html>"""

        with open(template_path, 'w', encoding='utf-8') as f:
            f.write(default_template)

        print(f"✓ Created template file: {template_path}")
        print("  You can now edit this file to customize your emails!")
        return default_template

def attach_files(msg, attachment_paths):
    """Attach multiple files to the email"""
    attached_files = []

    for file_path in attachment_paths:
        if os.path.exists(file_path):
            try:
                # Open file in binary mode
                with open(file_path, "rb") as attachment:
                    # Instance of MIMEBase and named as part
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())

                # Encode file
                encoders.encode_base64(part)

                # Add header
                filename = os.path.basename(file_path)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {filename}'
                )

                # Attach the part to message
                msg.attach(part)
                attached_files.append(filename)

            except Exception as e:
                print(f"Warning: Could not attach file {file_path}: {str(e)}")
        else:
            print(f"Warning: Attachment file not found: {file_path}")

    return attached_files

def personalize_content(template, row_data):
    """Replace placeholders in template with actual data"""
    content = template

    # Standard replacements
    replacements = {
        '{{NAME}}': str(row_data.get('First Name', '')),
        '{{DATE}}': datetime.now().strftime("%B %d, %Y"),
        '{{COMPANY}}': str(row_data.get('Company', '')),
        '{{EMAIL}}': str(row_data.get('To', '')),
    }

    # Add any other columns as potential placeholders
    for key, value in row_data.items():
        placeholder = '{{' + key.upper().replace(' ', '_') + '}}'
        if placeholder not in replacements:
            replacements[placeholder] = str(value) if not pd.isna(value) else ''

    # Replace all placeholders
    for placeholder, value in replacements.items():
        content = content.replace(placeholder, value)

    return content

def send_personalized_email(row_data, config, template, test_mode=False):
    """Send personalized email with multiple recipients and attachments"""
    # Get email credentials
    sender_email = os.getenv('EMAIL_SENDER')
    sender_password = os.getenv('EMAIL_PASSWORD')
    smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
    smtp_port = int(os.getenv('SMTP_PORT', 587))

    if not sender_email or not sender_password:
        raise ValueError("Email credentials not found. Please set EMAIL_SENDER and EMAIL_PASSWORD in .env file")

    # Parse recipients
    to_emails = parse_email_list(row_data.get(config.COLUMNS['to'], ''))
    cc_emails = parse_email_list(row_data.get(config.COLUMNS['cc'], ''))
    bcc_emails = parse_email_list(row_data.get(config.COLUMNS['bcc'], ''))

    if not to_emails:
        return False, "No valid TO email addresses"

    # Create message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ', '.join(to_emails)

    if cc_emails:
        msg['Cc'] = ', '.join(cc_emails)

    # Use custom subject if provided, otherwise use default
    custom_subject = row_data.get(config.COLUMNS['custom_subject'], '')
    if custom_subject and not pd.isna(custom_subject) and str(custom_subject).strip():
        msg['Subject'] = personalize_content(str(custom_subject), row_data)
    else:
        msg['Subject'] = personalize_content(config.EMAIL_SUBJECT, row_data)

    # Personalize content
    personalized_content = personalize_content(template, row_data)
    msg.attach(MIMEText(personalized_content, 'html'))

    # Handle attachments
    attached_files = []

    # Global attachments (from attachments folder)
    if os.path.exists(config.ATTACHMENTS_FOLDER):
        global_attachments = [os.path.join(config.ATTACHMENTS_FOLDER, f)
                              for f in os.listdir(config.ATTACHMENTS_FOLDER)
                              if os.path.isfile(os.path.join(config.ATTACHMENTS_FOLDER, f))]
        attached_files.extend(attach_files(msg, global_attachments))

    # Specific attachments for this recipient
    specific_attachments = row_data.get(config.COLUMNS['attachments'], '')
    if specific_attachments and not pd.isna(specific_attachments):
        attachment_list = [a.strip() for a in str(specific_attachments).split(',')]
        attached_files.extend(attach_files(msg, attachment_list))

    # All recipients for sending
    all_recipients = to_emails + cc_emails + bcc_emails

    # Test mode - just print what would be sent
    if test_mode:
        print("\n" + "="*50)
        print("TEST MODE - Email Preview")
        print("="*50)
        print(f"From: {sender_email}")
        print(f"To: {', '.join(to_emails)}")
        if cc_emails:
            print(f"CC: {', '.join(cc_emails)}")
        if bcc_emails:
            print(f"BCC: {', '.join(bcc_emails)}")
        print(f"Subject: {msg['Subject']}")
        print(f"Attachments: {', '.join(attached_files) if attached_files else 'None'}")
        print("\nEmail Content Preview (first 500 chars):")
        print(personalized_content[:500] + "..." if len(personalized_content) > 500 else personalized_content)
        print("="*50 + "\n")
        return True, "Test mode - not sent"

    # Send the email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, all_recipients, msg.as_string())
        server.quit()

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status_message = f"Sent successfully"

        print(f"✓ Email sent to {row_data.get(config.COLUMNS['name'], 'Unknown')} ({', '.join(to_emails)})")
        if attached_files:
            print(f"  Attachments: {', '.join(attached_files)}")

        return True, status_message, timestamp

    except Exception as e:
        error_message = f"Failed: {str(e)}"
        print(f"✗ Failed to send email to {', '.join(to_emails)}: {str(e)}")
        return False, error_message, ""

def create_sample_files():
    """Create sample configuration and template files for users"""
    print("🔧 Setting up email sender files...\n")

    # Create sample .env file
    if not os.path.exists('.env'):
        with open('.env', 'w') as f:
            f.write("""# Email Configuration
EMAIL_SENDER=your_email@gmail.com
EMAIL_PASSWORD=your_app_password
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587

# For Gmail, use App Password instead of regular password
# Go to: https://myaccount.google.com/apppasswords
""")
        print("✓ Created .env file - Please update with your email credentials")
    else:
        print("✓ .env file already exists")

    # Create sample config file
    if not os.path.exists('email_config.json'):
        sample_config = {
            "EXCEL_FILE": "aarushi_main.xlsx",
            "EMAIL_SUBJECT": "Your Email Subject Here - Can use {{NAME}} placeholders",
            "TEMPLATE_FILE": "email_template.html",
            "ATTACHMENTS_FOLDER": "attachments",
            "DELAY_BETWEEN_EMAILS": 2,
            "MAX_EMAILS_PER_RUN": 100,
            "TEST_MODE": True,
            "_comment": "Set TEST_MODE to false when ready to send real emails"
        }
        with open('email_config.json', 'w') as f:
            json.dump(sample_config, f, indent=4)
        print("✓ Created email_config.json - Easy configuration file")
    else:
        print("✓ email_config.json already exists")

    # Create attachments folder
    if not os.path.exists('attachments'):
        os.makedirs('attachments')
        with open('attachments/README.txt', 'w') as f:
            f.write("Place any files here that you want to attach to ALL emails.\n")
            f.write("For specific attachments per recipient, use the 'Attachments' column in Excel.\n")
        print("✓ Created attachments folder - Place files here to attach to emails")
    else:
        print("✓ attachments folder already exists")

    print("\n📄 Files created successfully!")
    print("   1. Edit .env with your email credentials")
    print("   2. Edit email_template.html to customize your email")
    print("   3. Edit email_config.json for settings")
    print("   4. Prepare your Excel file with recipient data\n")

def create_sample_excel():
    """Create a sample Excel file with proper columns"""
    sample_data = {
        'First Name': ['John', 'Jane', 'Bob'],
        'To': ['john@example.com', 'jane@example.com,jane2@example.com', 'bob@example.com'],
        'CC': ['manager@example.com', 'boss@example.com,hr@example.com', ''],
        'BCC': ['', 'secret@example.com', ''],
        'Company': ['ABC Corp', 'XYZ Inc', 'Demo Ltd'],
        'Custom Subject': ['', 'Special invitation for {{NAME}}', ''],
        'Attachments': ['', 'report.pdf,presentation.pptx', 'invoice.pdf'],
        'Email Status': ['', '', ''],
        'Sent Timestamp': ['', '', '']
    }

    df = pd.DataFrame(sample_data)
    df.to_excel('sample_email_list.xlsx', index=False)
    print("Created sample_email_list.xlsx file")

def main():
    print("=== Enhanced Email Sender ===\n")

    # Create sample files if they don't exist
    create_sample_files()

    # Load configuration
    config = EmailConfig()
    external_config = read_config_file()

    # Update config with external settings
    for key, value in external_config.items():
        if hasattr(config, key):
            setattr(config, key, value)

    # Load email template from separate HTML file
    email_template = read_email_template(config.TEMPLATE_FILE)

    print(f"\n📧 Email Configuration:")
    print(f"   Template: {config.TEMPLATE_FILE}")
    print(f"   Excel file: {config.EXCEL_FILE}")
    print(f"   Attachments folder: {config.ATTACHMENTS_FOLDER}")
    print(f"   Subject: {config.EMAIL_SUBJECT[:50]}..." if len(config.EMAIL_SUBJECT) > 50 else f"   Subject: {config.EMAIL_SUBJECT}")
    print()

    # Check if Excel file exists
    if not os.path.exists(config.EXCEL_FILE):
        print(f"Error: Excel file '{config.EXCEL_FILE}' not found!")
        print("\nWould you like to create a sample Excel file? (yes/no): ", end='')
        if input().lower().startswith('y'):
            create_sample_excel()
        return

    # Read client data
    try:
        df = read_client_data(config.EXCEL_FILE)

        # Validate columns
        missing_columns = []
        for col_name in config.COLUMNS.values():
            if col_name not in df.columns and col_name != 'BCC' and col_name != 'Attachments' and col_name != 'Custom Subject':
                missing_columns.append(col_name)

        if missing_columns:
            print(f"Error: Missing required columns: {', '.join(missing_columns)}")
            print(f"Available columns: {', '.join(df.columns)}")
            return

        # Add missing optional columns
        for col in ['BCC', 'Attachments', 'Custom Subject', 'Sent Timestamp']:
            if config.COLUMNS.get(col.lower(), col) not in df.columns:
                df[config.COLUMNS.get(col.lower(), col)] = ''

        # Test mode check
        if config.TEST_MODE:
            print("*** RUNNING IN TEST MODE - NO EMAILS WILL BE SENT ***\n")
            print("Set TEST_MODE to False in config to send actual emails\n")

        # Summary
        total_rows = len(df)
        pending_emails = df[df[config.COLUMNS['status']].isna() | (df[config.COLUMNS['status']] == '')].shape[0]

        print(f"Total recipients: {total_rows}")
        print(f"Pending emails: {pending_emails}")
        print(f"Already sent: {total_rows - pending_emails}\n")

        if pending_emails == 0:
            print("No pending emails to send.")
            return

        # Confirmation
        if not config.TEST_MODE:
            print(f"Ready to send {min(pending_emails, config.MAX_EMAILS_PER_RUN)} emails.")
            print("Continue? (yes/no): ", end='')
            if not input().lower().startswith('y'):
                print("Cancelled.")
                return

        print("\nStarting email campaign...\n")

        # Process emails
        successful_emails = 0
        failed_emails = 0
        emails_sent_this_run = 0

        for index, row in df.iterrows():
            # Check if already sent
            current_status = row[config.COLUMNS['status']]
            if not pd.isna(current_status) and str(current_status).strip():
                continue

            # Check max emails per run
            if emails_sent_this_run >= config.MAX_EMAILS_PER_RUN:
                print(f"\nReached maximum emails per run ({config.MAX_EMAILS_PER_RUN}). Stopping.")
                break

            # Send email
            result = send_personalized_email(row, config, email_template, config.TEST_MODE)

            if len(result) == 3:
                success, status_message, timestamp = result
            else:
                success, status_message = result
                timestamp = ""

            # Update DataFrame
            df.at[index, config.COLUMNS['status']] = status_message
            if timestamp:
                df.at[index, config.COLUMNS['timestamp']] = timestamp

            if success:
                successful_emails += 1
            else:
                failed_emails += 1

            emails_sent_this_run += 1

            # Delay between emails
            if emails_sent_this_run < config.MAX_EMAILS_PER_RUN and index < len(df) - 1:
                time.sleep(config.DELAY_BETWEEN_EMAILS)

        # Save updated DataFrame
        df.to_excel(config.EXCEL_FILE, index=False)
        print(f"\n✓ Updated Excel file saved: {config.EXCEL_FILE}")

        # Print summary
        print("\n" + "="*50)
        print("EMAIL CAMPAIGN SUMMARY")
        print("="*50)
        print(f"Total recipients in file: {total_rows}")
        print(f"Emails processed this run: {emails_sent_this_run}")
        print(f"Successful: {successful_emails}")
        print(f"Failed: {failed_emails}")
        print(f"Remaining: {pending_emails - emails_sent_this_run}")

        if config.TEST_MODE:
            print("\n*** This was a TEST RUN - no actual emails were sent ***")
            print("Set TEST_MODE to False in config to send real emails")

    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()

def read_client_data(file_path):
    """Read client data from Excel or CSV file"""
    if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
        return pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    else:
        raise ValueError("Unsupported file format. Please use Excel (.xlsx, .xls) or CSV (.csv)")

if __name__ == "__main__":
    main()