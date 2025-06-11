# Email Automation Tool

A powerful Python-based email automation tool that allows you to send personalized bulk emails using HTML templates and Excel/CSV data sources. This tool is perfect for sending newsletters, invitations, or any other type of bulk email communication while maintaining personalization and tracking capabilities.

## Features

- 📧 Send personalized bulk emails using HTML templates
- 📊 Read recipient data from Excel (.xlsx) or CSV files
- 📝 Support for HTML email templates with dynamic placeholders
- 📎 Attach multiple files to emails
- 📈 Track email sending status and timestamps
- 🔄 Support for CC and BCC recipients
- ⚡ Configurable delay between emails to avoid spam filters
- 🛡️ Secure SMTP connection with TLS support
- 🧪 Test mode for previewing emails without sending
- 📋 Detailed logging and status tracking

## Prerequisites

- Python 3.6 or higher
- Required Python packages (install using `pip install -r requirements.txt`):
  - pandas
  - python-dotenv
  - openpyxl (for Excel file support)

## Installation

1. Clone this repository:
```bash
git clone https://github.com/TheAniketTayade/mail_gun.git
cd mail_gun
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the project root with your email credentials:
```env
EMAIL_SENDER=your.email@example.com
EMAIL_PASSWORD=your_app_password
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
```

## Configuration

1. Prepare your recipient data in an Excel or CSV file with the following columns:
   - First Name
   - To (recipient email)
   - CC (optional)
   - BCC (optional)
   - Email Status (will be updated automatically)
   - Sent Timestamp (will be updated automatically)

2. Create or modify your HTML email template in `email_template.html`. Available placeholders:
   - `{{NAME}}` - Recipient's first name
   - `{{DATE}}` - Current date
   - `{{TO}}` - Recipient's email
   - `{{COMPANY}}` - Company name (if column exists)
   - Any other column from your Excel file (in UPPERCASE)

3. (Optional) Configure additional settings in `email_config.json`:
```json
{
    "EXCEL_FILE": "your_data.xlsx",
    "EMAIL_SUBJECT": "Your Email Subject",
    "TEMPLATE_FILE": "email_template.html",
    "ATTACHMENTS_FOLDER": "attachments",
    "DELAY_BETWEEN_EMAILS": 2,
    "MAX_EMAILS_PER_RUN": 100,
    "TEST_MODE": false
}
```

## Usage

1. Place your recipient data file in the project directory
2. Update the email template in `email_template.html`
3. Run the script:
```bash
python main.py
```

For test mode (preview emails without sending):
```bash
python main.py --test
```

## File Structure

```
mail_gun/
├── main.py              # Main script for email automation
├── email_sender.py      # Core email sending functionality
├── email_template.html  # HTML template for emails
├── email_config.json    # Configuration file
├── .env                 # Environment variables (create this)
├── requirements.txt     # Python dependencies
└── attachments/         # Folder for email attachments
```

## Security Notes

- Never commit your `.env` file or expose your email credentials
- Use app passwords instead of your main email password
- Keep your recipient data secure and comply with data protection regulations

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

If you encounter any issues or have questions, please open an issue in the GitHub repository.