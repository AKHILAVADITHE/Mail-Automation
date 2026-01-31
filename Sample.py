import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
from datetime import datetime
import io

CONFIG = {
    "EMAIL_SENDER": "vaditheakhila162@gmail.com",
    "EMAIL_PASSWORD": "sjek zunx kxvf alpc",
    "SMTP_SERVER": "smtp.gmail.com",
    "SMTP_PORT": 587,
    "TO_EMAIL": "vaditheakhila162@gmail.com",
    "CC_EMAILS": ["akhilavadithe123@gmail.com"],
    "VIRAL_VIDEOS_CSV": "viral_videos_data.csv",
    "MARKET_SHARE_CSV": "market_share_data.csv",
    "PRODUCTIVITY_CSV": "productivity_data.csv",
    "ENGAGEMENT_CSV": "engagement_metrics.csv"
}

def load_csv_data(file_path, report_type):
    try:
        if not os.path.exists(file_path):
            print(f" CSV file not found: {file_path}")
            return None

        df = pd.read_csv(file_path)

        if report_type == "viral_videos" and 'published_date' in df.columns:
            df['published_date'] = pd.to_datetime(df['published_date'])

        print(f"Loaded {len(df)} rows from {os.path.basename(file_path)}")
        return df

    except Exception as e:
        print(f" Error loading CSV {file_path}: {str(e)}")
        return None

def format_views(value):
    if pd.isna(value) or value == 0:
        return "0"
    elif value >= 1000000000:
        return f"{value/1000000000:.1f}B"
    elif value >= 1000000:
        return f"{value/1000000:.1f}M"
    elif value >= 1000:
        return f"{value/1000:.1f}K"
    else:
        return f"{int(value):,}"

def send_gmail(subject, html_content, to_email, cc_emails=None, attachment=None, attachment_name=""):
    try:
        msg = MIMEMultipart('mixed')
        msg['From'] = CONFIG["EMAIL_SENDER"]
        msg['To'] = to_email

        if cc_emails:
            msg['Cc'] = ", ".join(cc_emails)
            all_recipients = [to_email] + cc_emails
        else:
            all_recipients = [to_email]

        msg['Subject'] = subject

        html_part = MIMEText(html_content, 'html')
        msg.attach(html_part)

        if attachment:
            excel_part = MIMEApplication(attachment)
            excel_part.add_header(
                'Content-Disposition',
                'attachment',
                filename=attachment_name
            )
            msg.attach(excel_part)

        server = smtplib.SMTP(CONFIG["SMTP_SERVER"], CONFIG["SMTP_PORT"])
        server.starttls()
        server.login(CONFIG["EMAIL_SENDER"], CONFIG["EMAIL_PASSWORD"])
        server.send_message(msg, to_addrs=all_recipients)
        server.quit()

        print(f" Email sent successfully to: {', '.join(all_recipients)}")
        return True

    except Exception as e:
        print(f"Failed to send email: {str(e)}")
        return False

def create_simple_table(data, title, report_type):
    if data is None or data.empty:
        return f"<p>No {title.lower()} data available.</p>"

    display_data = data.head(20)

    table_html = f"""
    <div style="margin: 25px 0; background: #ffffff; border-radius: 12px; padding: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); border: 1px solid #e2e8f0;">
        <h3 style="color: #1a202c; margin: 0 0 20px 0; text-align: center; font-size: 18px; font-weight: bold;">
            📊 {title}
        </h3>
        <div style="width: 100%; overflow-x: auto;">
            <table style="border-collapse: collapse; width: 100%; border: 1px solid #cbd5e1;">
                <thead>
                    <tr style="background: #f8fafc; color: #374151;">
    """

    for col in display_data.columns:
        table_html += f'<th style="padding: 12px; border: 1px solid #cbd5e1; font-weight: bold; text-align: left;">{col.replace("_", " ").title()}</th>'

    table_html += """
                    </tr>
                </thead>
                <tbody>
    """

    for i, (_, row) in enumerate(display_data.iterrows()):
        bg_color = '#f9fafb' if i % 2 == 0 else '#ffffff'
        table_html += f'<tr style="background-color: {bg_color};">'

        for col in display_data.columns:
            value = row[col]

            if 'view' in col.lower() and 'count' in col.lower():
                formatted_value = format_views(value)
            elif 'date' in col.lower():
                if pd.notna(value):
                    formatted_value = pd.to_datetime(value).strftime('%Y-%m-%d') if not isinstance(value, str) else value
                else:
                    formatted_value = 'N/A'
            else:
                formatted_value = str(value) if pd.notna(value) else 'N/A'

            table_html += f'<td style="padding: 10px; border: 1px solid #cbd5e1;">{formatted_value}</td>'

        table_html += '</tr>'

    table_html += """
                </tbody>
            </table>
        </div>
        <p style="color: #6b7280; font-size: 12px; text-align: center; margin: 15px 0 0 0;">
            📎 Complete dataset available in attached Excel file
        </p>
    </div>
    """

    return table_html

def create_excel_attachment(data, sheet_name):
    if data is None or data.empty:
        return None

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        data.to_excel(writer, sheet_name=sheet_name, index=False)

        worksheet = writer.sheets[sheet_name]

        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass

            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

    output.seek(0)
    return output.read()

def create_email_html(content_table, title, report_type):
    current_date = datetime.now().strftime("%B %d, %Y")

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="margin: 0; padding: 0; font-family: Arial, sans-serif; background-color: #f1f5f9;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #f1f5f9;">
            <tr>
                <td align="center" style="padding: 30px 15px;">
                    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width: 900px; background-color: #ffffff; border-radius: 12px; box-shadow: 0 10px 25px rgba(0,0,0,0.1);">
                        <tr>
                            <td style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 35px 30px; border-radius: 12px 12px 0 0;">
                                <h1 style="margin: 0; color: #ffffff; font-size: 24px; text-align: center;">
                                    📊 {title}
                                </h1>
                                <p style="margin: 10px 0 0 0; color: #e2e8f0; font-size: 14px; text-align: center;">
                                    Generated on {current_date}
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td style="padding: 20px 30px; background-color: #ffffff;">
                                {content_table}
                            </td>
                        </tr>
                        <tr>
                            <td style="background: #f8fafc; padding: 25px 30px; text-align: center; border-radius: 0 0 12px 12px;">
                                <p style="margin: 0; color: #64748b; font-size: 14px;">
                                    Best Regards,<br>
                                    <strong style="color: #3b82f6;">Analytics Team</strong>
                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """

def main():
    print("🚀 Starting Standalone Mail Automation...")
    print("=" * 50)

    reports = {
        "viral_videos": {
            "title": "Daily Viral Videos Report",
            "csv_file": CONFIG["VIRAL_VIDEOS_CSV"],
            "sheet_name": "Viral Videos"
        },
        "market_share": {
            "title": "Market Share Analysis",
            "csv_file": CONFIG["MARKET_SHARE_CSV"],
            "sheet_name": "Market Share"
        },
        "productivity": {
            "title": "Productivity Dashboard",
            "csv_file": CONFIG["PRODUCTIVITY_CSV"],
            "sheet_name": "Productivity"
        },
        "engagement": {
            "title": "Engagement Report",
            "csv_file": CONFIG["ENGAGEMENT_CSV"],
            "sheet_name": "Engagement"
        }
    }

    for report_type, config in reports.items():
        print(f"\n Processing {config['title']}...")

        data = load_csv_data(config["csv_file"], report_type)
        if data is None:
            print(f" Skipping {config['title']} - no data available")
            continue

        content_table = create_simple_table(data, config['title'], report_type)
        html_content = create_email_html(content_table, config['title'], report_type)
        excel_attachment = create_excel_attachment(data, config['sheet_name'])

        subject = f"{config['title']} - {datetime.now().strftime('%B %d, %Y')}"
        attachment_name = f"{report_type}_report_{datetime.now().strftime('%Y%m%d')}.xlsx"

        print(f" Sending {config['title']} email...")

        success = send_gmail(
            subject=subject,
            html_content=html_content,
            to_email=CONFIG["TO_EMAIL"],
            cc_emails=CONFIG["CC_EMAILS"],
            attachment=excel_attachment,
            attachment_name=attachment_name
        )

        if success:
            print(f" {config['title']} sent successfully!")
        else:
            print(f" Failed to send {config['title']}")

    print("\n" + "=" * 50)
    print("🎉 Mail automation completed!")

if __name__ == "__main__":
    main()
