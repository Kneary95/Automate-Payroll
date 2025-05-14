import os
import glob
import win32com.client as win32

def find_latest_report():
    files = glob.glob("*_Payroll.xlsx")
    if not files:
        print("❌ No payroll files found.")
        return None
    return max(files, key=os.path.getctime)

def send_email_with_attachment(filepath, subject, body, recipients):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  # 0: Mail item
    mail.To = "; ".join(recipients)
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(os.path.abspath(filepath))
    
    mail.Send()
    print(f"📤 Email sent with attachment: {filepath}")

if __name__ == "__main__":
    latest_file = find_latest_report()
    if latest_file:
        subject = "Biweekly Payroll Comparison Report"
        body = (
            "Hi Team,\n\n"
            "Please find attached the latest payroll comparison report.\n\n"
            "Let me know if you have any questions.\n\n"
            "Best,\nKevin"
        )
        recipients = ["kevin.neary@matrixaba.com"]  # <-- Update these
        send_email_with_attachment(latest_file, subject, body, recipients)