import pandas as pd
from datetime import datetime
import win32com.client as win32

def check_expiry():
    # Load the Excel file
    df = pd.read_excel("data/certificate_expiry.xlsx")
    
    # Convert dates and calculate Days Remaining
    df["Expiry Date"] = pd.to_datetime(df["Expiry Date"]).dt.date
    df["Days Remaining"] = (df["Expiry Date"] - datetime.today().date()).apply(lambda x: x.days)
    
    # Notify only if today matches the 'How many days before' exactly
    df["Notify"] = df["Days Remaining"] == df["How many days before"]
    
    print(f"\n🕒 Today: {datetime.today().date()}")
    print("\n📋 Certificate Status:\n")
    print(df[["Email Subject", "Expiry Date", "How many days before", "Days Remaining", "Notify"]])

    expiring = df[df["Notify"] == True]
    if not expiring.empty:
        print("\n⚠️ Expiring Today According to Alert Window:\n")
        print(expiring[["Email Subject", "Expiry Date", "Days Remaining"]])
        return expiring
    else:
        print("\n✅ No items expiring exactly today based on 'How many days before'.")
        return None

def send_outlook_email(row):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.Subject = row["Email Subject"]
    mail.To = row["To Email Address"]

    if not pd.isna(row["CC Email Address"]):
        mail.CC = row["CC Email Address"]

    # Format content for HTML
    email_content = row["Email Content"].replace('\n', '<br>')

    mail.HTMLBody = f"""
    <p>Dear Team,</p>
    <p>{email_content}</p>
    <p><b>Expiry Date:</b> {row['Expiry Date']}<br>
    <b>Days Remaining:</b> {row['Days Remaining']}</p>
    <p>Regards,<br>Automation Bot</p>
    """

    mail.Send()
    print(f"📧 Email sent: {row['Email Subject']}")

# === MAIN SCRIPT ENTRYPOINT ===
if __name__ == "__main__":
    expiring_df = check_expiry()
    if expiring_df is not None:
        for _, row in expiring_df.iterrows():
            send_outlook_email(row)
