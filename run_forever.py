import time
import subprocess
from datetime import datetime, timedelta
import pytz

# Set timezone to Singapore
SG_TIMEZONE = pytz.timezone("Asia/Singapore")

def wait_until_next_10am():
    now = datetime.now(SG_TIMEZONE)
    today_10am = now.replace(hour=10, minute=0, second=0, microsecond=0)

    if now >= today_10am:
        # If it's already past 10 AM today, wait until 10 AM tomorrow
        next_10am = today_10am + timedelta(days=1)
    else:
        # If it's before 10 AM today
        next_10am = today_10am

    seconds_to_wait = (next_10am - now).total_seconds()
    print(f"🕒 Waiting {int(seconds_to_wait)} seconds until next 10 AM (Singapore)...")
    time.sleep(seconds_to_wait)

while True:
    wait_until_next_10am()
    print(f"\n🚀 Running main.py at {datetime.now(SG_TIMEZONE)}")
    subprocess.run(["python", "main.py"])
