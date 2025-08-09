import schedule
import time
import subprocess

def job():
    subprocess.run(["python", "agent.py"])

schedule.every().day.at("08:00").do(job)

while True:
    schedule.run_pending()
    time.sleep(60)
