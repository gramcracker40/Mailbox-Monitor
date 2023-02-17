import subprocess
import time


while True:
    run_check = subprocess.run(['python', 'main.py'])
    time.sleep(30)

