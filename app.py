from flask import Flask, request, render_template, send_file, redirect, url_for
import openpyxl
import threading
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

class WhatsAppAutomation:

    def __init__(self):
        self.driver = None
        self.wait = None

    def start_driver(self):
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        service = Service("/usr/local/bin/chromedriver")
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.wait = WebDriverWait(self.driver, 10)

    def login(self):
        self.driver.get("https://web.whatsapp.com/")
        return "Opened WhatsApp Web. Please login."

    def send_messages(self, filepath, language="en"):
        self.start_driver()
        wb = openpyxl.load_workbook(filepath)
        sheet = wb.active

        wb_sent = openpyxl.Workbook()
        sheet_sent = wb_sent.active
        sheet_sent.append(["Phone Number", "Message", "Status"])

        wb_unsent = openpyxl.Workbook()
        sheet_unsent = wb_unsent.active
        for row in sheet.iter_rows(values_only=True):
            sheet_unsent.append(row)

        unsent_row = 2

        for row in sheet.iter_rows(min_row=2, values_only=True):
            phone_number, message = str(row[0]), str(row[1])

            try:
                self.driver.get(f"https://web.whatsapp.com/send?phone={phone_number}")
                time.sleep(5)  # Give time for the page to load

                message_box_xpath = "//div[@aria-label='Type a message']" if language == "en" else "//div[@aria-label='اكتب رسالة']"
                message_box = self.wait.until(EC.presence_of_element_located((By.XPATH, message_box_xpath)))
                message_box.send_keys(message + Keys.ENTER)

                # Wait for the message to be sent and determine its status
                status = "Unknown"
                icons = [
                    ("//span[@data-icon='msg-check']", "Sent"),
                    ("//span[@data-icon='msg-dblcheck']", "Delivered"),
                    ("//span[@data-icon='msg-dblcheck-ack']", "Read")
                ]

                for icon, state in icons:
                    try:
                        self.wait.until(EC.presence_of_element_located((By.XPATH, icon)))
                        status = state
                        break
                    except:
                        continue

                sheet_sent.append([phone_number, message, status])
                sheet_unsent.delete_rows(unsent_row)

            except Exception as e:
                print(f"Error for {phone_number}: {str(e)}")
                unsent_row += 1

        wb_sent.save(os.path.join(app.config['UPLOAD_FOLDER'], "sent_messages.xlsx"))
        wb_unsent.save(os.path.join(app.config['UPLOAD_FOLDER'], "unsent_messages.xlsx"))

        self.driver.quit()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        automation_thread = threading.Thread(target=start_whatsapp_automation, args=(file_path,))
        automation_thread.start()
        return redirect(url_for('status'))

@app.route('/status')
def status():
    # Here you can add code to display the current status
    return "Sending messages..."

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename))

def start_whatsapp_automation(filepath):
    automation = WhatsAppAutomation()
    automation.send_messages(filepath)

if __name__ == "__main__":
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(host='0.0.0.0', port=5000)
