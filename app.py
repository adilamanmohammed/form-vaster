from flask import Flask, render_template, request
import os
import pandas as pd
from datetime import datetime
import uuid

app = Flask(__name__)

# Path to store submissions
SUBMISSION_FOLDER = 'submissions'
if not os.path.exists(SUBMISSION_FOLDER):
    os.makedirs(SUBMISSION_FOLDER)

# Path to centralized Excel file
EXCEL_FILE = os.path.join(SUBMISSION_FOLDER, 'employee_submissions.xlsx')

# Ensure the Excel file exists
if not os.path.exists(EXCEL_FILE):
    pd.DataFrame(columns=['APID', 'Name', 'Email', 'Details', 'Submission Time']).to_excel(EXCEL_FILE, index=False)

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form['name']
    email = request.form['email']
    details = request.form['details']
    files = request.files.getlist('files')
    
    # Generate unique APID
    apid = f"APID{uuid.uuid4().hex[:6].upper()}"

    # Create folder for submission
    folder_name = f"{apid}_{name.replace(' ', '_')}"
    folder_path = os.path.join(SUBMISSION_FOLDER, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    
    # Save uploaded files
    for file in files:
        file.save(os.path.join(folder_path, file.filename))
    
    # Save submission data in Excel
    submission_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    new_data = pd.DataFrame([[apid, name, email, details, submission_time]], 
                             columns=['APID', 'Name', 'Email', 'Details', 'Submission Time'])
    existing_data = pd.read_excel(EXCEL_FILE)
    updated_data = pd.concat([existing_data, new_data], ignore_index=True)
    updated_data.to_excel(EXCEL_FILE, index=False)
    
    return f"Thanks for submitting! Your Application ID is {apid}."

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)