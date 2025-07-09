from flask import Flask, render_template, request, send_from_directory
import os
import csv
import statistics
import google.generativeai as genai
import base64
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import sqlite3
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# ✅ Your Gemini API Key
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs("downloads", exist_ok=True)

SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# ✅ Gmail Authentication Setup
def get_gmail_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def send_email(recipient_email, subject, body):
    if os.getenv("DISABLE_EMAIL", "false").lower() == "true":
        print(f"[DEBUG] Skipped sending email to {recipient_email} (DISABLE_EMAIL enabled)")
        return
    service = get_gmail_service()
    message = MIMEMultipart()
    message['to'] = recipient_email
    message['subject'] = subject
    message.attach(MIMEText(body, 'plain'))

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    send_message = {'raw': raw_message}
    try:
        service = build('gmail', 'v1', credentials=service)
        service.users().messages().send(userId="me", body=send_message).execute()
        print(f"Email sent to {recipient_email}")
    except Exception as error:
        print(f"An error occurred while sending email to {recipient_email}: {error}")

# ✅ Initialize Database
def init_db():
    conn = sqlite3.connect("feedback.db")
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS feedbacks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            email TEXT,
            physics INTEGER,
            chemistry INTEGER,
            maths INTEGER,
            cs INTEGER,
            english INTEGER,
            feedback TEXT
        )
    ''')
    conn.commit()
    conn.close()

# ✅ Insert into DB
def store_feedback(name, email, phy, chem, maths, cs, eng, feedback):
    conn = sqlite3.connect("feedback.db")
    c = conn.cursor()
    c.execute("INSERT INTO feedbacks (name, email, physics, chemistry, maths, cs, english, feedback) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
              (name, email, phy, chem, maths, cs, eng, feedback))
    conn.commit()
    conn.close()

# ✅ Extract stats from data
def compute_subject_stats(data):
    stats = {}
    subjects = ['Physics', 'Chemistry', 'Maths', 'CS', 'English']
    for subject in subjects:
        marks = [int(row[subject]) for row in data]
        stats[subject] = {
            "avg": round(statistics.mean(marks), 2),
            "max": max(marks)
        }
    return stats

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/download/<filename>")
def download_xlsx(filename):
    return send_from_directory("downloads", filename, as_attachment=True)

@app.route("/generate", methods=["POST"])
def generate():
    teacher_name = request.form['teacher_name']
    school_name = request.form['school_name']
    class_name = request.form['class_name']
    max_marks = request.form['max_marks']
    uploaded_file = request.files['csv_file']

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], uploaded_file.filename)
    uploaded_file.save(filepath)

    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(filepath)
    elif ext in [".xlsx", ".xls"]:
        df = pd.read_excel(filepath)
    else:
        return "Unsupported file type!", 400

    data = df.to_dict(orient='records')
    stats = compute_subject_stats(data)
    init_db()

    excel_data = []

    for student in data:
        student_name = student['Name']
        email = student['Email']
        phy = student['Physics']
        chem = student['Chemistry']
        maths = student['Maths']
        cs = student['CS']
        eng = student['English']

        prompt = f"""
        You are a teacher writing a personalized academic feedback email for a student based on their marks in five subjects, along with class performance statistics.

        📘 Student Details:
        - Name: {student_name}
        - Class: {class_name}
        - Maximum Marks per Subject: {max_marks}

        🎯 Student's Performance:
        - Physics: {phy} (Class Avg: {stats['Physics']['avg']}, Max: {stats['Physics']['max']})
        - Chemistry: {chem} (Class Avg: {stats['Chemistry']['avg']}, Max: {stats['Chemistry']['max']})
        - Mathematics: {maths} (Class Avg: {stats['Maths']['avg']}, Max: {stats['Maths']['max']})
        - Computer Science: {cs} (Class Avg: {stats['CS']['avg']}, Max: {stats['CS']['max']})
        - English: {eng} (Class Avg: {stats['English']['avg']}, Max: {stats['English']['max']})

        📜 Instructions for Feedback Generation:
        1. Use a polite and motivating tone.
        2. For subjects where the student scored above class average, praise them.
        3. For subjects below 85% of the maximum marks, suggest improvements and helpful YouTube links (e.g., "Class 12 CS Basics").
        4. Keep the message personalized and under 200 words.

        Teacher's Name: {teacher_name}

        Provide with a final email, that I will directly and automatically send to students via a program. Do not add any such things which show that it is being generated from AI. Only give me final result no such lines like  \"Here is your email\" or anything.
        Just provide email body content do not start with subject.
        """

        try:
            response = model.generate_content(prompt)
            feedback = response.text.strip()
            send_email(email, "Academic Performance Feedback", feedback)
            store_feedback(student_name, email, phy, chem, maths, cs, eng, feedback)
            excel_data.append({
                'Name': student_name,
                'Email': email,
                'Physics': phy,
                'Chemistry': chem,
                'Maths': maths,
                'CS': cs,
                'English': eng,
                'Feedback': feedback
            })

        except Exception as e:
            print(f"Error generating or sending feedback for {student_name}: {str(e)}")

    df_result = pd.DataFrame(excel_data)
    filename = f"feedback_summary_{teacher_name.replace(' ', '_')}.xlsx"
    download_path = os.path.join("downloads", filename)
    df_result.to_excel(download_path, index=False)

    return render_template("index.html",
                           email="Feedback emails sent and stored in DB.",
                           download_link=f"/download/{filename}")

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))  # Render sets PORT env var
    app.run(host='0.0.0.0', port=port)

