import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import os
import shutil
import re
import textract
import PyPDF2
import docx
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Extract text from file
def extract_text(filepath):
    try:
        if filepath.endswith('.pdf'):
            with open(filepath, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                return ' '.join(page.extract_text() or '' for page in reader.pages)
        elif filepath.endswith('.docx'):
            doc = docx.Document(filepath)
            return ' '.join(para.text for para in doc.paragraphs)
    except:
        return ""
    return ""

# Extract Name, Email
def extract_details(text):
    email = re.findall(r'\b[\w.-]+?@\w+?\.\w+?\b', text)
    lines = text.strip().split('\n')
    name = lines[0] if lines else 'Candidate'
    return name.strip().title(), email[0] if email else None

# Experience Matching
def experience_matches(text, exp):
    patterns = [
        rf'{exp}\s*\+?\s*(years?|yrs?)',
        rf'over\s*{exp}\s*(years?|yrs?)',
        rf'at\s*least\s*{exp}\s*(years?|yrs?)',
        rf'minimum\s*{exp}\s*(years?|yrs?)',
    ]
    return any(re.search(pat, text.lower()) for pat in patterns)

# Suggestions
def generate_suggestions(missing_skills):
    if not missing_skills:
        return "Try improving resume clarity and keyword optimization."
    tips = [f"Learn {skill.title()} from wscube tech.com/Coursera/YouTube." for skill in missing_skills]
    return "\n".join(tips)

# Send Email
def send_email(to_email, subject, body):
    try:
        sender_email = email_entry.get()
        sender_pass = pass_entry.get()

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_pass)
            server.send_message(msg)
    except Exception as e:
        print(f"Email failed for {to_email}: {e}")

def shortlist_candidates():
    folder = folder_entry.get()
    skills = skills_entry.get().split(',')
    education = education_entry.get().lower()
    experience = experience_entry.get()

    if not folder or not skills or not experience or not education:
        messagebox.showwarning("Input Missing", "Please fill all fields!")
        return

    try:
        experience = int(experience)
    except ValueError:
        messagebox.showwarning("Invalid Experience", "Experience must be a number!")
        return

    shortlisted_tree.delete(*shortlisted_tree.get_children())
    not_shortlisted_tree.delete(*not_shortlisted_tree.get_children())

    short_dir = os.path.join(folder, 'Shortlisted')
    not_short_dir = os.path.join(folder, 'Not_Shortlisted')
    os.makedirs(short_dir, exist_ok=True)
    os.makedirs(not_short_dir, exist_ok=True)

    for file in os.listdir(folder):
        filepath = os.path.join(folder, file)
        if not file.endswith(('.pdf', '.docx')):
            continue

        text = extract_text(filepath).lower()
        name, email_id = extract_details(text)

        skill_match = [skill.strip().lower() for skill in skills if skill.strip().lower() in text]
        missing_skills = [skill.strip().lower() for skill in skills if skill.strip().lower() not in skill_match]
        exp_match = experience_matches(text, experience)
        edu_match = education in text

        if len(skill_match) >= len(skills)//2 and exp_match and edu_match:
            shutil.copy(filepath, short_dir)
            shortlisted_tree.insert('', 'end', values=(name, email_id))
            if email_id:
                subject = "Congratulations! You are Selected for Interview"
                body = f"Hi {name},\n\nCongratulations! You are shortlisted for interview.\n\nBest Wishes,\nHR Team"
                send_email(email_id, subject, body)
        else:
            shutil.copy(filepath, not_short_dir)
            not_shortlisted_tree.insert('', 'end', values=(name, email_id))
            if email_id:
                subject = "Regarding Your Application Status"
                body = f"Hi {name},\n\nUnfortunately, you are not shortlisted due to missing skills.\n\nMissing Skills: {', '.join(missing_skills)}\n\nSuggestions:\n{generate_suggestions(missing_skills)}\n\nBest Wishes,\nHR Team"
                send_email(email_id, subject, body)

    messagebox.showinfo("Process Completed", "Candidate Shortlisting Done!")

# GUI Setup
root = tk.Tk()
root.title("Resume Shortlister with Email")
root.geometry("850x600")

tk.Label(root, text="Folder Path:").pack()
folder_entry = tk.Entry(root, width=80)
folder_entry.pack()

tk.Button(root, text="Browse", command=lambda: folder_entry.insert(0, filedialog.askdirectory())).pack()

tk.Label(root, text="Required Skills (comma separated):").pack()
skills_entry = tk.Entry(root, width=80)
skills_entry.pack()

tk.Label(root, text="Education (e.g., b.tech, mca, bsc):").pack()
education_entry = tk.Entry(root, width=80)
education_entry.pack()

tk.Label(root, text="Minimum Experience (years):").pack()
experience_entry = tk.Entry(root, width=80)
experience_entry.pack()

tk.Label(root, text="Your Gmail:").pack()
email_entry = tk.Entry(root, width=50)
email_entry.pack()

tk.Label(root, text="Your App Password:").pack()
pass_entry = tk.Entry(root, width=50, show="*")
pass_entry.pack()

tk.Button(root, text="Start Shortlisting", command=shortlist_candidates, bg='green', fg='white').pack(pady=10)

frame = tk.Frame(root)
frame.pack(pady=10)

tk.Label(frame, text="Shortlisted Candidates").grid(row=0, column=0, padx=20)
shortlisted_tree = ttk.Treeview(frame, columns=("Name", "Email"), show='headings', height=15)
shortlisted_tree.heading('Name', text='Name')
shortlisted_tree.heading('Email', text='Email')
shortlisted_tree.grid(row=1, column=0, padx=20)

tk.Label(frame, text="Not Shortlisted Candidates").grid(row=0, column=1, padx=20)
not_shortlisted_tree = ttk.Treeview(frame, columns=("Name", "Email"), show='headings', height=15)
not_shortlisted_tree.heading('Name', text='Name')
not_shortlisted_tree.heading('Email', text='Email')
not_shortlisted_tree.grid(row=1, column=1, padx=20)

root.mainloop()
