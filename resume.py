import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import shutil
import re
import textract
import PyPDF2
import docx
import pandas as pd

# Extract text from any file type
def extract_text(filepath):
    text = ""
    try:
        if filepath.endswith('.pdf'):
            with open(filepath, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                text = ' '.join(page.extract_text() or '' for page in reader.pages)
        elif filepath.endswith('.docx'):
            doc = docx.Document(filepath)
            text = ' '.join(para.text for para in doc.paragraphs)
        else:
            raise ValueError("Unsupported file format")
    except Exception as e:
        text = f"Error reading file: {e}"
    return text

# Extract Name, Email, Phone
def extract_details(text):
    email = re.findall(r'\b[\w.-]+?@\w+?\.\w+?\b', text)
    phone = re.findall(r'(\+?\d{1,3}[\s-]?)?\d{10}', text)
    lines = text.strip().split('\n')
    name = lines[0] if lines else 'Not Found'
    return name.strip(), email[0] if email else 'Not Found', phone[0] if phone else 'Not Found'

# Experience Matching
def experience_matches(text, exp):
    patterns = [
        rf'{exp}\s*\+?\s*(years?|yrs?)',
        rf'over\s*{exp}\s*(years?|yrs?)',
        rf'at\s*least\s*{exp}\s*(years?|yrs?)',
        rf'minimum\s*{exp}\s*(years?|yrs?)',
        rf'{exp}\s*-\s*(years?|yrs?)',
        rf'{exp}\s*(years?|yrs?)\s*of\s*experience',
        rf'experience\s*[:\-]?\s*{exp}\s*(years?|yrs?)',
    ]
    return any(re.search(pat, text.lower()) for pat in patterns)

# Check Resume
def check_resume(filepath, criteria):
    text = extract_text(filepath).lower()
    if not text or 'error' in text:
        return False, ('', '', '')

    skills_match = any(skill.strip().lower() in set(text.split()) for skill in set(criteria['skills']))

    exp_match = experience_matches(text, criteria['experience'])
    edu_match = criteria['education'].lower() in text

    if skills_match and exp_match and edu_match:
        return True, extract_details(text)
    return False, ('', '', '')

# Shortlist Resumes
def shortlist_candidates():
    folder = folder_entry.get()
    skills = skills_entry.get().split(',')
    experience = experience_entry.get()
    education = education_entry.get()

    if not folder or not skills or not experience or not education:
        messagebox.showwarning("Missing Input", "Please fill all fields!")
        return

    try:
        experience = int(experience)
    except ValueError:
        messagebox.showwarning("Invalid Input", "Experience must be a number!")
        return

    criteria = {'skills': skills, 'experience': experience, 'education': education}
    shortlisted = []
    shortlisted_folder = os.path.join(folder, 'Shortlisted')
    os.makedirs(shortlisted_folder, exist_ok=True)

    output_text.delete(1.0, tk.END)

    for file in os.listdir(folder):
        filepath = os.path.join(folder, file)
        if os.path.isfile(filepath) and filepath.endswith(('.pdf','.docx')):
            match, details = check_resume(filepath, criteria)
            if match:
                shutil.copy(filepath, shortlisted_folder)
                name, email, phone = details
                shortlisted.append([name, email, phone, file])
                output_text.insert(tk.END, f"{name} | {email} | {phone} | {file}\n")

    if shortlisted:
        df = pd.DataFrame(shortlisted, columns=['Name', 'Email', 'Phone', 'Filename'])
        df.to_csv(os.path.join(shortlisted_folder, 'shortlisted.csv'), index=False)
        with open(os.path.join(shortlisted_folder, 'shortlisted.txt'), 'w') as f:
            for item in shortlisted:
                f.write(" | ".join(item) + "\n")

    status_label.config(text=f"Shortlisting Complete! {len(shortlisted)} Resumes Shortlisted")

#Preview Resume
def preview_resume(event):
    selected_line = output_text.get(output_text.index(tk.CURRENT) + " linestart", output_text.index(tk.CURRENT) + " lineend")
    parts = selected_line.split('|')
    if len(parts) >= 4:
        filename = parts[3].strip()
        folder = folder_entry.get()
        filepath = os.path.join(folder, filename)
        if os.path.exists(filepath):
            content = extract_text(filepath)
            preview_win = tk.Toplevel(root)
            preview_win.title(f"Preview - {filename}")
            text_area = scrolledtext.ScrolledText(preview_win, width=80, height=30)
            text_area.pack(padx=10, pady=10)
            text_area.insert(tk.END, content)
            text_area.config(state=tk.DISABLED)

# Browse Folder
def browse_folder():
    folder = filedialog.askdirectory()
    folder_entry.delete(0, tk.END)
    folder_entry.insert(0, folder)

# Clear All
def clear_all():
    folder_entry.delete(0, tk.END)
    skills_entry.delete(0, tk.END)
    experience_entry.delete(0, tk.END)
    education_entry.delete(0, tk.END)
    output_text.delete(1.0, tk.END)
    status_label.config(text="")

# Main GUI
root = tk.Tk()
root.title("Resume Shortlister")
root.geometry('700x600')
root.resizable(False, False)
root.eval('tk::PlaceWindow . center')

tk.Label(root, text="Folder Path").grid(row=0, column=0, padx=5, pady=5)
folder_entry = tk.Entry(root, width=60)
folder_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=browse_folder).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="Skills (comma-separated)").grid(row=1, column=0, padx=5, pady=5)
skills_entry = tk.Entry(root, width=60)
skills_entry.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Experience (years)").grid(row=2, column=0, padx=5, pady=5)
experience_entry = tk.Entry(root, width=60)
experience_entry.grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text="Education").grid(row=3, column=0, padx=5, pady=5)
education_entry = tk.Entry(root, width=60)
education_entry.grid(row=3, column=1, padx=5, pady=5)

tk.Button(root, text="Shortlist", command=shortlist_candidates, bg="lightgreen").grid(row=4, column=1, padx=5, pady=5)
tk.Button(root, text="Clear", command=clear_all, bg="lightblue").grid(row=4, column=2, padx=5, pady=5)

output_text = scrolledtext.ScrolledText(root, width=80, height=20)
output_text.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

output_text.bind("<Double-Button-1>", preview_resume)

status_label = tk.Label(root, text="", fg="green")
status_label.grid(row=6, column=0, columnspan=3)

root.mainloop()
