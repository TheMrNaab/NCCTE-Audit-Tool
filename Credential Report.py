# Copyright © 2025 David Naab
# This software is provided free of charge for individual, educational use only.
# Attribution is required in all derivative works and distributed outputs.
# Organization-level, institutional, or district-wide use — including deployment, modification, or redistribution within a school system, company, or agency — is strictly prohibited without prior written permission from the author.
# To request licensing or usage approval for school district or institutional use, contact the author directly.
# URL: https://github.com/TheMrNaab/NCCTE-Audit-Tool

import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tabulate import tabulate
import os
from datetime import datetime

# --- License Agreement Prompt ---
print("Copyright © 2025 David Naab\n")
print("This software is provided free of charge for individual, educational use only.\n")
print("Attribution is required in all derivative works and distributed outputs.\n")
print("Organization-level, institutional, or district-wide use — including deployment,")
print("modification, or redistribution within a school system, company, or agency —")
print("is strictly prohibited without prior written permission from the author.\n")
print("To request licensing or usage approval for school district or institutional use,")
print("contact the author directly.\n\n")

agreement = input("Do you agree to these terms? (yes/no): ").strip().lower()
if agreement != "yes":
    print("\n\nYou must accept the license agreement to use this software. Exiting.")
    exit(1)

print("\n\n\n\n")

# --- Formatting helpers ---
def bold(text):
    return f"\033[1m{text}\033[0m"

def file_info(path):
    name = os.path.basename(path)
    mtime = os.path.getmtime(path)
    date = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
    return f"{name} (last modified: {date})"

def color_result(result):
    if isinstance(result, str) and result.lower() == 'pass':
        return f"\033[92m{result}\033[0m"
    elif isinstance(result, str) and result.lower() == 'fail':
        return f"\033[91m{result}\033[0m"
    return result

def html_color_result(result):
    if isinstance(result, str) and result.lower() == 'pass':
        return '<span style="color:green;font-weight:bold">Pass</span>'
    elif isinstance(result, str) and result.lower() == 'fail':
        return '<span style="color:red;font-weight:bold">Fail</span>'
    return result

# --- File dialogs ---
def select_file(prompt_title):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(title=prompt_title, filetypes=[("CSV files", "*.csv")])

def get_save_path():
    root = tk.Tk()
    root.withdraw()
    return filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML files", "*.html")])

# --- File selection ---
results_path = select_file("Select 'Results (264) Plus.csv'")
credentials_path = select_file("Select 'Credentials_ByStudentCredential' CSV")

# --- Load data ---
results_df = pd.read_csv(results_path, skiprows=3)
credentials_df = pd.read_csv(credentials_path, skiprows=5)

# --- Exam selection ---
unique_exams = sorted(results_df['Exam'].dropna().unique())
print("\nAvailable Exams:")
for i, exam in enumerate(unique_exams, start=1):
    print(f"{i}. {exam}")

selected_indices = input("\nEnter the numbers of the exams to analyze (comma-separated): ")
selected_exam_indices = [int(i.strip()) - 1 for i in selected_indices.split(",") if i.strip().isdigit()]
selected_exams = [unique_exams[i] for i in selected_exam_indices]

print("\nSelected Exams:")
for ex in selected_exams:
    print(f" - {ex}")

# --- Merge and process data ---
merged_df = credentials_df.merge(
    results_df,
    how='left',
    left_on='Student ID',
    right_on='StudentEmployeeID'
)

def determine_status(row):
    if row['Met'] == 1: return 'Met'
    if row['Not Met'] == 1: return 'Not Met'
    if row['Not Attempted'] == 1: return 'Not Attempted'
    if row['Not Offered'] == 1: return 'Not Offered'
    if row['Previously Reported'] == 1: return 'Previously Reported'
    if row['Previous Enrollment'] == 1: return 'Previous Enrollment'
    return 'Unknown'

merged_df['Status'] = merged_df.apply(determine_status, axis=1)

student_info = merged_df[['Student ID', 'Student First Name', 'Student Last Name', 'Status']].drop_duplicates()
exam_info = merged_df[['Student ID', 'Exam', 'ExamDate', 'Score', 'Result']].dropna(subset=['Exam'])

exam_info_grouped = exam_info.groupby('Student ID').apply(
    lambda df: df[['Exam', 'ExamDate', 'Score', 'Result']].to_dict('records')
).to_dict()

# Certiport status for selected exams
certiport_results = results_df[results_df['Exam'].isin(selected_exams)][['StudentEmployeeID', 'Result']]
certiport_status = (
    certiport_results.groupby('StudentEmployeeID')['Result']
    .apply(lambda results: 'Pass' if 'Pass' in results.values else 'Fail')
    .to_dict()
)

def get_certiport_status(sid):
    return certiport_status.get(sid, "")

def get_flag(status, certiport):
    if status == "Not Attempted":
        if certiport in ["Pass", "Fail"]:
            return "⚠️ Student has exam results but is marked as 'Not Attempted'"
        else:
            return "⚠️ Student is marked as 'Not Attempted' and has no selected exam attempts"
    elif certiport == "":
        return "⚠️ No selected exam attempts found for this student"
    return ""

student_info['Exams'] = student_info['Student ID'].map(lambda sid: exam_info_grouped.get(sid, []))
student_info['Has Exams'] = student_info['Exams'].apply(lambda x: isinstance(x, list) and len(x) > 0)
student_info['Certiport Status'] = student_info['Student ID'].map(get_certiport_status)
student_info['Flag'] = student_info.apply(lambda row: get_flag(row['Status'], row['Certiport Status']), axis=1)

# Sort students
students_with_exams = student_info[student_info['Has Exams']].sort_values(by='Student Last Name')
students_without_exams = student_info[~student_info['Has Exams']].sort_values(by='Student Last Name')
sorted_students = pd.concat([students_with_exams, students_without_exams])

# --- HTML head with CSS ---
html_head = """
<head>
    <meta charset="UTF-8">
    <title>Credential Report</title>
    
</head>
<style>
    body {
        font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
        background-color: #fff;
        color: #000;
        padding: 10px 20px;
        line-height: 1.4;
    }

    .comment-toggle button {
    font-size: 12px;
    padding: 4px 8px;
    margin-top: 6px;
    cursor: pointer;
    background-color: #eee;
    border: 1px solid #bbb;
    border-radius: 4px;
    }

    h1 {
        color: #2c3e50;
        font-size: 22px;
        margin: 0 0 5px;
    }

    h2 {
        font-size: 16px;
        color: #555;
        margin: 10px 0;
    }

    p, li {
        font-size: 13px;
        margin: 4px 0;
    }

    ul {
        margin: 4px 0 12px 20px;
        padding-left: 0;
    }

    hr {
        border: none;
        border-top: 1px solid #ccc;
        margin: 20px 0 10px;
    }

    table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 5px;
    }

    th {
        background-color: #2c3e50;
        color: white;
        padding: 6px;
        font-size: 13px;
        text-align: left;
    }

    td {
        padding: 6px;
        border-bottom: 1px solid #ddd;
        font-size: 13px;
    }

    .flag {
        color: #c0392b;
        font-weight: bold;
        font-size: 13px;
        margin-bottom: 8px;
    }

    .section {
        background-color: #fff;
        padding: 10px 15px;
        margin-bottom: 10px;
        border: 1px solid #ccc;
        border-radius: 4px;
    }

    .comments-box {
        width: 100%;
        height: 60px;
        padding: 6px;
        border: 1px solid #ccc;
        border-radius: 4px;
        font-family: inherit;
        font-size: 13px;
        resize: vertical;
        background-color: #fff;
    }

    footer, footer p, footer ul, footer li {
        font-size: 10px;
        color: #666;
        text-align: left;
    }

    footer {
        padding-top: 10px;
        margin-top: 30px;
        border-top: 1px solid #ccc;
    }

    @media print {
        body {
            background: white;
        }
        .section {
            box-shadow: none;
            page-break-inside: avoid;
        }
    }
</style>
"""

# --- Build HTML blocks ---
html_blocks = []
for _, student in sorted_students.iterrows():
    exams = student['Exams']
    rows = ""
    for e in exams:
        rows += f"<tr><td>{e['Exam']}</td><td>{e['ExamDate']}</td><td>{e['Score']}</td><td>{html_color_result(e['Result'])}</td></tr>"

    table_html = f"""
    <table>
        <tr><th>Exam</th><th>Exam Date</th><th>Score</th><th>Result</th></tr>
        {rows if rows else '<tr><td colspan="4">No exam records found.</td></tr>'}
    </table>
    """

    flag_html = f"<p class='flag'>{student['Flag']}</p>" if student['Flag'] else ""

    html_block = f"""
    <div class="section">
        {flag_html}
        <p><strong>Name:</strong> {student['Student First Name']} {student['Student Last Name']}</p>
        <p><strong>Student ID:</strong> {student['Student ID']}</p>
        <p><strong>Credential Status:</strong> {student['Status']}</p>
        <p><strong>Certiport Status:</strong> {student['Certiport Status']}</p>
        <p><strong>Exams:</strong></p>
        {table_html}
        <div class="comment-toggle">
            <button type="button" onclick="toggleCommentBox(this)">Add Comments</button>
            <div class="comments-box-container" style="display: none;">
                <textarea class="comments-box" placeholder="Enter comments here..."></textarea>
            </div>
        </div> 
    </div>
    """
    html_blocks.append(html_block)

footer_html = f"""
<footer>
    <p><strong>Report generated using:</strong></p>
    <ul>
        <li>Certiport Results File: {file_info(results_path)}</li>
        <li>Credentials File (NCCTE Admin): {file_info(credentials_path)}</li>
    </ul>
</footer>
"""

script = """
<script>
    function toggleCommentBox(button) {
        const container = button.nextElementSibling;
        const isVisible = container.style.display === "block";
        container.style.display = isVisible ? "none" : "block";
        button.textContent = isVisible ? "Add Comments" : "Hide Comments";
    }
</script>
"""

# --- Final HTML document ---
html_template = f"""
<html>
{html_head}
<body>
    <h1>Credential & Certiport Report</h1>
    <h2>Selected Exams:</h2>
    <ul>
        {''.join(f"<li>{ex}</li>" for ex in selected_exams)}
    </ul>
    <p>This report compares Certiport exam results with NC CTE Admin credential records. It matches students by ID, lists all exam attempts, and evaluates credential status. Flags are shown for students who are marked as ‘Not Attempted’ but have results, as well as students who have no record of taking the selected exams.</p>
    {''.join(html_blocks)}
    {footer_html}
    
    
</body>
{script}
</html>
"""

# --- Save HTML ---
html_path = get_save_path()
if html_path:
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_template)
    print(f"\n✅ HTML report saved to: {html_path}")
else:
    print("\n❌ HTML export cancelled.")

# --- Console output ---
for _, student in sorted_students.iterrows():
    print("==========================================")
    if student['Flag']:
        print(f"{bold('⚠️ Flag')} : {student['Flag']}")
    print(f"{bold('Student ID')}       : {student['Student ID']}")
    print(f"{bold('Name')}             : {student['Student First Name']} {student['Student Last Name']}")
    print(f"{bold('Credential Status')}: {student['Status']}")
    print(f"{bold('Certiport Status')} : {color_result(student['Certiport Status'])}")
    print(f"{bold('Exams')}:")

    exams = student['Exams']
    if not exams:
        print("  No exam records found.")
    else:
        table_data = [
            [e['Exam'], e['ExamDate'], e['Score'], color_result(e['Result'])]
            for e in exams
        ]
        print(tabulate(table_data, headers=["Exam", "Exam Date", "Score", "Result"], tablefmt="grid"))
    print("\n\n")
