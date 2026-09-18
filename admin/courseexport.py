import os
import re
import csv
import sys
import time
import requests
import threading
import webbrowser
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from canvasapi import Canvas
from flask import Flask, render_template_string, request, Response, stream_with_context

app = Flask(__name__)

# --- ORIGINAL CORE LOGIC ---

def sanitize_filename(name):
    if not name:
        return "Untitled"
    return re.sub(r'[\\/*?:"<>|]', "", name)[:50]

def clean_html(html_content):
    if not html_content:
        return ""
    soup = BeautifulSoup(html_content, "html.parser")
    return soup.get_text(separator="\n").strip()

def html_to_docx(html_content, doc_path, title):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    if not html_content:
        doc.add_paragraph("[No content or body text provided.]")
        doc.save(doc_path)
        return

    soup = BeautifulSoup(html_content, "html.parser")
    for element in soup.descendants:
        if element.name in ["p", "div"]:
            text = element.get_text().strip()
            if text:
                p = doc.add_paragraph(text)
                p.paragraph_format.space_after = Pt(6)
        elif element.name in ["h2", "h3", "h4"]:
            text = element.get_text().strip()
            if text:
                level = int(element.name[1]) - 1
                h = doc.add_heading(text, level=level)
                h.paragraph_format.space_before = Pt(12)
                h.paragraph_format.space_after = Pt(4)
        elif element.name == "li":
            text = element.get_text().strip()
            if text:
                p = doc.add_paragraph(text, style="List Bullet")
                p.paragraph_format.space_after = Pt(3)

    if len(doc.paragraphs) == 2 and doc.paragraphs[0].text == title:
        doc.add_paragraph(soup.get_text())

    doc.save(doc_path)

class QuestionMock:
    def __init__(self, d):
        self.question_name = d.get("question_name", "Question")
        self.question_text = d.get("question_text", "")
        self.question_type = d.get("question_type", "")
        self.points_possible = d.get("points_possible", 0)
        self.answers = d.get("answers", [])
        self.matches = d.get("matches", [])

def generate_question_structure_docx(questions, doc_path, title, include_answers=False, is_bank=False):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    suffix = " (Answer Key)" if include_answers else ""
    prefix = "Question Bank: " if is_bank else ""
    main_title = f"{prefix}{title}{suffix}"
    
    heading = doc.add_heading(main_title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    if not include_answers:
        p_meta = doc.add_paragraph()
        p_meta.add_run("Name: _________________________\t\tDate: _____________\n")
        p_meta.paragraph_format.space_after = Pt(18)
    
    doc.add_heading("Questions", level=2)
    
    for i, q in enumerate(questions, start=1):
        q_name = getattr(q, "question_name", f"Question {i}")
        q_text = clean_html(getattr(q, "question_text", ""))
        q_type = getattr(q, "question_type", "")
        points = getattr(q, "points_possible", 0)
        
        if not q_text and hasattr(q, 'answers') and not getattr(q, 'answers'):
            continue
            
        qh = doc.add_heading(f"{i}. {q_name} ({points} pts)", level=3)
        qh.paragraph_format.space_before = Pt(14)
        qh.paragraph_format.space_after = Pt(6)
        
        qp = doc.add_paragraph(q_text)
        qp.paragraph_format.space_after = Pt(8)
        
        answers = getattr(q, "answers", [])
        if not answers:
            answers = []
        
        if q_type in ["multiple_choice_question", "multiple_answers_question", "true_false_question"]:
            for ans in answers:
                ans_text = clean_html(ans.get("text", ""))
                if not ans_text and ans.get("html"):
                    ans_text = clean_html(ans.get("html"))
                
                weight = ans.get("weight", 0)
                if weight is None:
                    weight = 0
                is_correct = float(weight) > 0
                
                if include_answers and is_correct:
                    p_ans = doc.add_paragraph()
                    run_box = p_ans.add_run("[X] ")
                    run_box.bold = True
                    run_text = p_ans.add_run(f"{ans_text}  <-- CORRECT ANSWER")
                    run_text.bold = True
                else:
                    p_ans = doc.add_paragraph("[  ] " + ans_text)
                p_ans.paragraph_format.left_indent = Inches(0.25)
                p_ans.paragraph_format.space_after = Pt(4)
                
        elif q_type == "essay_question":
            p_space = doc.add_paragraph()
            if include_answers:
                p_space.add_run("[ Essay Question: Student responses will vary. ]")
            else:
                p_space.add_run("\n" * 6)
            p_space.paragraph_format.space_after = Pt(12)
            
        elif q_type in ["short_answer_question", "numerical_question"]:
            p_blank = doc.add_paragraph()
            if include_answers:
                correct_answers = [clean_html(str(a.get("text", ""))) for a in answers if float(a.get("weight", 0) or 0) > 0 or "text" in a]
                ans_string = ", ".join(filter(None, correct_answers))
                p_blank.add_run(f"Correct Answer(s): {ans_string}").bold = True
            else:
                p_blank.add_run("Answer: _____________________________________")
            p_blank.paragraph_format.space_after = Pt(12)
            
        elif q_type == "matching_question":
            matches = getattr(q, "matches", [])
            for ans in answers:
                left_side = clean_html(ans.get("text", ""))
                right_side_correct = clean_html(ans.get("match_text", ""))
                
                p_match = doc.add_paragraph()
                if include_answers:
                    p_match.add_run(f"_____  {left_side}  ==> Correct Match: {right_side_correct}")
                else:
                    p_match.add_run(f"_____  {left_side}")
                p_match.paragraph_format.left_indent = Inches(0.25)
                p_match.paragraph_format.space_after = Pt(4)
                
            if not include_answers and matches:
                doc.add_paragraph("\nPossible Choices:").bold = True
                all_choices = sorted(list(set([clean_html(m.get("text", "")) for m in matches if m.get("text", "")])))
                for choice in all_choices:
                    p_choice = doc.add_paragraph(f"- {choice}")
                    p_choice.paragraph_format.left_indent = Inches(0.5)
                    p_choice.paragraph_format.space_after = Pt(2)
        else:
            if include_answers and answers:
                p_fallback = doc.add_paragraph("Answers: ")
                for ans in answers:
                    if float(ans.get("weight", 0) or 0) > 0:
                        p_fallback.add_run(f"[{clean_html(ans.get('text',''))}] ")
            else:
                p_fallback = doc.add_paragraph("_____________________________________")
            p_fallback.paragraph_format.space_after = Pt(12)

    doc.save(doc_path)

def generate_rubric_docx(rubric_data, doc_path):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    title = rubric_data.get("title", "Unnamed Rubric")
    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    criteria = rubric_data.get("data", [])
    if not criteria:
        doc.add_paragraph("[This rubric contains no criteria rows.]")
        doc.save(doc_path)
        return

    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Criterion Description'
    hdr_cells[1].text = 'Ratings / Performance Levels'
    hdr_cells[2].text = 'Pts'
    
    for item in criteria:
        row_cells = table.add_row().cells
        desc = item.get("description", "No description")
        long_desc = item.get("long_description", "")
        cell_0_text = f"{desc}"
        if long_desc:
            cell_0_text += f"\n\nDetails: {long_desc}"
        row_cells[0].text = cell_0_text
        
        ratings = item.get("ratings", [])
        ratings_text_list = []
        for r in ratings:
            r_desc = r.get("description", "No rating name")
            r_long = r.get("long_description", "")
            r_pts = r.get("points", 0)
            r_str = f"• {r_pts} pts - {r_desc}"
            if r_long:
                r_str += f" ({r_long})"
            ratings_text_list.append(r_str)
            
        row_cells[1].text = "\n".join(ratings_text_list)
        row_cells[2].text = f"{item.get('points', 0)} pts"

    doc.save(doc_path)

def generate_rubric_csv(rubric_data, csv_path):
    criteria = rubric_data.get("data", [])
    if not criteria:
        return

    max_ratings = max(len(item.get("ratings", [])) for item in criteria) if criteria else 0
    header = ["Rubric Name", "Criteria Name", "Criteria Description", "Criteria Enable Range"]
    for _ in range(max_ratings):
        header.extend(["Rating Name", "Rating Description", "Rating Points"])

    with open(csv_path, mode='w', newline='', encoding='utf-8') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(header)
        
        rubric_title = rubric_data.get("title", "Unnamed Rubric")
        for item in criteria:
            row = [
                rubric_title,
                item.get("description", ""),
                item.get("long_description", ""),
                str(item.get("criterion_use_range", False)).lower()
            ]
            
            ratings = item.get("ratings", [])
            for i in range(max_ratings):
                if i < len(ratings):
                    r = ratings[i]
                    row.extend([
                        r.get("description", ""),
                        r.get("long_description", ""),
                        r.get("points", 0)
                    ])
                else:
                    row.extend(["", "", ""])
            writer.writerow(row)

# --- WEB APPLICATION BACKEND & STREAMING INTERFACE ---

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Canvas Course Backup Tool</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background-color: #f4f6f8; color: #333; margin: 0; padding: 40px; }
        .container { max-width: 650px; background: white; margin: 0 auto; padding: 30px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
        h2 { color: #1E3A8A; margin-top: 0; }
        p { color: #666; font-size: 14px; line-height: 1.5; }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-weight: bold; margin-bottom: 5px; font-size: 14px; color: #222; }
        .explanation { font-size: 12px; color: #666; margin-top: 4px; line-height: 1.4; font-weight: normal; }
        input[type="text"], input[type="password"] { width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; font-size: 15px; }
        .inst-wrapper { display: flex; align-items: center; }
        .inst-wrapper span { background: #eee; padding: 10px; border: 1px solid #ccc; border-radius: 4px 0 0 4px; border-right: none; font-size: 14px; color: #555; }
        .inst-wrapper input { border-radius: 0 4px 4px 0; }
        
        /* Interactive Buttons Style Layout */
        button { background-color: #1E3A8A; color: white; border: none; padding: 14px 20px; font-size: 16px; border-radius: 4px; cursor: pointer; width: 100%; font-weight: bold; transition: background-color 0.2s; }
        button:hover { background-color: #172554; }
        button:disabled { background-color: #64748B; cursor: not-allowed; }
        
        .action-button-container { display: flex; gap: 12px; margin-top: 10px; }
        #resetBtn { display: none; background-color: #059669; flex: 1; }
        #resetBtn:hover { background-color: #047857; }
        #exitBtn { display: none; background-color: #DC2626; flex: 1; }
        #exitBtn:hover { background-color: #B91C1C; }
        
        #log-window { background: #1e1e1e; color: #00ff00; padding: 15px; border-radius: 4px; height: 250px; overflow-y: auto; font-family: monospace; font-size: 13px; display: none; margin-top: 20px; white-space: pre-wrap; line-height: 1.4; }
        
        #shutdown-notice { display: none; background-color: #FEF3C7; border: 1px solid #F59E0B; color: #92400E; padding: 15px; border-radius: 6px; font-weight: bold; font-size: 14px; margin-top: 20px; text-align: center; line-height: 1.5; }
    </style>
</head>
<body>
    <div class="container" id="mainContainer">
        <h2>Canvas Course Content Exporter</h2>
        <p>This utility downloads course content onto your desktop computer, formatted into Word (.docx) documents and Excel (.csv) spreadsheets.</p>
        <hr style="border: 0; border-top: 1px solid #eee; margin-bottom: 25px;">
        
        <form id="backupForm" action="/run-backup" method="POST" target="logFrame">
            <div class="form-group">
                <label for="institution">Your School's Canvas Instance</label>
                <div class="inst-wrapper">
                    <span>https://</span>
                    <input type="text" id="institution" name="institution" placeholder="schoolname" required>
                    <span>.instructure.com</span>
                </div>
            </div>
            
            <div class="form-group">
                <label for="api_key">Access Token provided by your Canvas Admin</label>
                <input type="password" id="api_key" name="api_key" placeholder="Paste full alphanumeric token here..." required>
            </div>
            
            <div class="form-group">
                <label for="course_id">Canvas Course ID</label>
                <input type="text" id="course_id" name="course_id" placeholder="e.g., 1234567" required>
                <div class="explanation">
                    This Course ID can be found in your course homepage URL. E.g., <strong>1234567</strong> is the Course ID with a homepage URL of <code>https://school.instructure.com/courses/1234567</code>
                </div>
            </div>
            
            <button type="submit" id="submitBtn">Download Course Content to your Computer</button>
            
            <div class="action-button-container">
                <button type="button" id="resetBtn">Export Another Course</button>
                <button type="button" id="exitBtn">Exit Exporter</button>
            </div>
        </form>

        <div id="log-window"></div>
        <div id="shutdown-notice">
            The Exporter has been shut down successfully.<br>
            <strong>It is now safe to close this web browser tab.</strong>
        </div>
        
        <iframe name="logFrame" style="display:none;"></iframe>
    </div>

    <script>
        const form = document.getElementById('backupForm');
        const logWindow = document.getElementById('log-window');
        const submitBtn = document.getElementById('submitBtn');
        const resetBtn = document.getElementById('resetBtn');
        const exitBtn = document.getElementById('exitBtn');
        const courseIdInput = document.getElementById('course_id');
        const shutdownNotice = document.getElementById('shutdown-notice');

        form.onsubmit = function() {
            logWindow.style.display = 'block';
            logWindow.innerHTML = "Initializing connection to Canvas...<br>";
            submitBtn.disabled = true;
            submitBtn.innerText = "Downloading Content...";
            
            window.onMessageReceived = function(messageData) {
                if (messageData === "__FINISHED__") {
                    submitBtn.style.display = 'none'; 
                    resetBtn.style.display = 'block'; 
                    exitBtn.style.display = 'block'; // Reveal native app close pipeline trigger
                    return;
                }
                
                let cleanedData = messageData.replace(/\\\\n/g, '<br>');
                logWindow.innerHTML += cleanedData + "<br>";
                logWindow.scrollTop = logWindow.scrollHeight;
            };
        };

        // Reset click pipeline
        resetBtn.onclick = function() {
            logWindow.innerHTML = "";
            logWindow.style.display = 'none';
            
            courseIdInput.value = ""; 
            resetBtn.style.display = 'none'; 
            exitBtn.style.display = 'none';
            
            submitBtn.style.display = 'block'; 
            submitBtn.disabled = false;
            submitBtn.innerText = "Download Course Content to your Computer";
            
            courseIdInput.focus(); 
        };

        // Server Kill Script Sequence Execution
        exitBtn.onclick = function() {
            // Hide control forms instantly to provide feedback that the app is stopping
            form.style.display = 'none';
            logWindow.style.display = 'none';
            shutdownNotice.style.display = 'block';
            
            // Post death-signal to internal flask routing pipeline natively
            fetch('/shutdown', { method: 'POST' }).catch(err => {
                // Ignore fallback network disruption drop as server dies mid-request handler lifecycle
            });
        };

        window.addEventListener('message', function(e) {
            if (window.onMessageReceived) {
                window.onMessageReceived(e.data);
            }
        });
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/run-backup', methods=['POST'])
def run_backup():
    institution = request.form.get('institution').strip()
    api_key = request.form.get('api_key').strip()
    course_id_str = request.form.get('course_id').strip()
    
    api_url = f"https://{institution}.instructure.com"
    
    try:
        course_id = int(course_id_str)
    except ValueError:
        return Response("<script>window.parent.postMessage('Error: Course ID must be numeric.', '*');</script>", mimetype='text/html')

    def execute_backup_stream():
        def log(msg):
            return f"<script>window.parent.postMessage({repr(msg)}, '*');</script>\n"

        yield log(f"Connecting to Canvas at {api_url}...")
        canvas = Canvas(api_url, api_key)
        try:
            course = canvas.get_course(course_id, include=["syllabus_body"])
            yield log(f"Successfully Connected to Course: {course.name}")
        except Exception as e:
            yield log(f"Connection Failure: {str(e)}")
            yield log("__FINISHED__")
            return

        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        root_dir = os.path.join(desktop_path, sanitize_filename(f"Canvas_Backup_{course.name}_{course_id}"))
        os.makedirs(root_dir, exist_ok=True)

        folders = ["Assignments", "Quizzes", "Question_Banks", "Rubrics", "Discussions", "Announcements", "Pages", "Syllabus", "Gradebook", "Course_Files"]
        for folder in folders:
            os.makedirs(os.path.join(root_dir, folder), exist_ok=True)

        # 1. Assignments
        yield log("Exporting Assignments...")
        assignments = list(course.get_assignments())
        for assign in assignments:
            filename = f"{sanitize_filename(assign.name)}.docx"
            filepath = os.path.join(root_dir, "Assignments", filename)
            html_to_docx(getattr(assign, "description", ""), filepath, assign.name)

        # 2. Quizzes
        yield log("Exporting Quizzes...")
        for quiz in course.get_quizzes():
            safe_title = sanitize_filename(quiz.title)
            try:
                questions = list(quiz.get_questions())
            except Exception:
                continue
            generate_question_structure_docx(questions, os.path.join(root_dir, "Quizzes", f"{safe_title}_Printable_Quiz.docx"), quiz.title, include_answers=False)
            generate_question_structure_docx(questions, os.path.join(root_dir, "Quizzes", f"{safe_title}_Answer_Key.docx"), quiz.title, include_answers=True)

        # 3. Question Banks
        yield log("Exporting Question Banks...")
        try:
            headers = {"Authorization": f"Bearer {api_key}"}
            base_api = api_url.rstrip('/')
            banks_url = f"{base_api}/api/v1/question_banks?context_type=Course&context_id={course_id}"
            response = requests.get(banks_url, headers=headers)
            if response.status_code == 200:
                for bank in response.json():
                    bank_id = bank.get("id")
                    bank_title = bank.get("title", "Unnamed Bank")
                    safe_bank_title = sanitize_filename(bank_title)
                    questions_url = f"{base_api}/api/v1/question_banks/{bank_id}/questions?per_page=100"
                    q_response = requests.get(questions_url, headers=headers)
                    if q_response.status_code == 200:
                        bank_questions = [QuestionMock(q) for q in q_response.json()]
                        if not bank_questions: continue
                        generate_question_structure_docx(bank_questions, os.path.join(root_dir, "Question_Banks", f"{safe_bank_title}_Printable_Bank.docx"), bank_title, include_answers=False, is_bank=True)
                        generate_question_structure_docx(bank_questions, os.path.join(root_dir, "Question_Banks", f"{safe_bank_title}_Answer_Key.docx"), bank_title, include_answers=True, is_bank=True)
        except Exception as e:
            yield log(f"  Warning (Question Banks): {str(e)}")

        # 4. Rubrics
        yield log("Exporting Rubrics...")
        try:
            rubrics = course.get_rubrics(include=["assessments", "associations"])
            for r in rubrics:
                detailed_rubric_url = f"{api_url.rstrip('/')}/api/v1/courses/{course_id}/rubrics/{r.id}"
                headers = {"Authorization": f"Bearer {api_key}"}
                r_response = requests.get(detailed_rubric_url, headers=headers)
                if r_response.status_code == 200:
                    rubric_data = r_response.json()
                    safe_rubric_title = sanitize_filename(rubric_data.get("title", f"Rubric_{r.id}"))
                    generate_rubric_docx(rubric_data, os.path.join(root_dir, "Rubrics", f"{safe_rubric_title}.docx"))
                    generate_rubric_csv(rubric_data, os.path.join(root_dir, "Rubrics", f"{safe_rubric_title}_Importable.csv"))
        except Exception as e:
            yield log(f"  Warning (Rubrics): {str(e)}")

        # 5. Discussions
        yield log("Exporting Discussions...")
        for topic in course.get_discussion_topics(only_announcements=False):
            if not getattr(topic, "is_announcement", False):
                filename = f"{sanitize_filename(topic.title)}.docx"
                filepath = os.path.join(root_dir, "Discussions", filename)
                html_to_docx(getattr(topic, "message", ""), filepath, topic.title)

        # 6. Announcements
        yield log("Exporting Announcements...")
        for announce in course.get_discussion_topics(only_announcements=True):
            filename = f"{sanitize_filename(announce.title)}.docx"
            filepath = os.path.join(root_dir, "Announcements", filename)
            html_to_docx(getattr(announce, "message", ""), filepath, announce.title)

        # 7. Content Pages
        yield log("Exporting Pages...")
        for page in course.get_pages():
            detailed_page = course.get_page(page.url)
            filename = f"{sanitize_filename(detailed_page.title)}.docx"
            filepath = os.path.join(root_dir, "Pages", filename)
            html_to_docx(getattr(detailed_page, "body", ""), filepath, detailed_page.title)

        # 8. Syllabus
        yield log("Exporting Syllabus...")
        syllabus_html = getattr(course, "syllabus_body", "")
        html_to_docx(syllabus_html, os.path.join(root_dir, "Syllabus", "Syllabus.docx"), f"{course.name} Syllabus")

        # 9. Gradebook
        yield log("Exporting Gradebook...")
        try:
            students = list(course.get_users(enrollment_type=['student']))
            csv_headers = ["Student Name", "Canvas User ID"]
            assignment_id_map = {}
            for assign in assignments:
                header_name = f"{assign.name} (Max: {getattr(assign, 'points_possible', 0)})"
                csv_headers.append(header_name)
                assignment_id_map[assign.id] = header_name
            csv_filepath = os.path.join(root_dir, "Gradebook", f"Gradebook_Backup_{course_id}.csv")
            with open(csv_filepath, mode='w', newline='', encoding='utf-8') as csv_file:
                writer = csv.DictWriter(csv_file, fieldnames=csv_headers)
                writer.writeheader()
                for student in students:
                    row = {"Student Name": student.name, "Canvas User ID": student.id}
                    submissions = course.get_multiple_submissions(student_ids=[student.id], include=['assignment'])
                    for sub in submissions:
                        assign_id = getattr(sub, "assignment_id", None)
                        if assign_id in assignment_id_map:
                            col_name = assignment_id_map[assign_id]
                            score = getattr(sub, "score", "")
                            row[col_name] = score if score is not None else "Missing"
                    writer.writerow(row)
        except Exception as e:
            yield log(f"  Warning (Gradebook): {str(e)}")

        # 10. Course Files
        yield log("Downloading Course Files (This may take a while)...")
        try:
            canvas_folders = list(course.get_folders())
            folder_id_to_path = {}
            for folder in canvas_folders:
                clean_path_parts = [sanitize_filename(p) for p in folder.full_name.split('/') if p]
                if clean_path_parts and clean_path_parts[0].lower() == "course files":
                    clean_path_parts.pop(0)
                local_target_dir = os.path.join(root_dir, "Course_Files", *clean_path_parts)
                os.makedirs(local_target_dir, exist_ok=True)
                folder_id_to_path[folder.id] = local_target_dir

            canvas_files = course.get_files()
            headers_file = {"Authorization": f"Bearer {api_key}"}
            for file in canvas_files:
                parent_folder_id = getattr(file, "folder_id", None)
                target_directory = folder_id_to_path.get(parent_folder_id, os.path.join(root_dir, "Course_Files"))
                is_unpublished = getattr(file, "locked", False) or getattr(file, "hidden", False)
                prefix_str = "[UNPUBLISHED] " if is_unpublished else ""
                filename = re.sub(r'[\\/*?:"<>|]', "", prefix_str + file.display_name)
                file_extract_path = os.path.join(target_directory, filename)
                try:
                    with requests.get(file.url, headers=headers_file, stream=True) as r:
                        r.raise_for_status()
                        with open(file_extract_path, 'wb') as f:
                            for chunk in r.iter_content(chunk_size=8192):
                                f.write(chunk)
                except Exception:
                    continue
        except Exception as e:
            yield log(f"  Warning (Files): {str(e)}")

        yield log(f"SUCCESS! Complete course archive saved to your Desktop at: {root_dir}")
        yield log("__FINISHED__")

    return Response(stream_with_context(execute_backup_stream()), mimetype='text/html')

@app.route('/shutdown', methods=['POST'])
def shutdown():
    # Programmatically sends a termination signal to the running engine process framework cleanly
    os._exit(0)

def open_browser():
    time.sleep(1)
    webbrowser.open("http://127.0.0.1:5000/")

if __name__ == '__main__':
    # 1. MANDATORY FOR WINDOWS: Prevents PyInstaller from spawning infinite background processes
    if sys.platform == "win32":
        import multiprocessing
        multiprocessing.freeze_support()

    # 2. Start the background thread to automatically launch the user's web browser
    threading.Thread(target=open_browser, daemon=True).start()
    
    # 3. Launch the web app frame local interface
    app.run(host="127.0.0.1", port=5000, debug=False)