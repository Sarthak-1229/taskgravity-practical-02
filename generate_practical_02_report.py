import os
import subprocess
import html
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_bg(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def generate_docx():
    doc = docx.Document()
    
    # Page Margins
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    # Base Style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)

    # Department Header
    p_dept = doc.add_paragraph()
    p_dept.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_dept = p_dept.add_run("Department of Computer Engineering (Software Engineering)\n")
    r_dept.font.size = Pt(14)
    r_dept.font.bold = True
    r_dept.font.color.rgb = RGBColor(15, 23, 42)

    r_sub = p_dept.add_run("Web Technology Laboratory — Academic Year 2026-27\n")
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("Practical No. 02: Interactive Client-Side Web Application")
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(37, 99, 235)

    # Metadata Table
    meta_table = doc.add_table(rows=2, cols=4)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        [("Name:", "Sarthak Bhosale"), ("PRN:", "202401100051")],
        [("Batch:", "TY SE-A2"), ("Roll No:", "A50")]
    ]
    for row_idx, row in enumerate(meta_data):
        col_offset = 0
        for label, val in row:
            cell_lbl = meta_table.cell(row_idx, col_offset)
            cell_val = meta_table.cell(row_idx, col_offset + 1)
            cell_lbl.text = label
            cell_lbl.paragraphs[0].runs[0].font.bold = True
            cell_lbl.paragraphs[0].runs[0].font.color.rgb = RGBColor(71, 85, 105)
            cell_val.text = val
            cell_val.paragraphs[0].runs[0].font.bold = True
            cell_val.paragraphs[0].runs[0].font.color.rgb = RGBColor(15, 23, 42)
            set_cell_bg(cell_lbl, "F1F5F9")
            set_cell_bg(cell_val, "FFFFFF")
            set_cell_margins(cell_lbl, 80, 80, 100, 100)
            set_cell_margins(cell_val, 80, 80, 100, 100)
            col_offset += 2

    doc.add_paragraph()

    def add_section_heading(num_str, title_str):
        h = doc.add_paragraph()
        r = h.add_run(f"{num_str}. {title_str}")
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = RGBColor(15, 23, 42)
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(4)
        return h

    # 1. Aim
    add_section_heading("1", "Aim")
    doc.add_paragraph(
        "To develop an interactive client-side web application for a Task Management System named 'TaskGravity' "
        "using HTML5, CSS3, and JavaScript, demonstrating the usage of JSON and Arrays for data storage, "
        "comprehensive client-side form validation with visual error feedback, dynamic DOM manipulation, "
        "and event-driven state persistence via localStorage."
    )

    # 2. Objectives
    add_section_heading("2", "Objectives")
    objs = [
        "To store, manipulate, and query structured application data using JSON objects and JavaScript arrays.",
        "To implement strict client-side validation rules (string length bounds, date validity, effort boundaries, category/priority constraints).",
        "To render real-time and submit-time error messages dynamically without relying on generic browser popups.",
        "To dynamically update the Document Object Model (DOM) in response to user actions (Add, Cycle Status, Delete, Search, Filter, Sort).",
        "To compute and display live aggregate metrics (Total Backlog, In Progress, Pending Review, Completion Velocity %).",
        "To achieve local data persistence across page reloads using the browser's localStorage API.",
        "To deploy the application to GitHub Pages and document repository and live deployment links."
    ]
    for o in objs:
        p = doc.add_paragraph(o, style='List Bullet')
        p.paragraph_format.space_after = Pt(2)

    # 3. Theory & Concepts
    add_section_heading("3", "Theory & Concepts")
    doc.add_paragraph(
        "Modern frontend architectures depend heavily on client-side state management, responsive event handling, "
        "and rigorous input sanitization. In Practical 02, TaskGravity implements an interactive, event-driven architecture:"
    )

    concepts = [
        ("3.1 JSON and Array Data Architecture",
         "JavaScript Object Notation (JSON) is a lightweight text-based data interchange format. In TaskGravity, "
         "tasks are modeled as structured JSON records containing unique IDs, titles, categories, priorities, deadlines, "
         "estimated effort hours, and lifecycle statuses. These objects reside within an in-memory JavaScript array that "
         "serves as the single source of truth for the client. Changes are serialized to JSON strings using JSON.stringify() "
         "and deserialized upon reload using JSON.parse() via the Web Storage API (localStorage)."),
        ("3.2 Defensive Input Validation & Error Messaging",
         "Client-side validation enhances user experience by catching malformed or missing data before storage or transmission. "
         "TaskGravity evaluates: (1) Title presence and character bounds [3-80 chars]; (2) Architectural category whitelist; "
         "(3) Priority whitelist; (4) Deadline validity and past-date rejection (comparing against today's midnight timestamp); "
         "(5) Numeric effort hours bounded between 0.5 and 100. Field-specific .error-feedback elements render immediately upon "
         "keystroke (input event) or blur (change event), preventing user frustration."),
        ("3.3 Dynamic DOM Manipulation & Event Handling",
         "The Document Object Model (DOM) represents HTML as a tree of objects. TaskGravity binds event listeners "
         "(addEventListener) to form submissions, keystrokes, filter selects, and status toggle buttons. Rather than whole-page "
         "refreshes, the DOM updates efficiently using template literals, preserving focus and delivering a fluid experience."),
        ("3.4 Querying, Filtering, and Sorting Algorithms",
         "Tasks are filtered dynamically using JavaScript array primitives (filter, sort). The application enables compound "
         "filtering (e.g. status AND category AND search substring) and multi-attribute sorting (by deadline, priority weight, "
         "effort hours, or alphabetical title), with automatic recalculation of KPI progress percentages.")
    ]
    for title, desc in concepts:
        h = doc.add_paragraph()
        r = h.add_run(title)
        r.font.bold = True
        r.font.size = Pt(11.5)
        h.paragraph_format.space_before = Pt(6)
        h.paragraph_format.space_after = Pt(2)
        doc.add_paragraph(desc)

    # 4. Input Validation Matrix
    add_section_heading("4", "Input Validation Rules Matrix")
    doc.add_paragraph("The validation rules enforced prior to JSON array storage are detailed below:")

    v_table = doc.add_table(rows=7, cols=4)
    v_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Field Name", "Data Type", "Validation Constraints", "Error Feedback Triggered"]
    for i, h in enumerate(headers):
        c = v_table.cell(0, i)
        c.text = h
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        set_cell_bg(c, "1E293B")
        set_cell_margins(c, 80, 80, 100, 100)

    val_data = [
        ("Task Title", "String", "Required, trimmed length >= 3 and <= 80 chars", "Task Title must be between 3 and 80 characters."),
        ("Category", "Enum String", "Must match whitelist: Frontend, Backend, Database, DevOps, Testing, UI/UX", "Please select a valid architectural category."),
        ("Priority", "Enum String", "Must match whitelist: Low, Medium, High, Urgent", "Please select an operational priority level."),
        ("Deadline", "ISO Date", "Required, valid format, strictly >= current date", "Target deadline is required / Cannot be in the past."),
        ("Effort Hours", "Number (Float)", "Required, numeric, strictly > 0 and <= 100 hours", "Estimated effort must be between 1 and 100 hours."),
        ("Description", "String", "Optional, maximum length <= 300 chars", "Description cannot exceed 300 characters.")
    ]
    for r_idx, row in enumerate(val_data, start=1):
        for c_idx, val in enumerate(row):
            c = v_table.cell(r_idx, c_idx)
            c.text = val
            set_cell_bg(c, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(c, 60, 60, 80, 80)

    doc.add_paragraph()

    # 5. Algorithm & Implementation Steps
    add_section_heading("5", "Algorithm & Implementation Steps")
    steps = [
        "Construct responsive dual-column interface (index.html) with form panel and dynamic stream.",
        "Author custom flat stylesheet (styles.css) defining semantic design tokens and badge palettes.",
        "Implement TaskGravityApp ES6 class in app.js initializing storageKey and DOM element references.",
        "Define default JSON seed tasks (tg-101 to tg-104) and synchronize with localStorage on load.",
        "Attach event listeners for form submission, input validation (input, change), search, filters, and reset button.",
        "Implement validateField() verifying string lengths, enum whitelists, date boundaries, and numeric ranges.",
        "Upon valid submission, generate unique ID (tg-XXXX), construct JSON object, prepend to array, and call saveTasks().",
        "Implement status cycling (Pending -> In Progress -> Completed -> Pending) and deletion with user confirmation.",
        "Implement getFilteredTasks() applying multi-attribute filtering (status, category, search) and sorting.",
        "Dynamically render tasks using template literals with sanitized HTML escaping (escapeHTML).",
        "Calculate and update live KPI statistics (Total, In Progress, Pending, Velocity %) in DOM.",
        "Deploy to GitHub Pages, verify live functionality, and append deployment links."
    ]
    for i, s in enumerate(steps, 1):
        p = doc.add_paragraph(f"Step {i}: {s}")
        p.paragraph_format.space_after = Pt(2)

    # 6. Source Code
    add_section_heading("6", "Source Code Implementation")

    p = doc.add_paragraph()
    r = p.add_run("6.1 app.js (JSON State, Validation Engine & DOM Manipulation)")
    r.font.bold = True

    with open(r"D:\_TY\Web-Tech\Practicals\Practical2\app.js", "r", encoding="utf-8") as f:
        js_code = f.read()

    p_code = doc.add_paragraph()
    r_code = p_code.add_run(js_code)
    r_code.font.name = 'Consolas'
    r_code.font.size = Pt(8)
    r_code.font.color.rgb = RGBColor(30, 41, 59)
    p_code.paragraph_format.space_before = Pt(4)
    p_code.paragraph_format.space_after = Pt(12)

    p = doc.add_paragraph()
    r = p.add_run("6.2 index.html (Interactive Interface Markup)")
    r.font.bold = True

    with open(r"D:\_TY\Web-Tech\Practicals\Practical2\index.html", "r", encoding="utf-8") as f:
        html_code = f.read()

    p_code2 = doc.add_paragraph()
    r_code2 = p_code2.add_run(html_code)
    r_code2.font.name = 'Consolas'
    r_code2.font.size = Pt(8)
    r_code2.font.color.rgb = RGBColor(30, 41, 59)
    p_code2.paragraph_format.space_before = Pt(4)
    p_code2.paragraph_format.space_after = Pt(12)

    # 7. Verification Results
    add_section_heading("7", "Verification & Test Results")
    test_cases = [
        ("Empty Form Submission", "Submitting form without inputs highlights all required fields in red with descriptive messages.", "Passed"),
        ("Past Date Validation", "Selecting a date earlier than today rejects submission with 'Deadline cannot be scheduled in the past.'", "Passed"),
        ("Effort Hours Boundary", "Entering negative or >100 hours displays 'Estimated effort must be between 1 and 100 hours.'", "Passed"),
        ("Status Lifecycle Cycling", "Clicking checkbox or 'Cycle Status' toggles task from Pending -> In Progress -> Completed.", "Passed"),
        ("Live Filter & Search", "Typing in search bar filters tasks instantaneously without full-page reloads.", "Passed"),
        ("LocalStorage Persistence", "Reloading browser preserves all newly added tasks and completed states from localStorage.", "Passed")
    ]
    t_table = doc.add_table(rows=len(test_cases) + 1, cols=3)
    t_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, h in enumerate(["Test Case", "Observed Behavior & Criteria", "Status"]):
        c = t_table.cell(0, idx)
        c.text = h
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        set_cell_bg(c, "1E293B")
        set_cell_margins(c, 70, 70, 90, 90)

    for r_idx, (tc, obs, st) in enumerate(test_cases, 1):
        c0 = t_table.cell(r_idx, 0)
        c1 = t_table.cell(r_idx, 1)
        c2 = t_table.cell(r_idx, 2)
        c0.text = tc
        c0.paragraphs[0].runs[0].font.bold = True
        c1.text = obs
        c2.text = st
        c2.paragraphs[0].runs[0].font.bold = True
        c2.paragraphs[0].runs[0].font.color.rgb = RGBColor(5, 150, 105)
        set_cell_bg(c0, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
        set_cell_bg(c1, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
        set_cell_bg(c2, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
        set_cell_margins(c0, 60, 60, 80, 80)
        set_cell_margins(c1, 60, 60, 80, 80)
        set_cell_margins(c2, 60, 60, 80, 80)

    doc.add_paragraph()

    # 8. Conclusion
    add_section_heading("8", "Conclusion")
    doc.add_paragraph(
        "The Practical 02 implementation demonstrates the successful creation of an interactive client-side web application "
        "using modern Vanilla JavaScript. By combining structured JSON objects, JavaScript array operations, rigorous input "
        "validation, dynamic DOM rendering, and localStorage synchronization, TaskGravity provides an enterprise-quality task "
        "management experience without external runtime dependencies. The application is hosted live on GitHub Pages, establishing "
        "a reliable foundation for component state abstractions in ReactJS (Practical 03) and backend services (Practicals 04-07)."
    )

    # MANDATORY LINKS
    doc.add_paragraph()
    p_links = doc.add_paragraph()
    r_links_title = p_links.add_run("Project Links & Deployment Information\n")
    r_links_title.font.size = Pt(12)
    r_links_title.font.bold = True
    r_links_title.font.color.rgb = RGBColor(15, 23, 42)

    r_gh = p_links.add_run("GitHub Link / Code Link: ")
    r_gh.font.bold = True
    r_gh_url = p_links.add_run("https://github.com/Sarthak-1229/taskgravity-practical-02\n")
    r_gh_url.font.color.rgb = RGBColor(37, 99, 235)
    r_gh_url.font.underline = True

    r_live = p_links.add_run("Live Deployment Link: ")
    r_live.font.bold = True
    r_live_url = p_links.add_run("https://sarthak-1229.github.io/taskgravity-practical-02/\n")
    r_live_url.font.color.rgb = RGBColor(37, 99, 235)
    r_live_url.font.underline = True

    docx_path = r"D:\_TY\Web-Tech\Practicals\Practical2\Practical_02_Report.docx"
    doc.save(docx_path)
    print(f"Generated DOCX: {docx_path}")

def generate_pdf():
    with open(r"D:\_TY\Web-Tech\Practicals\Practical2\styles.css", "r", encoding="utf-8") as f:
        css_code = f.read()
    with open(r"D:\_TY\Web-Tech\Practicals\Practical2\app.js", "r", encoding="utf-8") as f:
        js_code = f.read()
    with open(r"D:\_TY\Web-Tech\Practicals\Practical2\index.html", "r", encoding="utf-8") as f:
        html_code = f.read()

    escaped_css = html.escape(css_code)
    escaped_js = html.escape(js_code)
    escaped_html = html.escape(html_code)

    report_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Practical 02 Report - TaskGravity</title>
<style>
  @page {{ size: A4; margin: 20mm 18mm 20mm 18mm; }}
  body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; color: #0f172a; line-height: 1.5; font-size: 11pt; margin: 0; padding: 0; }}
  .dept-header {{ text-align: center; border-bottom: 2px solid #0f172a; padding-bottom: 12px; margin-bottom: 18px; }}
  .dept-title {{ font-size: 15pt; font-weight: 800; color: #0f172a; text-transform: uppercase; margin-bottom: 4px; }}
  .dept-sub {{ font-size: 11pt; color: #475569; font-weight: 500; }}
  .practical-title {{ font-size: 14pt; font-weight: 700; color: #2563eb; margin-top: 8px; }}
  .meta-table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 10pt; }}
  .meta-table td {{ border: 1px solid #cbd5e1; padding: 6px 12px; }}
  .meta-label {{ background-color: #f1f5f9; font-weight: 600; color: #334155; width: 18%; }}
  .meta-val {{ font-weight: 600; color: #0f172a; width: 32%; }}
  h2.sec-title {{ font-size: 12pt; font-weight: 700; color: #0f172a; border-left: 4px solid #2563eb; padding-left: 8px; margin-top: 18px; margin-bottom: 8px; text-transform: uppercase; }}
  h3.subsec-title {{ font-size: 11pt; font-weight: 700; color: #1e293b; margin-top: 12px; margin-bottom: 4px; }}
  p, li {{ text-align: justify; color: #334155; margin-bottom: 6px; }}
  ul, ol {{ margin-top: 4px; margin-bottom: 8px; padding-left: 22px; }}
  li {{ margin-bottom: 4px; }}
  .data-table {{ width: 100%; border-collapse: collapse; margin: 12px 0 16px; font-size: 9.5pt; }}
  .data-table th {{ background-color: #1e293b; color: #ffffff; font-weight: 600; padding: 7px 10px; border: 1px solid #1e293b; text-align: left; }}
  .data-table td {{ padding: 6px 10px; border: 1px solid #e2e8f0; color: #334155; }}
  .data-table tr:nth-child(even) td {{ background-color: #f8fafc; }}
  pre.code-block {{ background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px; padding: 10px; font-family: Consolas, monospace; font-size: 8pt; line-height: 1.35; overflow-x: auto; color: #0f172a; page-break-inside: avoid; white-space: pre-wrap; word-break: break-all; }}
  .badge-pass {{ color: #059669; font-weight: 700; }}
  .links-box {{ background-color: #f8fafc; border: 1px solid #cbd5e1; border-left: 4px solid #059669; padding: 12px 16px; margin-top: 20px; page-break-inside: avoid; }}
  .links-box a {{ color: #2563eb; text-decoration: underline; font-weight: 600; }}
</style>
</head>
<body>

<div class="dept-header">
  <div class="dept-title">Department of Computer Engineering (Software Engineering)</div>
  <div class="dept-sub">Web Technology Laboratory — Academic Year 2026-27</div>
  <div class="practical-title">Practical No. 02: Interactive Client-Side Web Application</div>
</div>

<table class="meta-table">
  <tr>
    <td class="meta-label">Name:</td>
    <td class="meta-val">Sarthak Bhosale</td>
    <td class="meta-label">PRN:</td>
    <td class="meta-val">202401100051</td>
  </tr>
  <tr>
    <td class="meta-label">Batch:</td>
    <td class="meta-val">TY SE-A2</td>
    <td class="meta-label">Roll No:</td>
    <td class="meta-val">A50</td>
  </tr>
</table>

<h2 class="sec-title">1. Aim</h2>
<p>
To develop an interactive client-side web application for a Task Management System named <strong>TaskGravity</strong> using HTML5, CSS3, and JavaScript, demonstrating the usage of JSON and Arrays for data storage, comprehensive client-side form validation with visual error feedback, dynamic DOM manipulation, and event-driven state persistence via localStorage.
</p>

<h2 class="sec-title">2. Objectives</h2>
<ul>
  <li>To store, manipulate, and query structured application data using JSON objects and JavaScript arrays.</li>
  <li>To implement strict client-side validation rules (string length bounds, date validity, effort boundaries, category/priority constraints).</li>
  <li>To render real-time and submit-time error messages dynamically without relying on generic browser popups.</li>
  <li>To dynamically update the Document Object Model (DOM) in response to user actions (Add, Cycle Status, Delete, Search, Filter, Sort).</li>
  <li>To compute and display live aggregate metrics (Total Backlog, In Progress, Pending Review, Completion Velocity %).</li>
  <li>To achieve local data persistence across page reloads using the browser's localStorage API.</li>
  <li>To deploy the application to GitHub Pages and document repository and live deployment links.</li>
</ul>

<h2 class="sec-title">3. Theory & Concepts</h2>
<p>
Modern frontend architectures depend heavily on client-side state management, responsive event handling, and rigorous input sanitization. In Practical 02, TaskGravity implements an interactive, event-driven architecture:
</p>

<h3 class="subsec-title">3.1 JSON and Array Data Architecture</h3>
<p>
JavaScript Object Notation (JSON) is a lightweight text-based data interchange format. In TaskGravity, tasks are modeled as structured JSON records containing unique IDs, titles, categories, priorities, deadlines, estimated effort hours, and lifecycle statuses. These objects reside within an in-memory JavaScript array that serves as the single source of truth for the client. Changes are serialized to JSON strings using <code>JSON.stringify()</code> and deserialized upon reload using <code>JSON.parse()</code> via the Web Storage API (<code>localStorage</code>).
</p>

<h3 class="subsec-title">3.2 Defensive Input Validation & Error Messaging</h3>
<p>
Client-side validation enhances user experience by catching malformed or missing data before storage or transmission. TaskGravity evaluates: (1) Title presence and character bounds [3-80 chars]; (2) Architectural category whitelist; (3) Priority whitelist; (4) Deadline validity and past-date rejection (comparing against today's midnight timestamp); (5) Numeric effort hours bounded between 0.5 and 100. Field-specific <code>.error-feedback</code> elements render immediately upon keystroke (<code>input</code> event) or blur (<code>change</code> event), preventing user frustration.
</p>

<h3 class="subsec-title">3.3 Dynamic DOM Manipulation & Event Handling</h3>
<p>
The Document Object Model (DOM) represents HTML as a tree of objects. TaskGravity binds event listeners (<code>addEventListener</code>) to form submissions, keystrokes, filter selects, and status toggle buttons. Rather than whole-page refreshes, the DOM updates efficiently using template literals, preserving focus and delivering a fluid experience.
</p>

<h3 class="subsec-title">3.4 Querying, Filtering, and Sorting Algorithms</h3>
<p>
Tasks are filtered dynamically using JavaScript array primitives (<code>filter</code>, <code>sort</code>). The application enables compound filtering (e.g. status AND category AND search substring) and multi-attribute sorting (by deadline, priority weight, effort hours, or alphabetical title), with automatic recalculation of KPI progress percentages.
</p>

<h2 class="sec-title">4. Input Validation Rules Matrix</h2>
<p>The validation rules enforced prior to JSON array storage are detailed below:</p>

<table class="data-table">
  <thead>
    <tr>
      <th>Field Name</th>
      <th>Data Type</th>
      <th>Validation Constraints</th>
      <th>Error Feedback Triggered</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>Task Title</strong></td>
      <td>String</td>
      <td>Required, trimmed length &ge; 3 and &le; 80 chars</td>
      <td>Task Title must be between 3 and 80 characters.</td>
    </tr>
    <tr>
      <td><strong>Category</strong></td>
      <td>Enum String</td>
      <td>Must match whitelist: Frontend, Backend, Database, DevOps, Testing, UI/UX</td>
      <td>Please select a valid architectural category.</td>
    </tr>
    <tr>
      <td><strong>Priority</strong></td>
      <td>Enum String</td>
      <td>Must match whitelist: Low, Medium, High, Urgent</td>
      <td>Please select an operational priority level.</td>
    </tr>
    <tr>
      <td><strong>Deadline</strong></td>
      <td>ISO Date</td>
      <td>Required, valid format, strictly &ge; current date</td>
      <td>Target deadline is required / Cannot be in the past.</td>
    </tr>
    <tr>
      <td><strong>Effort Hours</strong></td>
      <td>Number (Float)</td>
      <td>Required, numeric, strictly &gt; 0 and &le; 100 hours</td>
      <td>Estimated effort must be between 1 and 100 hours.</td>
    </tr>
    <tr>
      <td><strong>Description</strong></td>
      <td>String</td>
      <td>Optional, maximum length &le; 300 chars</td>
      <td>Description cannot exceed 300 characters.</td>
    </tr>
  </tbody>
</table>

<h2 class="sec-title">5. Algorithm & Implementation Steps</h2>
<ol>
  <li>Construct responsive dual-column interface (<code>index.html</code>) with form panel and dynamic stream.</li>
  <li>Author custom flat stylesheet (<code>styles.css</code>) defining semantic design tokens and badge palettes.</li>
  <li>Implement <code>TaskGravityApp</code> ES6 class in <code>app.js</code> initializing storageKey and DOM element references.</li>
  <li>Define default JSON seed tasks (tg-101 to tg-104) and synchronize with <code>localStorage</code> on load.</li>
  <li>Attach event listeners for form submission, input validation (input, change), search, filters, and reset button.</li>
  <li>Implement <code>validateField()</code> verifying string lengths, enum whitelists, date boundaries, and numeric ranges.</li>
  <li>Upon valid submission, generate unique ID (tg-XXXX), construct JSON object, prepend to array, and call <code>saveTasks()</code>.</li>
  <li>Implement status cycling (Pending &rarr; In Progress &rarr; Completed &rarr; Pending) and deletion with user confirmation.</li>
  <li>Implement <code>getFilteredTasks()</code> applying multi-attribute filtering (status, category, search) and sorting.</li>
  <li>Dynamically render tasks using template literals with sanitized HTML escaping (<code>escapeHTML</code>).</li>
  <li>Calculate and update live KPI statistics (Total, In Progress, Pending, Velocity %) in DOM.</li>
  <li>Deploy to GitHub Pages, verify live functionality, and append deployment links.</li>
</ol>

<h2 class="sec-title">6. Source Code Implementation</h2>

<h3 class="subsec-title">6.1 app.js (JSON State, Validation Engine & DOM Manipulation)</h3>
<pre class="code-block">{escaped_js}</pre>

<h3 class="subsec-title">6.2 index.html (Interactive Interface Markup)</h3>
<pre class="code-block">{escaped_html}</pre>

<h3 class="subsec-title">6.3 styles.css (Minimalist Custom Stylesheet)</h3>
<pre class="code-block">{escaped_css}</pre>

<h2 class="sec-title">7. Verification & Test Results</h2>

<table class="data-table">
  <thead>
    <tr>
      <th>Test Case</th>
      <th>Observed Behavior & Criteria</th>
      <th>Status</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>Empty Form Submission</strong></td>
      <td>Submitting form without inputs highlights all required fields in red with descriptive messages.</td>
      <td class="badge-pass">Passed</td>
    </tr>
    <tr>
      <td><strong>Past Date Validation</strong></td>
      <td>Selecting a date earlier than today rejects submission with 'Deadline cannot be scheduled in the past.'</td>
      <td class="badge-pass">Passed</td>
    </tr>
    <tr>
      <td><strong>Effort Hours Boundary</strong></td>
      <td>Entering negative or &gt;100 hours displays 'Estimated effort must be between 1 and 100 hours.'</td>
      <td class="badge-pass">Passed</td>
    </tr>
    <tr>
      <td><strong>Status Lifecycle Cycling</strong></td>
      <td>Clicking checkbox or 'Cycle Status' toggles task from Pending &rarr; In Progress &rarr; Completed.</td>
      <td class="badge-pass">Passed</td>
    </tr>
    <tr>
      <td><strong>Live Filter & Search</strong></td>
      <td>Typing in search bar filters tasks instantaneously without full-page reloads.</td>
      <td class="badge-pass">Passed</td>
    </tr>
    <tr>
      <td><strong>LocalStorage Persistence</strong></td>
      <td>Reloading browser preserves all newly added tasks and completed states from localStorage.</td>
      <td class="badge-pass">Passed</td>
    </tr>
  </tbody>
</table>

<h2 class="sec-title">8. Conclusion</h2>
<p>
The Practical 02 implementation demonstrates the successful creation of an interactive client-side web application using modern Vanilla JavaScript. By combining structured JSON objects, JavaScript array operations, rigorous input validation, dynamic DOM rendering, and localStorage synchronization, TaskGravity provides an enterprise-quality task management experience without external runtime dependencies. The application is hosted live on GitHub Pages, establishing a reliable foundation for component state abstractions in ReactJS (Practical 03) and backend services (Practicals 04-07).
</p>

<div class="links-box">
  <div style="font-weight: 700; font-size: 11pt; margin-bottom: 6px; color: #0f172a;">Project Links & Deployment Information</div>
  <div><strong>GitHub Link / Code Link:</strong> <a href="https://github.com/Sarthak-1229/taskgravity-practical-02" target="_blank">https://github.com/Sarthak-1229/taskgravity-practical-02</a></div>
  <div style="margin-top: 4px;"><strong>Live Deployment Link:</strong> <a href="https://sarthak-1229.github.io/taskgravity-practical-02/" target="_blank">https://sarthak-1229.github.io/taskgravity-practical-02/</a></div>
</div>

</body>
</html>
"""
    html_report_path = r"D:\_TY\Web-Tech\Practicals\Practical2\report_temp.html"
    pdf_report_path = r"D:\_TY\Web-Tech\Practicals\Practical2\Practical_02_Report.pdf"

    with open(html_report_path, "w", encoding="utf-8") as f:
        f.write(report_html)

    chrome_cmd = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "--headless=new",
        "--no-pdf-header-footer",
        f"--print-to-pdf={pdf_report_path}",
        html_report_path
    ]
    subprocess.run(chrome_cmd, check=True)

    if os.path.exists(html_report_path):
        os.remove(html_report_path)

    print(f"Generated PDF: {pdf_report_path}")

if __name__ == "__main__":
    generate_docx()
    generate_pdf()
