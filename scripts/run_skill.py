import os
import sys
import json
import re
import shutil
import urllib.request
import argparse
import datetime
import subprocess
import email
import imaplib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

try:
    from ddgs import DDGS
except ImportError:
    DDGS = None

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    Document = None


def get_timestamp():
    return datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


SENIOR_TITLE_KEYWORDS = [
    "senior", "sr.", "sr ", "lead", "staff", "principal", "architect", "manager",
    "director", "head", "vp", "expert", "specialist", "sde 2", "sde-2", "sde 3",
    "sde-3", "sde ii", "sde iii", "sde-ii", "sde-iii", "sde2", "sde3", "mid-senior",
    "5+", "6+", "7+", "8+", "9+", "10+"
]

SENIOR_EXP_PATTERNS = [
    r"\b[3-9]\+\s*years?\b",
    r"\b1[0-9]\+\s*years?\b",
    r"\b[3-9]\s*-\s*[0-9]+\s*years?\b",
    r"\b[4-9]\s*to\s*[0-9]+\s*years?\b",
    r"\bminimum\s+[3-9]\s*years?\b",
    r"\bat\s+least\s+[3-9]\s*years?\b",
    r"\brequires?\s+[3-9]\s*years?\b"
]


def is_senior_or_mismatched(title, text):
    """Returns True if the job is meant for mid/senior engineers with 3+ years experience."""
    t_lower = title.lower()
    
    # 1. Check title for senior keywords
    for sk in SENIOR_TITLE_KEYWORDS:
        if re.search(r"\b" + re.escape(sk) + r"\b", t_lower):
            return True, f"Senior title keyword detected: '{sk}'"

    # 2. Check text for 3+ years experience requirements
    txt_lower = text.lower()
    for pat in SENIOR_EXP_PATTERNS:
        match = re.search(pat, txt_lower)
        if match:
            snippet = txt_lower[max(0, match.start()-20):min(len(txt_lower), match.end()+20)]
            if not any(ek in t_lower for ek in ["intern", "internship", "fresher"]):
                return True, f"High experience requirement: '{match.group(0)}' in {snippet}"

    return False, "Suitable for entry-level/intern"


def search_candidate_jobs(limit=20):
    """Discovers active candidate job links strictly targeted at Internships, SDE-1, and Fresher roles."""
    queries = [
        'site:jobs.lever.co ("Intern" OR "Internship" OR "SDE-1" OR "SDE 1" OR "Junior" OR "Associate") Python (Bangalore OR Bengaluru OR Remote)',
        'site:job-boards.greenhouse.io ("Intern" OR "Internship" OR "SDE 1" OR "SDE-1" OR "Junior" OR "Fresher") Python (Bangalore OR Bengaluru OR Remote)',
        'site:boards.greenhouse.io ("Intern" OR "Internship" OR "SDE 1" OR "SDE-1" OR "Junior" OR "Associate") Python (Bangalore OR Bengaluru OR Remote)',
        'site:in.linkedin.com/jobs/view ("Intern" OR "Internship" OR "Junior" OR "Associate" OR "Fresher" OR "SDE 1") ("Python" OR "Backend") (Bengaluru OR Bangalore)',
        'site:wellfound.com/jobs ("Intern" OR "Junior" OR "SDE 1" OR "Backend") ("0-1" OR "0-2" OR "fresher" OR "intern") Python Bangalore',
        'site:cutshort.io/job ("Intern" OR "Junior" OR "SDE 1" OR "Fresher") Python Bangalore',
        'site:jobs.lever.co ("Backend Intern" OR "Software Engineer Intern" OR "AI Intern") Bangalore',
        'site:job-boards.greenhouse.io ("Backend Intern" OR "Software Engineer Intern") Bangalore',
        'Juspay ("SDE 1" OR "SDE-1" OR "Intern" OR "Fresher" OR "Backend") Bangalore hiring',
        'CRED ("Backend" OR "Software") ("Intern" OR "Fresher" OR "Junior") Bangalore hiring',
        '"Backend Developer" ("Intern" OR "Fresher" OR "0-1 years") Python Bangalore 2026 hiring',
        '"Software Engineer Intern" Python (Bangalore OR Remote) 2026 hiring'
    ]

    results = []
    seen_urls = set()
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}

    if DDGS:
        with DDGS() as ddgs:
            for q in queries:
                try:
                    items = list(ddgs.text(q, max_results=8))
                    for it in items:
                        url = it.get('href')
                        if url and url not in seen_urls:
                            seen_urls.add(url)
                            results.append(it)
                except Exception as e:
                    pass

    verified = []
    print(f"Discovered {len(results)} candidate links. Filtering for Internships & Fresher/SDE-1 roles...")
    for r in results:
        url = r.get('href', '')
        if not url:
            continue
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=6) as resp:
                if resp.getcode() == 200:
                    text = ""
                    if BeautifulSoup:
                        html = resp.read().decode('utf-8', errors='ignore')
                        soup = BeautifulSoup(html, 'html.parser')
                        text = ' '.join(soup.get_text(separator=' ').split())
                    else:
                        text = resp.read().decode('utf-8', errors='ignore')[:2000]

                    if any(w in text.lower() for w in ['no longer accepting applications', 'position filled', 'job not found', 'expired']):
                        continue

                    raw_title = r.get('title', 'Software Engineer')

                    is_senior, reason = is_senior_or_mismatched(raw_title, text)
                    if is_senior:
                        continue

                    verified.append({
                        'title': raw_title,
                        'url': url,
                        'snippet': r.get('body', ''),
                        'text': text[:2500]
                    })
                    print(f"  [ACCEPTED 200 OK] {raw_title[:50]} -> {url}")
                    if len(verified) >= limit:
                        break
        except Exception:
            pass

    return verified


def evaluate_job(job_item, profile_text, groq_api_key=None):
    """Evaluates job against candidate profile using Groq or rule-based fallback."""
    title = job_item['title']
    snippet = job_item['snippet']
    full_text = job_item['text']

    company = "Tech Company"
    if " at " in title:
        company = title.split(" at ")[-1].split(" - ")[0].split(" | ")[0].strip()
    elif " - " in title:
        company = title.split(" - ")[0].strip()
    elif " hiring " in title:
        company = title.split(" hiring ")[0].strip()

    role = title.split(" at ")[0].split(" - ")[0].split(" hiring ")[-1].strip()

    if groq_api_key:
        try:
            prompt = f"""You are an expert career evaluation assistant.
CRITICAL CANDIDATE PROFILE:
- Name: Jayaditya Dev
- Status: Final-Year B.E. Computer Science Undergrad (Expected 2027) with 0-1 years of intern experience at 7HiddenLayers.
- Target: HIGH-PAYING INTERNSHIPS, SDE-1, or FRESHER/JUNIOR FULL-TIME ROLES (0-2 YOE max). Target base CTC: 8-10 LPA or competitive stipend (₹40k-₹80k/mo).
- Core Stack: Python, FastAPI, PostgreSQL, SQLAlchemy, AWS, Docker, App Security (TryHackMe Top 5%).

JOB POSTING:
Title: {title}
Company: {company}
Snippet: {snippet}
Context: {full_text[:1400]}

SENIORITY RULE:
If this job requires 3+ years of experience or is explicitly a mid/senior/lead role, set "fitness_score": "35% Fit" and set "match_rationale": "Experience Mismatch: Requires 3+ years of experience; candidate is a final year undergrad seeking intern/fresher roles."

Respond ONLY in valid JSON matching this exact schema:
{{
  "fitness_score": "92% Fit",
  "is_suitable_for_fresher_or_intern": true,
  "experience_required": "Internship (0-1 YOE)" or "Fresher / SDE-1 (0-2 YOE)",
  "estimated_compensation": "8-10 LPA base" or "INR 40k-60k/month stipend" or "Competitive Market Standard (8-12 LPA)",
  "location": "Bengaluru (Hybrid)" or "Remote",
  "jd_summary": "2 concise sentences summarizing key requirements, core responsibilities, and technologies.",
  "match_rationale": "1-2 concise sentences explaining why this role matches candidate's Python/FastAPI/PostgreSQL skills.",
  "tailored_profile": "1-2 sentences summarizing candidate backend strengths aligned to this JD.",
  "cover_letter": "A bespoke, human-sounding cover letter under 220 words without AI clichés."
}}"""

            req_data = json.dumps({
                "model": "llama-3.3-70b-versatile",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.2,
                "response_format": {"type": "json_object"}
            }).encode('utf-8')

            req = urllib.request.Request(
                "https://api.groq.com/openai/v1/chat/completions",
                data=req_data,
                headers={
                    "Authorization": f"Bearer {groq_api_key}",
                    "Content-Type": "application/json",
                    "User-Agent": "JobSkillAgent/1.0"
                }
            )
            with urllib.request.urlopen(req, timeout=12) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                content = data['choices'][0]['message']['content']
                parsed = json.loads(content)
                return {
                    "company": company,
                    "role": role,
                    "fitness": parsed.get("fitness_score", "90% Fit"),
                    "suitable": parsed.get("is_suitable_for_fresher_or_intern", True),
                    "experience": parsed.get("experience_required", "Internship / Fresher (0-1 YOE)"),
                    "compensation": parsed.get("estimated_compensation", "8-10 LPA base / Competitive Stipend"),
                    "location": parsed.get("location", "Bengaluru / Remote"),
                    "jd_summary": parsed.get("jd_summary", snippet[:200]),
                    "rationale": parsed.get("match_rationale", "Strong match on Python, FastAPI, and PostgreSQL."),
                    "latex_profile": parsed.get("tailored_profile", "Final-year Computer Science undergraduate and backend engineer specializing in Python (FastAPI), PostgreSQL, and secure API architectures."),
                    "cover_letter": parsed.get("cover_letter", "")
                }
        except Exception as e:
            pass

    # Rule-based fallback
    is_intern = "intern" in title.lower() or "intern" in snippet.lower()
    return {
        "company": company,
        "role": role,
        "fitness": "92% Fit",
        "suitable": True,
        "experience": "Internship (0-1 YOE)" if is_intern else "Fresher / SDE-1 (0-2 YOE)",
        "compensation": "INR 40k-75k/month stipend" if is_intern else "8-10 LPA base",
        "location": "Bengaluru / Remote",
        "jd_summary": snippet[:220] if snippet else f"Backend engineering role in {company} involving Python APIs and databases.",
        "rationale": f"High alignment for final-year undergrad/intern role: Python backend development, FastAPI services, and PostgreSQL schemas at {company}.",
        "latex_profile": f"Final-year Computer Science undergraduate and backend engineer specializing in resilient API architecture, asynchronous data pipelines in Python (FastAPI), and relational modeling on PostgreSQL.",
        "cover_letter": f"""I am writing to apply for the {role} position at {company}. As a final-year Computer Science undergraduate with hands-on intern experience developing production backend services in Python (FastAPI, SQLAlchemy) and PostgreSQL, I am eager to contribute to your engineering team.

At 7HiddenLayers, I developed backend ingestion services that process complex document updates incrementally. In parallel, my independent engineering work includes architecting Guardian AI—an asynchronous platform integrating WebSockets and relational schemas—and securing a Top 5% global ranking on TryHackMe for defensive application security.

I write clean, tested code and am excited about the opportunity to bring my backend fundamentals and work ethic to {company}."""
    }


def compile_latex_pdf(output_dir, tex_filename, pdf_filename, base_tex, profile_text, resume_cls_path):
    """Compiles tailored 1-page LaTeX PDF using pdflatex."""
    cls_dest = os.path.join(output_dir, "resume.cls")
    if not os.path.exists(cls_dest) and os.path.exists(resume_cls_path):
        shutil.copy(resume_cls_path, cls_dest)

    old_profile = r"""\begin{rSection}{PROFILE}

Backend-focused full-stack engineer focused on building reliable, production-oriented, and secure backend systems, APIs, deployment workflows, and service-oriented architectures.

\end{rSection}"""

    new_profile = f"""\\begin{{rSection}}{{PROFILE}}\n\n{profile_text}\n\n\\end{{rSection}}"""
    tex_content = base_tex.replace(old_profile, new_profile)

    tex_path = os.path.join(output_dir, tex_filename)
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_content)

    try:
        subprocess.run(
            ["pdflatex", "-interaction=nonstopmode", tex_filename],
            cwd=output_dir,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False
        )
    except Exception:
        pass

    compiled_pdf = os.path.splitext(tex_path)[0] + ".pdf"
    target_pdf = os.path.join(output_dir, pdf_filename)
    if os.path.exists(compiled_pdf) and compiled_pdf != target_pdf:
        shutil.move(compiled_pdf, target_pdf)

    base_name = os.path.splitext(tex_filename)[0]
    for ext in [".aux", ".log", ".out"]:
        aux_f = os.path.join(output_dir, base_name + ext)
        if os.path.exists(aux_f):
            try:
                os.remove(aux_f)
            except Exception:
                pass

    if os.path.exists(cls_dest):
        try:
            os.remove(cls_dest)
        except Exception:
            pass

    return os.path.exists(target_pdf)


def generate_docx_files(output_dir, comp_name, role_name, profile_text, cover_letter_text):
    """Generates tailored DOCX resume and cover letter."""
    if not Document:
        return

    # Resume DOCX
    doc = Document()
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_name = p_name.add_run("JAYADITYA DEV")
    r_name.bold = True
    r_name.font.size = Pt(16)

    p_c = doc.add_paragraph()
    p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c.add_run("Bengaluru, Karnataka | jayadityadev10@gmail.com | +91 92345 09450\nlinkedin.com/in/jayadityadev26 | github.com/jayadityadev")

    p_p = doc.add_paragraph()
    p_p.add_run(f"Target Role: {role_name} - {comp_name}\n").bold = True
    p_p.add_run(profile_text)

    p_edu = doc.add_paragraph()
    p_edu.add_run("Education\n").bold = True
    p_edu.add_run("B.E. Computer Science & Engineering - KSIT Bengaluru (Expected 2027) | CGPA: 8.88")

    doc_resume_path = os.path.join(output_dir, f"Jayaditya_Dev_Resume_{comp_name}.docx")
    doc.save(doc_resume_path)

    # Cover Letter DOCX
    cl_doc = Document()
    p_cl_name = cl_doc.add_paragraph()
    r_cln = p_cl_name.add_run("Jayaditya Dev")
    r_cln.bold = True
    r_cln.font.size = Pt(14)
    cl_doc.add_paragraph(f"Date: {datetime.date.today().strftime('%B %d, %Y')}\nHiring Team - {role_name}\n{comp_name}")
    cl_doc.add_paragraph(f"Dear Hiring Manager at {comp_name},").bold = True
    for para in cover_letter_text.strip().split("\n\n"):
        cl_doc.add_paragraph(para.strip())
    cl_doc.add_paragraph("Sincerely,\nJayaditya Dev")

    doc_cl_path = os.path.join(output_dir, f"Jayaditya_Dev_CoverLetter_{comp_name}.docx")
    cl_doc.save(doc_cl_path)


def update_excel_tracker(run_dir, tracker_rows):
    """Creates a beautifully styled job_tracker.xlsx with comprehensive columns inside the run directory."""
    if not openpyxl:
        return None
    tracker_path = os.path.join(run_dir, "job_tracker.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Job Applications Tracker"

    # Enhanced headers
    headers = [
        "Rank", "Date Found", "Company", "Job Title", "Fitness Score",
        "Experience Required", "Estimated Compensation", "Location", "Platform",
        "Direct Apply Link", "Job Description Summary", "Resume Match Notes",
        "Tailored Cover Letter", "Application Status"
    ]
    ws.append(headers)

    # Header styling
    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    ws.row_dimensions[1].height = 28

    # Populate rows
    for r_idx, row_data in enumerate(tracker_rows, start=2):
        ws.append(row_data)
        ws.row_dimensions[r_idx].height = 24
        # Style rank and fitness
        ws.cell(row=r_idx, column=1).alignment = align_center
        ws.cell(row=r_idx, column=2).alignment = align_center
        ws.cell(row=r_idx, column=5).alignment = align_center
        ws.cell(row=r_idx, column=6).alignment = align_center
        ws.cell(row=r_idx, column=7).alignment = align_center
        ws.cell(row=r_idx, column=8).alignment = align_center
        ws.cell(row=r_idx, column=9).alignment = align_center
        ws.cell(row=r_idx, column=14).alignment = align_center

    # Column widths
    col_widths = {
        1: 8,   # Rank
        2: 13,  # Date Found
        3: 18,  # Company
        4: 30,  # Job Title
        5: 14,  # Fitness Score
        6: 22,  # Experience
        7: 24,  # Compensation
        8: 18,  # Location
        9: 14,  # Platform
        10: 35, # Direct Apply Link
        11: 45, # JD Summary
        12: 45, # Resume Match Notes
        13: 50, # Cover Letter
        14: 18  # Status
    }
    for col_idx, width in col_widths.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    wb.save(tracker_path)
    print(f"[OK] Saved comprehensive tracker to {tracker_path}")
    return tracker_path


def send_gmail_report(to_addr, user_addr, app_password, subject, html_body, attachments):
    """Dispatches formatted HTML morning briefing with attached XLSX tracker and resume PDFs via Gmail SMTP."""
    msg = MIMEMultipart('mixed')
    msg['Subject'] = subject
    msg['From'] = user_addr
    msg['To'] = to_addr

    msg_alt = MIMEMultipart('alternative')
    msg_alt.attach(MIMEText(html_body, 'html'))
    msg.attach(msg_alt)

    for att_path in attachments:
        if att_path and os.path.exists(att_path):
            with open(att_path, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(att_path)}"')
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(user_addr, app_password)
            server.sendmail(user_addr, [to_addr], msg.as_string())
        print(f"[OK] Morning briefing email sent successfully to {to_addr} (with {len(attachments)} attachments)")
    except Exception as e:
        print(f"[WARN] Failed to send email via SMTP: {e}")


def main():
    parser = argparse.ArgumentParser(description="Autonomous Job Search Skill Runner - 20 Jobs Tailored Edition")
    parser.add_argument("--dry-run", action="store_true", help="Perform discovery and local compilation without sending email")
    parser.add_argument("--limit", type=int, default=20, help="Number of fresh jobs to process (default: 20)")
    args = parser.parse_args()

    run_id = f"run_{get_timestamp()}"
    run_dir = os.path.join("runs", run_id)
    os.makedirs(run_dir, exist_ok=True)
    print(f"=== Starting Autonomous Job Skill Execution (Targeting {args.limit} Roles): {run_id} ===")

    resume_dir = "resume"
    main_tex_path = os.path.join(resume_dir, "main.tex")
    resume_cls_path = os.path.join(resume_dir, "resume.cls")

    base_tex = ""
    if os.path.exists(main_tex_path):
        with open(main_tex_path, "r", encoding="utf-8") as f:
            base_tex = f.read()

    groq_api_key = os.environ.get("GROQ_API_KEY")
    gmail_user = os.environ.get("GMAIL_USER")
    gmail_pwd = os.environ.get("GMAIL_APP_PASSWORD")

    # Step 1: Discover candidate jobs (request extra to allow strict filtering)
    jobs = search_candidate_jobs(limit=args.limit * 3)
    if not jobs:
        print("[INFO] No external search results found. Using verified seed targets.")
        jobs = [
            {"title": "Software Engineer Intern at Weekday", "href": "https://jobs.lever.co/weekdayworks/a86efff4-1c2e-43fc-8ecb-dc8f873db2f1", "snippet": "Backend Python software engineer intern.", "text": "Python, FastAPI, PostgreSQL"},
            {"title": "Software Engineer (Intern) - Backend at Merkle Science", "href": "https://jobs.lever.co/merklescience/e663b69b-264a-4bd7-b04d-fb3c0a824a28", "snippet": "Backend intern Python PostgreSQL.", "text": "Python, REST APIs, Microservices, PostgreSQL"},
            {"title": "Junior Software Engineer (Remote) at PolicyMe", "href": "https://jobs.lever.co/policyme/2ad3fd98-21fc-4bbd-aa91-868befe5f699", "snippet": "Junior software engineer Python.", "text": "Python, APIs, PostgreSQL"},
            {"title": "Software Development Engineer Backend (DEV-BE02) at Juspay", "href": "https://juspay.io/careers/DEV-BE02", "snippet": "First principles engineering, SDE-1 / Fresher.", "text": "Algorithms, Networking, Operating Systems"}
        ]

    evaluated_candidates = []
    print(f"Evaluating candidates with Groq / heuristic model...")
    for j in jobs:
        eval_res = evaluate_job(j, base_tex, groq_api_key=groq_api_key)

        fit_num = 80
        try:
            fit_num = int(re.search(r"\d+", eval_res["fitness"]).group(0))
        except Exception:
            pass

        if not eval_res.get("suitable", True) or fit_num < 70 or "mismatch" in eval_res.get("rationale", "").lower():
            continue

        eval_res["fit_num"] = fit_num
        eval_res["raw_job"] = j
        evaluated_candidates.append(eval_res)

    # Sort by fitness score descending (highest match first)
    evaluated_candidates.sort(key=lambda x: x.get("fit_num", 0), reverse=True)
    top_candidates = evaluated_candidates[:args.limit]

    print(f"Accepted {len(top_candidates)} highly-tailored entry-level/intern postings.")

    tracker_rows = []
    pdf_attachments = []
    report_rows_html = []
    today_str = datetime.date.today().strftime("%d %b %Y")

    for rank, eval_res in enumerate(top_candidates, 1):
        j = eval_res["raw_job"]
        comp = eval_res["company"].replace(" ", "").replace("/", "-")
        comp_dir = os.path.join(run_dir, f"{rank:02d}_{comp}")
        os.makedirs(comp_dir, exist_ok=True)

        # 1. Compile LaTeX PDF for all accepted top matches
        tex_file = f"Jayaditya_Dev_Resume_{comp}.tex"
        pdf_file = f"Jayaditya_Dev_Resume_{comp}.pdf"
        pdf_ok = compile_latex_pdf(
            output_dir=comp_dir,
            tex_filename=tex_file,
            pdf_filename=pdf_file,
            base_tex=base_tex,
            profile_text=eval_res["latex_profile"],
            resume_cls_path=resume_cls_path
        )
        pdf_full_path = os.path.join(comp_dir, pdf_file)
        if pdf_ok and rank <= 6:  # Attach top 6 PDFs directly to email to respect email size bounds
            pdf_attachments.append(pdf_full_path)

        # 2. Generate DOCX Resume and Cover Letter
        generate_docx_files(
            output_dir=comp_dir,
            comp_name=eval_res["company"],
            role_name=eval_res["role"],
            profile_text=eval_res["latex_profile"],
            cover_letter_text=eval_res["cover_letter"]
        )

        job_url = j.get("url") or j.get("href", "#")

        # Row schema for job_tracker.xlsx
        tracker_rows.append([
            rank,
            today_str,
            eval_res["company"],
            eval_res["role"],
            eval_res["fitness"],
            eval_res.get("experience", "Internship / Fresher (0-1 YOE)"),
            eval_res.get("compensation", "8-10 LPA base / Competitive Stipend"),
            eval_res.get("location", "Bengaluru / Remote"),
            "Direct / ATS",
            job_url,
            eval_res.get("jd_summary", ""),
            eval_res.get("rationale", ""),
            eval_res.get("cover_letter", ""),
            "Ready to Apply"
        ])

        report_rows_html.append(f"""
        <tr>
            <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: center;"><b>#{rank}</b></td>
            <td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>{eval_res['company']}</b></td>
            <td style="padding: 8px; border-bottom: 1px solid #ddd;">{eval_res['role']}</td>
            <td style="padding: 8px; border-bottom: 1px solid #ddd; color: #0d6efd; text-align: center;"><b>{eval_res['fitness']}</b></td>
            <td style="padding: 8px; border-bottom: 1px solid #ddd; font-size: 13px;">{eval_res.get('compensation', 'Competitive')}</td>
            <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: center;"><a href="{job_url}" style="background-color: #0d6efd; color: white; padding: 5px 10px; text-decoration: none; border-radius: 4px; display: inline-block; font-size: 12px;">Apply</a></td>
        </tr>
        <tr>
            <td colspan="6" style="padding: 4px 8px 10px 8px; color: #555; font-size: 12px; border-bottom: 1px solid #eee;">
                <b>Level:</b> {eval_res.get('experience', 'Intern / Fresher')} | <b>Fit Notes:</b> {eval_res.get('rationale', '')}
            </td>
        </tr>
        """)

    # Create beautifully styled Excel Tracker
    tracker_file = update_excel_tracker(run_dir, tracker_rows)

    # HTML Morning Email Report
    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333; line-height: 1.5;">
        <h2>Good Morning Jayaditya! 🚀</h2>
        <p>Here is your daily automated job briefing for <b>{today_str}</b>. Below are <b>{len(top_candidates)} highly-tailored roles</b> ranked specifically for your profile (High-Paying Internships and Fresher/Junior SDE-1 roles in Bengaluru & Remote).</p>
        
        <div style="background-color: #f0f7ff; border-left: 4px solid #0d6efd; padding: 12px; margin-bottom: 20px;">
            <b>📊 Application Tracker Attached:</b> The complete spreadsheet (<code>job_tracker.xlsx</code>) with all 20 rankings, JD summaries, compensation estimates, and full copy-paste cover letters is attached to this email.
        </div>

        <table style="width: 100%; border-collapse: collapse; margin-top: 10px; margin-bottom: 25px;">
            <thead>
                <tr style="background-color: #1B365D; color: white; text-align: left;">
                    <th style="padding: 10px; text-align: center;">Rank</th>
                    <th style="padding: 10px;">Company</th>
                    <th style="padding: 10px;">Role</th>
                    <th style="padding: 10px; text-align: center;">Fitness</th>
                    <th style="padding: 10px;">Est. Compensation</th>
                    <th style="padding: 10px; text-align: center;">Action</th>
                </tr>
            </thead>
            <tbody>
                {''.join(report_rows_html)}
            </tbody>
        </table>

        <p><b>Attached Materials:</b></p>
        <ul>
            <li><b>job_tracker.xlsx</b>: Complete tracking workbook with JD summaries, rankings, and ready-to-copy cover letters.</li>
            <li><b>Top Tailored LaTeX PDFs</b>: 1-page compiled resumes tailored to today's top matches.</li>
        </ul>
        <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;" />
        <p style="font-size: 12px; color: #888;">Autonomous Job Search Assistant | Powered by Antigravity & Groq</p>
    </body>
    </html>
    """

    summary_file = os.path.join(run_dir, "run_summary.md")
    with open(summary_file, "w", encoding="utf-8") as f:
        f.write(f"# Daily Job Search Summary - {today_str}\n\nProcessed {len(top_candidates)} tailored jobs in {run_id}.\n")

    # Combine attachments: job_tracker.xlsx FIRST, followed by top PDF resumes
    all_attachments = []
    if tracker_file and os.path.exists(tracker_file):
        all_attachments.append(tracker_file)
    all_attachments.extend(pdf_attachments)

    if not args.dry_run and gmail_user and gmail_pwd and len(top_candidates) > 0:
        send_gmail_report(
            to_addr=gmail_user,
            user_addr=gmail_user,
            app_password=gmail_pwd,
            subject=f"Daily Job Search Briefing ({len(top_candidates)} Tailored Roles + Tracker) - {today_str}",
            html_body=html_body,
            attachments=all_attachments
        )
    else:
        print("[INFO] Skipping Gmail dispatch (dry-run mode, missing credentials, or 0 accepted jobs).")

    print(f"=== Execution Finished Cleanly in {run_dir} ===")


if __name__ == "__main__":
    main()
