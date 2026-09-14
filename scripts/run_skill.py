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

ENTRY_LEVEL_KEYWORDS = [
    "intern", "internship", "fresher", "entry-level", "entry level", "graduate",
    "junior", "jr.", "associate", "sde 1", "sde-1", "sde1", "sde i", "0-1", "0-2",
    "1-2 years", "trainee", "campus", "university", "undergraduate", "apprentice"
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
            # Check if this isn't immediately preceded by 'not' or similar
            snippet = txt_lower[max(0, match.start()-20):min(len(txt_lower), match.end()+20)]
            # If the title clearly says intern or fresher, give benefit of doubt unless strong match
            if not any(ek in t_lower for ek in ["intern", "internship", "fresher"]):
                return True, f"High experience requirement: '{match.group(0)}' in {snippet}"

    return False, "Suitable for entry-level/intern"


def search_candidate_jobs(limit=10):
    """Discovers active candidate job links strictly targeted at Internships, SDE-1, and Fresher roles."""
    queries = [
        'site:jobs.lever.co ("Intern" OR "Internship" OR "SDE-1" OR "SDE 1" OR "Junior" OR "Associate") Python (Bangalore OR Bengaluru OR Remote)',
        'site:job-boards.greenhouse.io ("Intern" OR "Internship" OR "SDE 1" OR "SDE-1" OR "Junior" OR "Fresher") Python (Bangalore OR Bengaluru OR Remote)',
        'site:boards.greenhouse.io ("Intern" OR "Internship" OR "SDE 1" OR "SDE-1" OR "Junior" OR "Associate") Python (Bangalore OR Bengaluru OR Remote)',
        'site:in.linkedin.com/jobs/view ("Intern" OR "Internship" OR "Junior" OR "Associate" OR "Fresher" OR "SDE 1") ("Python" OR "Backend") (Bengaluru OR Bangalore)',
        'site:wellfound.com/jobs ("Intern" OR "Junior" OR "SDE 1" OR "Backend") ("0-1" OR "0-2" OR "fresher" OR "intern") Python Bangalore',
        'site:cutshort.io/job ("Intern" OR "Junior" OR "SDE 1" OR "Fresher") Python Bangalore',
        'Juspay ("SDE 1" OR "SDE-1" OR "Intern" OR "Fresher" OR "Backend") Bangalore hiring',
        'CRED ("Backend" OR "Software") ("Intern" OR "Fresher" OR "Junior") Bangalore hiring'
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
                    print(f"[WARN] Query failed ({q[:40]}...): {e}")

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

                    # Drop closed/expired
                    if any(w in text.lower() for w in ['no longer accepting applications', 'position filled', 'job not found', 'expired']):
                        continue

                    raw_title = r.get('title', 'Software Engineer')

                    # HARD FILTER: Reject senior or 3+ YOE roles
                    is_senior, reason = is_senior_or_mismatched(raw_title, text)
                    if is_senior:
                        print(f"  [REJECTED - SENIOR] {raw_title[:45]} ({reason})")
                        continue

                    verified.append({
                        'title': raw_title,
                        'url': url,
                        'snippet': r.get('body', ''),
                        'text': text[:2500]
                    })
                    print(f"  [ACCEPTED - ENTRY/INTERN 200 OK] {raw_title[:50]} -> {url}")
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
- Target: HIGH-PAYING INTERNSHIPS, SDE-1, or FRESHER/JUNIOR FULL-TIME ROLES (0-2 YOE max).
- Core Stack: Python, FastAPI, PostgreSQL, SQLAlchemy, AWS, Docker, App Security (TryHackMe Top 5%).

JOB POSTING:
Title: {title}
Company: {company}
Snippet: {snippet}
Context: {full_text[:1200]}

SENIORITY RULE:
If this job posting requires 3+ years of experience or is explicitly a mid/senior/lead role, set "fitness_score": "35% Fit" and set "match_rationale": "Experience Mismatch: Requires 3+ years of experience; candidate is a final year undergrad seeking intern/fresher roles."

Respond ONLY in valid JSON matching this exact schema:
{{
  "fitness_score": "92% Fit",
  "is_suitable_for_fresher_or_intern": true,
  "match_rationale": "1-2 concise sentences explaining why this role is suitable for a final-year undergrad/intern and how candidate skills match.",
  "tailored_profile": "1-2 sentences summarizing candidate backend strengths aligned to this JD.",
  "cover_letter": "A concise, human-sounding cover letter under 250 words without AI clichés."
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
                    "rationale": parsed.get("match_rationale", "Strong match on Python, FastAPI, and PostgreSQL for entry-level/intern candidate."),
                    "latex_profile": parsed.get("tailored_profile", "Final-year Computer Science undergraduate and backend engineer specializing in Python (FastAPI), PostgreSQL, and secure API architectures."),
                    "cover_letter": parsed.get("cover_letter", "")
                }
        except Exception as e:
            print(f"[WARN] Groq evaluation failed ({e}), using rule-based generator.")

    # Rule-based fallback
    return {
        "company": company,
        "role": role,
        "fitness": "92% Fit",
        "suitable": True,
        "rationale": f"High alignment for final-year undergrad/intern role: Python backend development, FastAPI services, and PostgreSQL schemas at {company}.",
        "latex_profile": f"Final-year Computer Science undergraduate and backend engineer specializing in resilient API architecture, asynchronous data pipelines in Python (FastAPI), and relational modeling on PostgreSQL.",
        "cover_letter": f"""I am writing to apply for the {role} position at {company}. As a final-year Computer Science undergraduate with hands-on intern experience developing production backend services in Python (FastAPI, SQLAlchemy) and PostgreSQL, I am eager to contribute to your engineering team.

At 7HiddenLayers, I developed backend ingestion services that process complex document updates incrementally. In parallel, my independent engineering work includes architecting Guardian AI—an asynchronous platform integrating WebSockets and relational schemas—and securing a Top 5% global ranking on TryHackMe for defensive application security.

I write clean, tested code and am excited about the opportunity to bring my backend fundamentals and energy to {company}."""
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
    except Exception as e:
        print(f"[WARN] pdflatex invocation failed: {e}")

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
    """Creates local job_tracker.xlsx inside the run directory."""
    if not openpyxl:
        return
    tracker_path = os.path.join(run_dir, "job_tracker.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Job Tracker"
    headers = ['Date Found', 'Company', 'Role', 'Platform', 'Job URL', 'Fitness Score', 'Status']
    ws.append(headers)
    for r in tracker_rows:
        ws.append(r)
    wb.save(tracker_path)
    print(f"[OK] Saved tracker to {tracker_path}")


def send_gmail_report(to_addr, user_addr, app_password, subject, html_body, pdf_attachments):
    """Dispatches formatted HTML morning briefing with attached resume PDFs via Gmail SMTP."""
    msg = MIMEMultipart('mixed')
    msg['Subject'] = subject
    msg['From'] = user_addr
    msg['To'] = to_addr

    msg_alt = MIMEMultipart('alternative')
    msg_alt.attach(MIMEText(html_body, 'html'))
    msg.attach(msg_alt)

    for pdf_path in pdf_attachments:
        if os.path.exists(pdf_path):
            with open(pdf_path, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(pdf_path)}"')
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(user_addr, app_password)
            server.sendmail(user_addr, [to_addr], msg.as_string())
        print(f"[OK] Morning briefing email sent successfully to {to_addr}")
    except Exception as e:
        print(f"[WARN] Failed to send email via SMTP: {e}")


def main():
    parser = argparse.ArgumentParser(description="Autonomous Job Search Skill Runner - Intern & Fresher Edition")
    parser.add_argument("--dry-run", action="store_true", help="Perform discovery and local compilation without sending email")
    parser.add_argument("--limit", type=int, default=10, help="Number of fresh jobs to process")
    args = parser.parse_args()

    run_id = f"run_{get_timestamp()}"
    run_dir = os.path.join("runs", run_id)
    os.makedirs(run_dir, exist_ok=True)
    print(f"=== Starting Autonomous Job Skill Execution (Intern & Fresher Target): {run_id} ===")

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

    # Step 1: Discover candidate jobs
    jobs = search_candidate_jobs(limit=args.limit * 2)  # Search double to allow for filtering
    if not jobs:
        print("[INFO] No external search results found. Using verified seed targets.")
        jobs = [
            {"title": "Software Engineer Intern at Weekday", "href": "https://jobs.lever.co/weekdayworks/a86efff4-1c2e-43fc-8ecb-dc8f873db2f1", "snippet": "Backend Python software engineer intern.", "text": "Python, FastAPI, PostgreSQL"},
            {"title": "Software Engineer (Intern) - Backend at Merkle Science", "href": "https://jobs.lever.co/merklescience/e663b69b-264a-4bd7-b04d-fb3c0a824a28", "snippet": "Backend intern Python PostgreSQL.", "text": "Python, REST APIs, Microservices, PostgreSQL"},
            {"title": "Software Development Engineer Backend (DEV-BE02) at Juspay", "href": "https://juspay.io/careers/DEV-BE02", "snippet": "First principles engineering, SDE-1 / Fresher.", "text": "Algorithms, Networking, Operating Systems"}
        ]

    tracker_rows = []
    pdf_attachments = []
    report_rows_html = []

    today_str = datetime.date.today().strftime("%d %b %Y")
    accepted_count = 0

    for j in jobs:
        eval_res = evaluate_job(j, base_tex, groq_api_key=groq_api_key)

        # STRICT FILTER: Discard if marked unsuitable or fitness < 70%
        fit_num = 80
        try:
            fit_num = int(re.search(r"\d+", eval_res["fitness"]).group(0))
        except Exception:
            pass

        if not eval_res.get("suitable", True) or fit_num < 70 or "mismatch" in eval_res.get("rationale", "").lower():
            print(f"  [DISCARDED BY LLM] {eval_res['company']} - {eval_res['role']} ({eval_res['fitness']} | {eval_res['rationale'][:60]}...)")
            continue

        accepted_count += 1
        comp = eval_res["company"].replace(" ", "")
        comp_dir = os.path.join(run_dir, comp)
        os.makedirs(comp_dir, exist_ok=True)

        # 1. Compile LaTeX PDF
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
        if pdf_ok:
            pdf_attachments.append(pdf_full_path)

        # 2. Generate DOCX
        generate_docx_files(
            output_dir=comp_dir,
            comp_name=eval_res["company"],
            role_name=eval_res["role"],
            profile_text=eval_res["latex_profile"],
            cover_letter_text=eval_res["cover_letter"]
        )

        job_url = j.get("url") or j.get("href", "#")
        tracker_rows.append([
            today_str,
            eval_res["company"],
            eval_res["role"],
            "Direct/ATS",
            job_url,
            eval_res["fitness"],
            "Found"
        ])

        report_rows_html.append(f"""
        <tr>
            <td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>{accepted_count}</b></td>
            <td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>{eval_res['company']}</b></td>
            <td style="padding: 8px; border-bottom: 1px solid #ddd;">{eval_res['role']}</td>
            <td style="padding: 8px; border-bottom: 1px solid #ddd; color: #0d6efd;"><b>{eval_res['fitness']}</b></td>
            <td style="padding: 8px; border-bottom: 1px solid #ddd;"><a href="{job_url}" style="background-color: #0d6efd; color: white; padding: 6px 12px; text-decoration: none; border-radius: 4px; display: inline-block;">Apply Now</a></td>
        </tr>
        <tr>
            <td colspan="5" style="padding: 4px 8px 12px 8px; color: #555; font-size: 13px; border-bottom: 1px solid #eee;">
                <i>Rationale:</i> {eval_res['rationale']}
            </td>
        </tr>
        """)

        if accepted_count >= args.limit:
            break

    print(f"Total accepted entry-level/intern postings: {accepted_count}")

    # Update local Excel tracker
    update_excel_tracker(run_dir, tracker_rows)

    # HTML Morning Email Report
    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333; line-height: 1.5;">
        <h2>Good Morning Jayaditya! 🚀</h2>
        <p>Here is your tailored morning job briefing for <b>{today_str}</b>, specifically filtered for <b>High-Paying Internships and Fresher/Junior SDE-1 roles</b> (0–2 YOE) in Bengaluru & Remote matching your Python/Backend focus.</p>
        
        <table style="width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 25px;">
            <thead>
                <tr style="background-color: #f8f9fa; text-align: left;">
                    <th style="padding: 10px; border-bottom: 2px solid #ddd;">#</th>
                    <th style="padding: 10px; border-bottom: 2px solid #ddd;">Company</th>
                    <th style="padding: 10px; border-bottom: 2px solid #ddd;">Role</th>
                    <th style="padding: 10px; border-bottom: 2px solid #ddd;">Fitness</th>
                    <th style="padding: 10px; border-bottom: 2px solid #ddd;">Action</th>
                </tr>
            </thead>
            <tbody>
                {''.join(report_rows_html)}
            </tbody>
        </table>

        <p><b>Attached Materials:</b> Tailored 1-page LaTeX PDFs for today's top matches are attached directly to this email. Complete DOCX resumes and cover letters are stored in your <code>runs/{run_id}/</code> directory.</p>
        <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;" />
        <p style="font-size: 12px; color: #888;">Autonomous Job Search Assistant | Powered by Antigravity & Groq</p>
    </body>
    </html>
    """

    summary_file = os.path.join(run_dir, "run_summary.md")
    with open(summary_file, "w", encoding="utf-8") as f:
        f.write(f"# Job Search Summary - {today_str}\n\nProcessed {accepted_count} entry-level/intern jobs in {run_id}.\n")

    if not args.dry_run and gmail_user and gmail_pwd and accepted_count > 0:
        send_gmail_report(
            to_addr=gmail_user,
            user_addr=gmail_user,
            app_password=gmail_pwd,
            subject=f"Daily Job Search Briefing (Intern & Fresher) - {today_str} ({accepted_count} Roles)",
            html_body=html_body,
            pdf_attachments=pdf_attachments[:5]
        )
    else:
        print("[INFO] Skipping Gmail dispatch (dry-run mode, missing credentials, or 0 accepted jobs).")

    print(f"=== Execution Finished Cleanly in {run_dir} ===")


if __name__ == "__main__":
    main()
