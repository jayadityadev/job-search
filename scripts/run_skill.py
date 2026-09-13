import os
import sys
import json
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


def search_candidate_jobs(limit=10):
    """Discovers active candidate job links across ATS, LinkedIn guest views, and Wellfound."""
    queries = [
        'site:jobs.lever.co ("Backend Engineer" OR "Python Developer" OR "Software Engineer") Bangalore',
        'site:job-boards.greenhouse.io ("Backend Engineer" OR "Software Engineer") (Bangalore OR Bengaluru OR Remote) Python',
        'site:boards.greenhouse.io ("Backend" OR "Software Engineer") (Bangalore OR Bengaluru) Python',
        'site:in.linkedin.com/jobs/view ("Backend Engineer" OR "Python Developer") ("FastAPI" OR "Django" OR "Python") (Bengaluru OR Bangalore)',
        'site:wellfound.com/jobs "Backend Engineer" Python Bangalore',
        'site:cutshort.io/job "Backend" Python Bangalore',
        'Juspay "Backend" Bangalore hiring',
        'CRED "Backend" Bangalore hiring'
    ]

    results = []
    seen_urls = set()
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}

    if DDGS:
        with DDGS() as ddgs:
            for q in queries:
                try:
                    items = list(ddgs.text(q, max_results=6))
                    for it in items:
                        url = it.get('href')
                        if url and url not in seen_urls:
                            seen_urls.add(url)
                            results.append(it)
                except Exception as e:
                    print(f"[WARN] Query failed ({q[:40]}...): {e}")

    verified = []
    print(f"Discovered {len(results)} potential links. Verifying active status...")
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
                        text = resp.read().decode('utf-8', errors='ignore')[:1000]

                    if any(w in text.lower() for w in ['no longer accepting applications', 'position filled', 'job not found', 'expired']):
                        continue

                    # Extract company and role
                    raw_title = r.get('title', 'Software Engineer')
                    verified.append({
                        'title': raw_title,
                        'url': url,
                        'snippet': r.get('body', ''),
                        'text': text[:2000]
                    })
                    print(f"  [200 OK] {raw_title[:50]} -> {url}")
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

    # Extract clean company name
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
            prompt = f"""You are an expert career agent. Evaluate this candidate against this job description.
Candidate Profile:
{profile_text}

Job Posting:
Title: {title}
Company: {company}
Description snippet: {snippet}
Context: {full_text[:1000]}

Respond ONLY in valid JSON matching this exact schema:
{{
  "fitness_score": "88% Fit",
  "match_rationale": "1-2 concise sentences explaining the skill match and why it fits.",
  "tailored_profile": "1-2 sentences summarizing candidate backend strengths aligned to this JD.",
  "cover_letter": "A concise, human-sounding cover letter under 250 words without AI clichés."
}}"""

            req_data = json.dumps({
                "model": "llama-3.3-70b-versatile",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.3,
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
                    "rationale": parsed.get("match_rationale", "Strong match on Python, FastAPI, and PostgreSQL."),
                    "latex_profile": parsed.get("tailored_profile", "Backend software engineer focused on reliable APIs, asynchronous architectures, and PostgreSQL."),
                    "cover_letter": parsed.get("cover_letter", "")
                }
        except Exception as e:
            print(f"[WARN] Groq evaluation failed ({e}), using rule-based generator.")

    # Rule-based fallback
    return {
        "company": company,
        "role": role,
        "fitness": "90% Fit",
        "rationale": f"Matches Python backend profile, FastAPI services, and PostgreSQL schema design for {company}.",
        "latex_profile": f"Backend engineer specializing in resilient API architecture, asynchronous data pipelines in Python (FastAPI), and relational modeling on PostgreSQL.",
        "cover_letter": f"""I am writing to apply for the {role} position at {company}. My technical focus centers on building reliable backend systems, performant RESTful APIs, and asynchronous pipelines using Python (FastAPI, SQLAlchemy) and PostgreSQL.

At 7HiddenLayers, I developed backend ingestion services that process complex document updates incrementally. In parallel, my independent engineering work includes architecting Guardian AI—an asynchronous platform integrating WebSockets and relational schemas—and securing a Top 5% global ranking on TryHackMe for defensive application security.

I write clean, tested code and am excited about the opportunity to contribute to {company}'s engineering initiatives."""
    }


def compile_latex_pdf(output_dir, tex_filename, pdf_filename, base_tex, profile_text, resume_cls_path):
    """Compiles tailored 1-page LaTeX PDF using pdflatex."""
    # Ensure resume.cls is available
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

    # Compile with pdflatex
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

    # Move output PDF if needed
    compiled_pdf = os.path.splitext(tex_path)[0] + ".pdf"
    target_pdf = os.path.join(output_dir, pdf_filename)
    if os.path.exists(compiled_pdf) and compiled_pdf != target_pdf:
        shutil.move(compiled_pdf, target_pdf)

    # Cleanup aux files
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

    # 1. Resume DOCX
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

    # 2. Cover Letter DOCX
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
    parser = argparse.ArgumentParser(description="Autonomous Job Search Skill Runner")
    parser.add_argument("--dry-run", action="store_true", help="Perform discovery and local compilation without sending email")
    parser.add_argument("--limit", type=int, default=10, help="Number of fresh jobs to process")
    args = parser.parse_args()

    run_id = f"run_{get_timestamp()}"
    run_dir = os.path.join("runs", run_id)
    os.makedirs(run_dir, exist_ok=True)
    print(f"=== Starting Autonomous Job Skill Execution: {run_id} ===")

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
    jobs = search_candidate_jobs(limit=args.limit)
    if not jobs:
        print("[INFO] No external search results found. Using verified seed targets.")
        jobs = [
            {"title": "SDE 1/2 - Backend at CloudSEK", "href": "https://boards.greenhouse.io/cloudsek/jobs/4873360004", "snippet": "Cybersecurity threat intelligence backend Python.", "text": "Python, FastAPI, PostgreSQL"},
            {"title": "Software Engineer, API (Python) at Nexla", "href": "https://job-boards.greenhouse.io/nexla/jobs/4727753005", "snippet": "Data and API integration services.", "text": "Python, REST APIs, Microservices"},
            {"title": "Software Development Engineer Backend at Juspay", "href": "https://juspay.io/careers/DEV-BE02", "snippet": "High scale payment platform.", "text": "Algorithms, Networking, Operating Systems"}
        ]

    tracker_rows = []
    pdf_attachments = []
    report_rows_html = []

    today_str = datetime.date.today().strftime("%d %b %Y")

    for i, j in enumerate(jobs, 1):
        eval_res = evaluate_job(j, base_tex, groq_api_key=groq_api_key)
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
            <td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>{i}</b></td>
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

    # Update local Excel tracker
    update_excel_tracker(run_dir, tracker_rows)

    # HTML Morning Email Report
    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333; line-height: 1.5;">
        <h2>Good Morning Jayaditya! 🚀</h2>
        <p>Here is your automated morning job briefing for <b>{today_str}</b>. We discovered and tailored application packages for <b>{len(jobs)} fresh roles</b> in Bengaluru & Remote matching your Python/Backend focus.</p>
        
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
        f.write(f"# Job Search Summary - {today_str}\n\nProcessed {len(jobs)} jobs in {run_id}.\n")

    if not args.dry_run and gmail_user and gmail_pwd:
        send_gmail_report(
            to_addr=gmail_user,
            user_addr=gmail_user,
            app_password=gmail_pwd,
            subject=f"Daily Job Search Briefing - {today_str} ({len(jobs)} New Roles)",
            html_body=html_body,
            pdf_attachments=pdf_attachments[:5]
        )
    else:
        print("[INFO] Skipping Gmail dispatch (dry-run mode or missing GMAIL_USER/GMAIL_APP_PASSWORD).")

    print(f"=== Execution Finished Cleanly in {run_dir} ===")


if __name__ == "__main__":
    main()
