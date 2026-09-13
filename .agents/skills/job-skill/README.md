# job-skill

AI-powered job search assistant for Indian professionals — built as a Claude Skill.

Searches 12+ Indian and global job platforms (Naukri, LinkedIn, Instahyre, Cutshort, Hirist,
Indeed India, Foundit, Shine, TimesJobs, Glassdoor, WeWorkRemotely, AngelList) plus direct
company career pages, generates ATS-optimized, human-sounding resumes and cover letters
tailored to each job posting, scores keyword match, tracks applications, and can automate
nightly searches with a morning report.

## Commands

- `/job-skill help` — usage guide
- `/job-skill search` — search all platforms for roles matching your profile
- `/job-skill automate` — set up a nightly automated search with a morning report
- `/job-skill status` — check application status (auto-detects rejections/interviews via Gmail if connected)

Natural language also works — e.g. "Find me React jobs in Bangalore", "Apply to this job: [URL]",
"What's the status of my applications?"

## Key features

- 13-channel search (12+ job boards + direct company career pages)
- Tailored resume + cover letter generated per job, written to sound human (no AI clichés,
  no templated bullet rhythm) rather than obviously AI-generated
- Internal ATS keyword optimization (target ≥70% match) — never shown as a raw score to the user
- Application tracker (`job_tracker.xlsx`) with status flow from Found → Applied → Offer/Rejected
- Gmail-based auto-detection of rejections, interview invites, and assessments
- **Auto-improvement on rejection patterns**: when Gmail shows repeated or same-day rejections,
  the skill diagnoses the likely cause (ATS filter vs. targeting vs. seniority mismatch) and
  automatically tightens keyword matching or search targeting going forward — and always tells
  you what changed and why
- Clean output only: no tool/JS/code-block narration, no "let me check X" pauses — just job
  listings, fitness scores, apply links, and status updates
- Weekly self-correction review of response/interview/ghosted rates

## Installation

Copy `SKILL.md` into your Claude Skills directory (e.g. `~/.claude/skills/job-skill/` or the
equivalent user-skills path for your Claude setup), or install via your organization's skill
catalog if available.

## Privacy

All data stays within the session. Optional Gmail integration only reads emails from companies
you've applied to — nothing else. The skill never submits applications, creates accounts, or
handles CAPTCHAs/OTPs/logins on your behalf; you review the morning report and apply manually.

## License

MIT — use and adapt freely.
