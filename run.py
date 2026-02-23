"""
run.py — Master orchestrator
─────────────────────────────
Single command to: tailor resume → log to Excel → auto-apply

Usage:
    python run.py \
        --company "Stripe" \
        --title "Backend Engineer" \
        --url "https://stripe.com/jobs/listing/backend-engineer/12345" \
        --platform "generic" \
        --job-desc job.txt \
        --apply   # omit to only tailor + log (no auto-apply)

Environment variables (or edit config.json):
    GROQ_API_KEY   your Groq key
"""

import argparse
import os
import subprocess
import sys
import json
from pathlib import Path

# ── Paths (edit these to match your setup) ────────────────────────────────────
BASE_DIR    = Path(__file__).parent
SCRIPTS_DIR = BASE_DIR / "scripts"
TRACKER     = BASE_DIR / "Job_Application_Tracker.xlsx"
RESUMES_DIR = BASE_DIR / "resumes"
CONFIG_FILE = BASE_DIR / "config.json"
BASE_RESUME = BASE_DIR / "base_resume.docx"

def run_step(cmd: list, label: str) -> bool:
    print(f"\n{'─'*50}")
    print(f"▶  {label}")
    print(f"{'─'*50}")
    result = subprocess.run(cmd, capture_output=False, text=True)
    if result.returncode != 0:
        print(f"❌ {label} failed (exit {result.returncode})")
        return False
    return True


def load_config() -> dict:
    if CONFIG_FILE.exists():
        return json.loads(CONFIG_FILE.read_text())
    return {}


def main():
    parser = argparse.ArgumentParser(description="Job Application Orchestrator")
    parser.add_argument("--company",  required=True)
    parser.add_argument("--title",    required=True)
    parser.add_argument("--url",      default="")
    parser.add_argument("--platform", default="generic",
                        choices=["linkedin", "indeed", "glassdoor", "generic"])
    parser.add_argument("--job-desc", required=True,
                        help="Path to job description .txt or paste text")
    parser.add_argument("--location", default="")
    parser.add_argument("--salary",   default="")
    parser.add_argument("--apply",    action="store_true",
                        help="Also run auto-apply bot after tailoring")
    parser.add_argument("--resume",   default=str(BASE_RESUME),
                        help="Path to base resume .docx")
    parser.add_argument("--headless", action="store_true")
    args = parser.parse_args()

    cfg     = load_config()
    api_key = os.getenv("GROQ_API_KEY") or cfg.get("groq_api_key", "")
    name    = cfg.get("full_name", "Your Name")

    if not api_key:
        print("❌ Set GROQ_API_KEY environment variable or add to config.json")
        sys.exit(1)

    # ── Step 1: Tailor resume + log to tracker ───────────────────────────────
    tailor_cmd = [
        sys.executable, str(SCRIPTS_DIR / "resume_tailor.py"),
        "--resume",   args.resume,
        "--job-desc", args.job_desc,
        "--company",  args.company,
        "--title",    args.title,
        "--platform", args.platform,
        "--url",      args.url,
        "--location", args.location,
        "--salary",   args.salary,
        "--tracker",  str(TRACKER),
        "--api-key",  api_key,
        "--name",     name,
        "--out-dir",  str(RESUMES_DIR),
    ]

    if not run_step(tailor_cmd, "Tailoring resume & logging to Excel"):
        sys.exit(1)

    # ── Step 2: Find the newly created resume ────────────────────────────────
    import re, datetime
    safe_company  = re.sub(r"[^\w]", "_", args.company)
    safe_title    = re.sub(r"[^\w]", "_", args.title)
    date_str      = datetime.date.today().strftime("%Y%m%d")
    resume_file   = RESUMES_DIR / f"Resume_{safe_company}_{safe_title}_{date_str}.docx"

    if not resume_file.exists():
        # fallback: pick newest docx in resumes dir
        docs = sorted(RESUMES_DIR.glob("*.docx"), key=os.path.getmtime, reverse=True)
        if docs:
            resume_file = docs[0]
        else:
            print("❌ Could not find tailored resume file")
            sys.exit(1)

    print(f"\n📄 Tailored resume: {resume_file}")

    # ── Step 3 (optional): Auto-apply ────────────────────────────────────────
    if args.apply:
        if not args.url:
            print("⚠️  --url required for auto-apply. Skipping.")
        else:
            apply_cmd = [
                sys.executable, str(SCRIPTS_DIR / "auto_apply.py"),
                "--url",      args.url,
                "--platform", args.platform,
                "--resume",   str(resume_file),
                "--tracker",  str(TRACKER),
                "--config",   str(CONFIG_FILE),
            ]
            if args.headless:
                apply_cmd.append("--headless")

            run_step(apply_cmd, f"Auto-applying on {args.platform.title()}")
    else:
        print("\n💡 Tip: Add --apply to also auto-submit the application.")
        print(f"   Or manually upload: {resume_file}")

    print(f"\n✅ All done! Check your tracker: {TRACKER}")


if __name__ == "__main__":
    main()
