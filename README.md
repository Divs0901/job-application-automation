# 🚀 Job Application Automation System

Automates the full job application pipeline:
1. **Tailor** your resume to each job using Claude AI
2. **Log** every application to Excel (with resume hyperlink)
3. **Auto-apply** on LinkedIn, Indeed, Glassdoor, and company pages

---

## 📁 Project Structure

```
job_automation/
├── run.py                          ← Master command (start here)
├── config.json                     ← Your credentials & personal info
├── base_resume.docx                ← Your base resume (YOU add this)
├── Job_Application_Tracker.xlsx    ← Auto-generated Excel tracker
├── resumes/                        ← Tailored resumes saved here
└── scripts/
    ├── create_tracker.py           ← Builds the Excel tracker
    ├── resume_tailor.py            ← Claude AI resume tailoring engine
    └── auto_apply.py               ← Selenium browser automation bot
```

---

## ⚙️ Setup (One-Time)

### 1. Install dependencies
```bash
pip install anthropic python-docx openpyxl selenium webdriver-manager
```

### 2. Add your base resume
Copy your resume as `base_resume.docx` into the `job_automation/` folder.

### 3. Fill in config.json
Edit `config.json` with your:
- **Anthropic API key** → get one at https://console.anthropic.com
- **Full name, email, phone, location**
- **LinkedIn / Indeed / Glassdoor credentials**

### 4. Create the Excel tracker
```bash
python scripts/create_tracker.py
```

### 5. Install ChromeDriver
The bot uses Chrome. Install ChromeDriver matching your Chrome version:
- Auto: `pip install webdriver-manager` (handles it automatically)
- Manual: https://chromedriver.chromium.org/downloads

---

## 🎯 Usage

### Option A — Tailor + Log only (you apply manually)
```bash
python run.py \
  --company "Stripe" \
  --title "Backend Engineer" \
  --url "https://stripe.com/jobs/..." \
  --platform "generic" \
  --job-desc job_description.txt
```

### Option B — Tailor + Log + Auto-Apply
```bash
python run.py \
  --company "Google" \
  --title "Software Engineer" \
  --url "https://www.linkedin.com/jobs/view/123456" \
  --platform "linkedin" \
  --job-desc job.txt \
  --apply
```

### Platform options
| Value       | What it does                          |
|-------------|---------------------------------------|
| `linkedin`  | LinkedIn Easy Apply                   |
| `indeed`    | Indeed Quick Apply                    |
| `glassdoor` | Glassdoor Easy Apply                  |
| `generic`   | Any company career page (Greenhouse, Lever, Workday) |

### Job description input
Pass as a `.txt` file path OR inline text:
```bash
--job-desc "We are looking for a Python engineer with 5+ years..."
# or
--job-desc ./jobs/stripe_backend.txt
```

---

## 📊 Excel Tracker

The tracker has 3 sheets:

| Sheet            | Purpose                                       |
|------------------|-----------------------------------------------|
| **Applications** | Every job row — status, resume link, notes    |
| **Dashboard**    | Live stats: total applied, interviews, offers |
| **Config**       | Settings reference (edit config.json instead) |

### Status color coding
| Color  | Status                          |
|--------|---------------------------------|
| 🟡 Yellow | Applied / Phone Screen / Test |
| 🔵 Blue   | Interview / Final Round       |
| 🟢 Green  | Offer                         |
| 🔴 Red    | Rejected / Withdrawn          |

---

## 💡 Tips

- **Review before submitting**: Run without `--apply` first to check the tailored resume
- **LinkedIn works best**: Easy Apply has the most consistent form structure
- **Glassdoor & Indeed**: May require 2FA — handle manually once, then cookies persist
- **Company pages**: The generic bot works for Greenhouse and Lever; Workday is more complex
- **Rate limiting**: Add `--headless` flag for background operation

---

## ⚠️ Important Notes

- The auto-apply bot uses browser automation — this is against some platforms' ToS
- Always review tailored resumes before submitting — AI can occasionally hallucinate
- Never include false information in your applications
- Store `config.json` securely — it contains your passwords

---

## 🔧 Advanced: Run resume tailoring standalone
```bash
python scripts/resume_tailor.py \
  --resume base_resume.docx \
  --job-desc job.txt \
  --company "Meta" \
  --title "ML Engineer" \
  --tracker Job_Application_Tracker.xlsx \
  --api-key sk-ant-...
```

## 🔧 Advanced: Update application status manually
```bash
python -c "
from scripts.auto_apply import update_tracker_status
update_tracker_status('Job_Application_Tracker.xlsx', 'https://linkedin.com/...', 'Interview')
"
```
