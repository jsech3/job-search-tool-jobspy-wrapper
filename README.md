# Job Search Tool

Search the major job boards for the roles you want and get back a **clean, ready-to-read
Excel file** — job title, company, salary, location, and a clickable apply link for every
posting, all in one place.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Powered by JobSpy](https://img.shields.io/badge/powered%20by-JobSpy-1B6B4B)

> **Not a coder?** Open **`START-HERE.html`** in your browser for a friendly, visual
> step-by-step walkthrough.

---

## Features

- 🔎 Search **multiple job titles** in a single run
- 🌐 Pulls from **Indeed, LinkedIn, ZipRecruiter, and Glassdoor**
- 🧹 **Removes duplicate** postings automatically
- 💵 Clean salary formatting — `$80,000 – $120,000/yr`
- 📊 Finished **Excel**: frozen header, filter dropdowns, sensible widths, clickable links
- ⚙️ One plain-English **settings block** — no coding required

---

## Quick start

```bash
# 1. Install dependencies (one time)
pip install -r requirements.txt

# 2. Open job_search.py and edit the SETTINGS block at the top
#    (job titles, location, how far back to search)

# 3. Run it
python job_search.py
```

Results are written to a date-stamped file like `Job_Postings_2026-07-02_1430.xlsx` in
this folder. Each run creates a new file, so earlier searches are never overwritten.

> Requires **Python 3.10+**. Get it at [python.org/downloads](https://www.python.org/downloads/)
> (on Windows, tick **“Add Python to PATH”** during install).

---

## Settings

All configuration lives in one block at the top of `job_search.py`:

| Setting | What it does | Example |
|---|---|---|
| `JOB_TITLES` | The roles to search for | `["Sales Representative", "Account Executive"]` |
| `LOCATION` | Where to search | `"San Francisco, CA"` or `"Remote"` |
| `HOURS_OLD` | How far back, in hours (days × 24) | `72` (last 3 days) |
| `RESULTS_PER_SITE` | Results per board, per title | `50` |
| `EXCLUDE_TITLE_KEYWORDS` | Hide titles containing these words | `["Senior", "Director"]` |
| `SITES` | Which boards to search | `["indeed", "linkedin"]` |

You never need to edit anything below the settings block.

---

## Output columns

| Column | Description |
|---|---|
| Job Title | The posting's title |
| Company | Who's hiring |
| Location | City/state, or “Remote” |
| Salary | Pay range when published, otherwise “Not listed” |
| Job Type | Full-time, contract, etc. |
| Remote? | Yes / No |
| Date Posted | Sorted newest-first |
| Source | Which board it came from |
| Link | Clickable link to the posting |
| Description | Full text when available, else “Click link to find description” |

---

## Good to know

- **Runs take a minute or two** — it searches every title across the boards, de-dupes, then writes the file.
- **Fewer results than expected?** Boards sometimes rate-limit automated searches. Indeed and LinkedIn are the most reliable; ZipRecruiter and Glassdoor can come up empty on a given run. Wait a bit and retry.
- **“Not listed” salaries are normal** — many employers don't publish pay.

---

## Responsible use

This tool is intended for **personal, educational job searching**. It uses JobSpy to read
publicly listed job postings — please respect each job board's Terms of Service, keep your
usage reasonable, and don't overload the sites. You are responsible for how you use it.

---

## Credits

Powered by **[JobSpy](https://github.com/speedyapply/JobSpy)** (`python-jobspy`), an
open-source library released under the MIT License. JobSpy does the job-board searching;
this project wraps it in a simple settings file and a clean, formatted Excel output.

## License

Released under the [MIT License](LICENSE). If you fork or share this project, please keep
the JobSpy credit above.
