# Scripts for memoQ Automation

This folder contains automation scripts designed to streamline tasks in memoQ, such as project setup, batch imports, and quality checks.

---

## Included Scripts

- **project_setup.py** — Automates creating new projects and assigning resources.
- **batch_import.ps1** — PowerShell script for batch importing files into memoQ projects.
- **qa_check.py** — Runs custom QA checks using regex patterns from this repo.
- **export_results.sh** — Shell script for exporting translation results for review.

---

## How to Use

1. Ensure you have the required runtime installed (Python 3.x, PowerShell, Bash).
2. Update any configuration variables in the scripts (e.g., file paths, memoQ server URLs, credentials).
3. Run scripts from the command line or terminal:
   ```bash
   python project_setup.py
