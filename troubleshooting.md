# Troubleshooting memoQ Issues

This file covers common problems users might face when working with memoQ and related scripts or regex patterns in this repo.

## Table of contents

- [Scripts Fail to Run or Throw Errors](#scripts-fail-to-run-or-throw-errors)
- [Slow Performance on Large Projects](#slow-performanceon-large-projects)
- [Sync Conflicts in Team Projects](#sync-conflicts-in-team-projects)
- [Corrupt or Missing File (desktop client)](#corrupt-or-missing-file-(-desktop-client-))
- [No Local Copy (desktop client)](#no-local-copy-(-desktop-client-))

---

## Common Issues and Solutions

### Regex Patterns Not Matching Correctly
- **Problem:** Your regex does not find the expected matches in memoQ.
- **Solutions:**
  - Ensure you are using **Perl-Compatible Regular Expressions (PCRE)** syntax, which memoQ supports.
  - Double-check all escape characters (e.g., use `\\` to match a literal backslash).
  - Test your regex patterns with external tools like [regex101.com](https://regex101.com/).
  - Avoid greedy quantifiers if you want non-greedy matches (`.*?` instead of `.*`).

---

### Scripts Fail to Run or Throw Errors
- **Problem:** Automation scripts in this repo do not execute properly.
- **Solutions:**
  - Confirm you have the correct runtime environment (e.g., Python 3.x, PowerShell).
  - Check for required dependencies or modules.
  - Run scripts from the directory where they reside to avoid path issues.
  - Read error messages carefully and check line numbers for debugging.
  - For memoQ API scripts, ensure your memoQ server URL and credentials are correctly configured.

---

### Slow Performance on Large Projects
- **Problem:** memoQ slows down or freezes when loading large files.
- **Solutions:**
  - Split large projects into smaller chunks.
  - Archive or clean unused translation memories.
  - Close other heavy applications to free system resources.
  - Check QA settings.

---

### Sync Conflicts in Team Projects
- **Problem:** Conflicts arise when multiple translators work on the same project.
- **Solutions:**
  - Always sync before starting work and after finishing.
  - Communicate with your team to avoid overlapping edits.
  - Use memoQ’s **Check-in/Check-out** features properly.
  - Contact your memoQ server admin if conflicts persist.

---

### Corrupt or Missing File (desktop client)
- **Problem:** Some memoQ-related files are either corrupt or missing from the computer.
- **Solutions:**
  - Uninstall memoQ – every version for the affected users.
  - If you have more than one, please delete the whole C:\Program Files\memoQ folder.
  - Then install the memoQ version you'd like to use: https://www.memoq.com/downloads.

---

### No Local Copy (desktop client)
- **Problem:** memoQ opens the wrong project.
- **Solutions:**
  - Check out a new project in memoQ. To do this, please do the following steps:
    1. Click on the 'Check out from Server' option in memoQ
    2. Connect to your server and choose the project you would like to use and check it out with a different name (memoQ server: [...].memoqworld.com)
    3. If the project is already checked out, a new pop-up window will appear, where you need to select the 'Check out new copy'
       Since the new project is a “Local copy”, you can find it in “My Computer” and not on the server. It is visible only to the creator of the local copy.

---

## Getting Further Help

- Check the official [memoQ Help Center](https://help.memoq.com/) for detailed guides.
