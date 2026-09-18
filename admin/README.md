# CanvExporter for Admins

The full tool. A Python application that pulls complete Canvas course content, including uploaded files, into an organized folder on your computer. Built for records requests, departmental archiving, tenure packets, and any job where you need a whole course out of Canvas in readable form and at scale.

Runs with a Canvas API token, so it can reach everything the token's owner can reach. When you launch it, it opens a simple form in your browser. No command-line wrangling past the one launch command.

---

## What it pulls

- Assignments, pages, announcements, and discussions as `.docx`
- Quiz questions and answer keys as `.docx` and `.csv`
- Rubrics formatted for reuse
- Uploaded course files
- All organized into named subfolders

Output is plain `.docx` and `.csv`. No `.imscc`, no XML, no importing back into an LMS to read it.

---

## What you need

1. **Python 3** installed on your computer. If you've never installed it: on macOS use [Homebrew](https://brew.sh) (`brew install python`), on Windows 11 use `winget install python3`. Full step-by-step install guides: **[macOS](install-macos.md)** · **[Windows 11](install-windows.md)**.
2. **A Canvas API access token.** Generated in Canvas under Account → Settings → Approved Integrations, or issued by your Canvas administrator. Treat it like a password.
3. **The Course ID** of the course you want to export (the number in the course URL).

---

## Setup

Open a terminal, go to the folder containing `canvexporter.py`, and set up an isolated environment for its libraries:

```bash
python3 -m venv .venv
source .venv/bin/activate          # macOS/Linux
# .venv\Scripts\activate           # Windows

pip install -r requirements.txt
```

That installs the five libraries the tool needs: `requests`, `beautifulsoup4`, `python-docx`, `canvasapi`, and `flask`.

> **Package vs. import names:** you install `beautifulsoup4` and `python-docx`, but the code imports them as `bs4` and `docx`. That mismatch is normal. `requirements.txt` handles it for you.

---

## Running it

```bash
python3 canvexporter.py
```

The tool starts a small web server on your own machine and **opens your browser automatically** to a form. You'll enter three things:

- **Institution** — your Canvas subdomain (the part before `.instructure.com`)
- **Access Token** — your Canvas API token
- **Course ID** — the number from the course URL

Click to run, and it writes the exported course into organized folders on your computer.

The server runs only on `127.0.0.1` (your machine, reachable by nobody else). When you're done, return to the terminal and press `Ctrl+C` to stop it.

---

## Troubleshooting

**It won't start, or says the port is in use.**
The tool uses port 5000. On macOS, AirPlay Receiver sometimes occupies that port. Turn it off under System Settings → General → AirDrop & Handoff → AirPlay Receiver, then try again.

**`ModuleNotFoundError` when I run it.**
Your virtual environment isn't active or the libraries aren't installed. Run `source .venv/bin/activate` (macOS/Linux) or the Windows equivalent, then `pip install -r requirements.txt`.

**The export is empty or errors out.**
Usually a wrong token or Course ID, or a token that doesn't have access to that course. Confirm the token's owner can open the course in Canvas.

**Still stuck?**
Open an [Issue](https://github.com/techconsigliere/CanvExporter/issues).

---

## A note on scope

This tool acts entirely with the permissions of the token you give it, through the public Canvas API, exactly as designed. It makes authenticated requests, touches only content that token can already reach, respects rate limits, and stores nothing anywhere but your own computer. Reviewed against Instructure's Acceptable Use Policy and Canvas API Policy before release.

---

*Part of [CanvExporter](https://github.com/techconsigliere/CanvExporter) by Chris Powell, Canvas LMS Administrator at Western Washington University. Written up at [Canvas Insider](https://canvasinsider.blog).*
