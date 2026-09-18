# CanvExporter

Get your Canvas course content out of Canvas and into files you can actually open, read, and print. No importing back into an LMS, no wrangling XML.

Built by a Canvas administrator after the May 2026 Instructure breach, when a lot of us learned the hard way that a course export hands you an `.imscc` file you can't read without Canvas.

There are two versions. Pick the one that fits you.

---

## 👩‍🏫 I'm a teacher — I want my own course content

Use the **[Teacher version](teacher/)**. It's a button you drag to your browser's bookmarks bar. No install, no account, no access token. Click it inside any Canvas course and it pulls your pages, assignments, quiz questions and answer keys, discussions, and rubrics into readable `.docx` and `.csv` files.

**→ [Go to the Teacher version](teacher/)** *(currently in beta — see the note there)*

---

## 🛠️ I'm an admin — I want the full tool

Use the **[Admin version](admin/)**. It's a Python tool that runs with a Canvas API token and pulls complete course content, including uploaded files, into an organized folder on your computer. Built for records requests, departmental work, and any job where you need everything in one place at scale.

**→ [Go to the Admin version](admin/)**

---

## What's the difference?

| | Teacher version | Admin version |
|---|---|---|
| **Who runs it** | Any teacher, in their own course | An admin with an API token |
| **Install** | Drag a button to your bookmarks bar | Python + a few libraries |
| **Access token needed** | No | Yes |
| **What it pulls** | Content you authored (pages, assignments, quizzes, discussions, rubrics) | Everything, including uploaded files |
| **Output** | `.docx` and `.csv` | `.docx` and `.csv`, organized folders |
| **Status** | Beta | Released |

Both do the same core job: turn Canvas course content into readable documents. The teacher version trades completeness for zero-friction access. The admin version trades convenience for a token and pulls the whole course.

A note on files: for downloading a course's uploaded **files** in bulk (PDFs, slide decks, images), Canvas already has you covered. In the course Files list, select all and click Download, and Canvas hands you a zip. Neither version reinvents that. The teacher version focuses on the authored content Canvas won't give you in readable form; the admin version pulls files too because the API token makes it possible.

---

## License

Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0). Free to use. Credit required. Not for commercial use without permission. [Full license](LICENSE).

---

## Who made this

Chris Powell, Canvas LMS Administrator at Western Washington University. Written up from the admin side of the desk at [Canvas Insider](https://canvasinsider.blog).
GitHub: [techconsigliere](https://github.com/techconsigliere)
