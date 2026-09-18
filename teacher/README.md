# CanvExporter for Teachers

Pull your Canvas course content into readable files, right from your browser. No account, no access token, no software to install. You drag one button to your bookmarks bar, and from then on it's one click inside any of your courses.

> ⚠️ **Beta.** This version works well on Chrome and Edge and on machines you control. It's still being tested on locked-down, IT-managed computers, where browser policies can sometimes block bookmarks-bar tools. If it doesn't work on your machine, that's useful to know: please [tell me what happened](https://github.com/techconsigliere/CanvExporter/issues). You won't break anything by trying.

---

## What you get

Click the button inside a Canvas course and CanvExporter pulls your authored content into readable files:

- Pages, assignments, and discussions as `.docx` files
- Quiz questions and answer keys as `.docx` and `.csv`
- Rubrics you can read and reuse

Everything comes out named in plain English and organized by type. No XML, no importing back into Canvas to read it.

---

## The four steps

**1. Open the install page.** Open [`install-exporter.html`](install-exporter.html) in your browser. (Download it and double-click, or open it directly.) You'll see a blue button.

**2. Drag the button to your bookmarks bar.** Click and drag the blue **CanvExporter** button up onto your browser's bookmarks bar. That's the install. The whole tool lives inside that bookmark.

**3. Go to a Canvas course and click it.** Open any course you teach in Canvas. With the course on screen, click the **CanvExporter** bookmark you just saved.

**4. Pick a folder and let it run.** On Chrome or Edge, you'll be asked once to choose a folder (your Desktop is fine). CanvExporter writes your content there in organized subfolders. On Safari or Firefox, it hands you a single `.zip` instead — same files inside, just zipped.

That's it.

---

## What browser you're using matters

CanvExporter uses a newer browser feature to write files into a folder you choose. That feature only exists in **Chrome and Edge**.

- **Chrome or Edge:** you get a real folder of files, organized and ready.
- **Safari or Firefox:** you get a `.zip` file with the same content inside. Unzip it and everything's there, readable.

If you're on a Mac, your default browser is probably Safari, so you'll get the zip unless you run it in Chrome. Either way, your content comes out readable.

---

## What this does *not* do

It pulls the content you **authored** in Canvas. It does not download the **files you uploaded** (PDFs, slide decks, images). You don't need this tool for those, Canvas already does it: go to your course **Files**, select all, and click **Download**, and Canvas gives you a zip of everything.

---

## Want to see what's inside the button?

The bookmarklet is one long line of JavaScript, which is hard to read. The full, readable source is right here: [`exporter_src.js`](exporter_src.js). If you're the cautious type, or your IT department is, read it before you run it. It only reads content your Canvas login already lets you see, and it sends nothing anywhere. Everything happens in your browser, on your machine.

---

## It didn't work?

That's genuinely useful information while this is in beta. Open an [Issue](https://github.com/techconsigliere/CanvExporter/issues) and tell me: your browser, whether your computer is managed by your institution's IT, and what happened when you clicked. Every failure report makes it better.

---

*Part of [CanvExporter](https://github.com/techconsigliere/CanvExporter) by Chris Powell. Written up at [Canvas Insider](https://canvasinsider.blog).*
