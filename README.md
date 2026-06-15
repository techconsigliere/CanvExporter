# CanvExporter

A free, open-source tool that exports Canvas LMS course content into 
editable, offline-accessible files — so teachers can access their own 
work whether Canvas is online or not.

Built by a Canvas Administrator in direct response to the May 2026 
Instructure security incident.

## What It Does

CanvExporter uses the Canvas API to extract everything in a course and 
deliver it as organized, editable files in clean subfolders:

- Announcements
- Assignments
- Discussions
- Quizzes (with answer key and print-ready versions)
- Content Pages
- Question Banks
- Rubrics (formatted for re-import)
- Files

Output formats: `.docx` and `.csv` — openable on any desktop computer 
without Canvas.

## What It Requires

- Python 3.x
- Three inputs: `API_URL`, `API_ACCESS_TOKEN`, and `COURSE_ID`
- A Canvas account with access to the course you want to export

## API Compliance

CanvExporter uses the public Canvas API exactly as designed. It makes 
authenticated requests only on behalf of the user running the tool, 
accesses only content that user owns or has authorized access to, 
respects Canvas API rate limits, and stores nothing server-side. 
Reviewed against Instructure's Acceptable Use Policy and Canvas API 
Policy — fully compliant.

## License

Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)

Free to use. Credit required. Not for commercial use without permission.

[Full license text](LICENSE)

## Author

Chris Powell — Canvas LMS Administrator, Western Washington University  
GitHub: [techconsigliere](https://github.com/techconsigliere)  
Blog: [Canvas Insider](https://canvasinsider.blog)

## Status

Active development. Initial release coming soon.
