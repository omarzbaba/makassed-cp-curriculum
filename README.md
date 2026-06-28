# Makassed General Hospital — Clinical Pathology Residency Lecture Library

A personalized web dashboard for **Dr. Tamima El-Jisr**, Chairperson, Department of Laboratory Medicine, to browse and download the full Clinical Pathology residency lecture curriculum.

- **34 lectures** across **7 subspecialty blocks** (Hematology & Coagulation, Transfusion Medicine, Clinical Chemistry, Microbiology, Immunology & Serology, Molecular & Cytogenetics, Lab Management & Quality)
- Color-coded topic tabs, search, single-download or multi-select **.zip** download
- Editable PowerPoint decks in `decks/`; curriculum master (Word + PDF) in `curriculum/`

## How it works
`index.html` is a self-contained static site (no build step). It reads the embedded manifest and links to the `.pptx` files. It runs as-is on **GitHub Pages**.

## Maintaining / updating a lecture
1. Replace the relevant file in `decks/` (keep the same filename), or add a new one.
2. If you add/rename/remove a deck, update `manifest.json` (the list the dashboard reads).
3. Commit and push — GitHub Pages redeploys automatically.

The owner controls all content; the shared Pages link always reflects the latest pushed version.
