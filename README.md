# 📄 Resume Builder — Boobalan V

**Complete resume project with HTML preview, PDF, Word (.docx), Excel (.xlsx), and generation scripts.**

---

## 📁 Project Structure

```
Resume Builder/
│
├── index.html                    ← 🌐 Open in browser — live resume with edit & export
├── css/style.css                 ← Styles + 4 color themes
├── js/app.js                     ← Edit mode, PDF export, Word export (browser-side)
│
├── generate_resume_docx.js       ← 📝 Node.js script → generates Boobalan_V_Resume.docx
├── generate_resume.js            ← 📊 Node.js script → generates Boobalan_V_Resume.xlsx
├── serve.js                      ← 🖥️ Local server at http://localhost:3456
├── package.json                  ← npm deps: docx ^9.14.0, exceljs ^4.4.0
│
├── Boobalan_V_Resume.pdf         ← ✅ Already-exported PDF (1 page, A4)
├── Boobalan_V_Resume.docx        ← ✅ Already-exported Word document
├── Boobalan_V_Resume.xlsx        ← ✅ Already-exported Excel (3 sheets)
│
└── README.md                     ← This file
```

---

## 🚀 Option 1 — Just Open in Browser (No install needed)

1. Double-click **`index.html`** in Chrome or Edge
2. Click **✏️ Edit Mode** → click any text on the resume to change it
3. Use the **Quick Edit panel** (sidebar) for fast header updates
4. Switch **themes** from the dropdown
5. Click **⬇️ Export PDF** or **📝 Export Word** to download

> ⚠️ Needs internet for Google Fonts on first open.

---

## 🖥️ Option 2 — Run the Local Server (Full experience)

Serves `preview.html` + direct download links for PDF/DOCX/XLSX:

```bash
cd "P:\Projects\Resume Builder"
npm install
node serve.js
```

Then open: **http://localhost:3456**

---

## 📝 Option 3 — Re-generate Word (.docx)

Uses the **docx** library — produces a properly formatted A4 Word doc:

```bash
cd "P:\Projects\Resume Builder"
npm install
node generate_resume_docx.js
```

Output: `Boobalan_V_Resume.docx`

---

## 📊 Option 4 — Re-generate Excel (.xlsx)

Uses **exceljs** — produces 3 sheets: clean view, gridlines view, ATS table:

```bash
cd "P:\Projects\Resume Builder"
npm install
node generate_resume.js
```

Output: `Boobalan_V_Resume.xlsx`

---

## 👤 Resume Data — Boobalan V

| Field | Value |
|---|---|
| Name | Boobalan V. |
| Title | Software Developer \| Full-Stack Developer |
| Email | boobalanvaithiyanathan1@gmail.com |
| Phone | +91 9025006329 |
| Location | Mayiladuthurai, Tamil Nadu, India |
| LinkedIn | linkedin.com/in/boobalan001 |
| GitHub | github.com/boobalanIT |
| Portfolio | alanz.netlify.app |

### Experience
- **Megam Technologies** — Software Developer (Dec 2024 – Present)
  - Datalytics: AI-Powered BI platform, Embed SDK & Charts/Tables UI library, LLM + MCP integration, Angular/NestJS/MongoDB, Nginx
  - LeatherLink TanPro 11.0 (TMS) ERP: PHP + PostgreSQL (schema-per-tenant), Apache with PHP-FPM
  - Automated mPDF reporting, production batch tracking, MongoDB audit workflows

### Projects
- **Datalytics Embed SDK & UI Library** — Component library & SDK (TypeScript, Angular, npm)
- **KYAAT** — Gesture-controlled mouse (Python, OpenCV)
- **Personal Portfolio** — alanz.netlify.app

### Education
- **B.Tech IT** — Mohamed Sathak A J College of Engineering (2020–2024, 80% First Class)

### Awards
- 🏆 **Versatile Rockstar Award** — Megam Technologies, 2026
- Full Stack Development — Infosys Springboard
- Python — Spoken Tutorial / CodeTantra

---

## 🛠️ Dependencies

| Package | Version | Use |
|---|---|---|
| docx | ^9.14.0 | Word .docx generation |
| exceljs | ^4.4.0 | Excel .xlsx generation |
| html2pdf.js | CDN | Browser PDF export |

> Original scripts from: `~/.gemini/antigravity-ide/brain/ebfbd316-1a9b-468e-90d4-8919e1ba8c63`
