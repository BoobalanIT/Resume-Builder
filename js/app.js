/* ═══════════════════════════════════════════════
   Resume Builder — app.js
   Edit Mode · Theme · PDF Export · Word Export
═══════════════════════════════════════════════ */

let isEditMode = false;

// ── Toggle Edit Mode ──────────────────────────
function toggleEdit() {
  isEditMode = !isEditMode;
  const body  = document.body;
  const panel = document.getElementById("editPanel");
  const btn   = document.getElementById("toggleEditBtn");

  document.querySelectorAll("[contenteditable]").forEach(el => {
    el.contentEditable = isEditMode ? "true" : "false";
  });

  if (isEditMode) {
    body.classList.add("edit-mode", "edit-open");
    panel.classList.add("open");
    btn.textContent = "✅ Done Editing";
    showToast("Edit mode ON — click any text to change it");
  } else {
    body.classList.remove("edit-mode", "edit-open");
    panel.classList.remove("open");
    btn.textContent = "✏️ Edit Mode";
    showToast("Changes saved to view");
  }
}

// ── Sync Quick Panel → Resume ─────────────────
function syncText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

// ── Theme Switcher ────────────────────────────
function applyTheme(theme) {
  document.body.className = document.body.className.replace(/theme-\S+/g, "").trim();
  if (theme && theme !== "classic") document.body.classList.add("theme-" + theme);
  if (isEditMode) document.body.classList.add("edit-mode", "edit-open");
  showToast("Theme: " + theme.charAt(0).toUpperCase() + theme.slice(1));
}

// ── PDF Export — uses browser print (exact HTML match, vector PDF) ──
function exportPDF() {
  const wasEdit = isEditMode;
  if (wasEdit) toggleEdit();
  window.print();
  if (wasEdit) setTimeout(() => toggleEdit(), 500);
}

// ── Word Export ───────────────────────────────
function exportWord() {
  if (typeof docx === "undefined") {
    showToast("⚠️ Word library not loaded — check your internet connection.");
    return;
  }

  showToast("Building Word document…");

  try {
    const { Document, Packer, Paragraph, TextRun, AlignmentType } = docx;

    const t  = id => { const el = document.getElementById(id); return el ? el.innerText.trim() : ""; };
    const li = sel => [...document.querySelectorAll(sel)].map(e => e.innerText.trim());

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 720, bottom: 720, left: 1000, right: 1000 }
          }
        },
        children: [
          // Header
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t("name"), bold: true, size: 46, color: "1B365D" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t("jobTitle"), bold: true, size: 22, color: "1E293B" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t("location") + "  |  " + t("phone") + "  |  " + t("email"), size: 18, color: "475569" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "LinkedIn: Boobalan V  |  GitHub: boobalanIT  |  Portfolio: alanz.netlify.app", size: 18, color: "475569" })] }),
          blank(),

          // Professional Summary
          heading("Professional Summary"),
          new Paragraph({ children: [new TextRun({ text: t("summary"), size: 19 })] }),
          blank(),

          // Technical Skills
          heading("Technical Skills"),
          bullet("Frontend: Angular, TypeScript, JavaScript, HTML5, CSS3, Bootstrap"),
          bullet("Backend: Node.js, NestJS, PHP, REST APIs"),
          bullet("Databases: PostgreSQL, MongoDB"),
          bullet("Web Servers & Deployment: Nginx, Apache, DigitalOcean (Ubuntu), Application Deployment"),
          bullet("Tools & CI/CD: Git, GitLab CI/CD, Postman, pgAdmin"),
          bullet("AI & Integration: Open-source LLM integration, Model Context Protocol (MCP), AI-powered chatbot integration"),
          blank(),

          // Professional Experience
          heading("Professional Experience"),
          new Paragraph({ children: [new TextRun({ text: "Megam Technologies", bold: true, size: 21, color: "111827" }), new TextRun({ text: "   December 2024 – Present", bold: true, size: 19, color: "1E293B" })] }),
          new Paragraph({ children: [new TextRun({ text: "Software Developer | Chennai, India", italics: true, size: 19, color: "475569" })] }),
          new Paragraph({ children: [new TextRun({ text: "Datalytics: AI-Powered Business Intelligence Platform", bold: true, size: 19, color: "1B365D" })] }),
          bullet("Integrated an open-source LLM via Model Context Protocol (MCP), enabling users to query business data in natural language and receive AI-generated insights."),
          bullet("Built interactive dashboards and reports to track key business metrics, including employee work hours, project progress, and issue tracking."),
          bullet("Created responsive Angular components and connected them to REST APIs built with Node.js and NestJS."),
          bullet("Configured Nginx as a reverse proxy to route frontend and backend API traffic, and set up GitLab CI/CD pipelines for automated deployment."),
          new Paragraph({ children: [new TextRun({ text: "Datalytics: UI Component Library & Embed SDK", bold: true, size: 19, color: "1B365D" })] }),
          bullet("Designed a reusable Angular component library with interactive charts, KPI widgets, and data tables supporting sorting, filtering, pagination, and report configuration."),
          bullet("Engineered an Embed SDK that lets third-party applications embed live analytics dashboards using token-based authentication."),
          bullet("Published versioned packages to a private GitLab npm registry, so other applications can install them through .npmrc configuration."),
          new Paragraph({ children: [new TextRun({ text: "LeatherLink TanPro 11.0: Tannery Management System (TMS)", bold: true, size: 19, color: "1B365D" })] }),
          bullet("Migrated PHP applications to a newly provisioned Ubuntu server on DigitalOcean, setting up the complete hosting environment and handling deployment, maintenance, bug fixes, and feature enhancements."),
          bullet("Maintained core PHP and PostgreSQL modules that support tannery operations across 300+ database tables."),
          bullet("Architected a PostgreSQL schema-per-tenant model to isolate data for each enterprise client."),
          bullet("Implemented MongoDB-based audit logging and multi-level approval workflows to track business transactions."),
          bullet("Optimized SQL queries, views, materialized views, functions, and triggers to support application workflows."),
          bullet("Automated PDF generation with mPDF for delivery challans, GST invoices, and job cards."),
          bullet("Configured Apache and PHP-FPM environments for multiple client applications, covering server setup, deployment, and environment configuration."),
          bullet("Scheduled cron jobs to automate recurring application tasks and background processes."),
          blank(),

          // Personal Projects
          heading("Personal Projects"),
          new Paragraph({ children: [new TextRun({ text: "KYAAT — Gesture-Controlled Mouse", bold: true, size: 20, color: "111827" }), new TextRun({ text: "  |  Python, OpenCV", italics: true, size: 18, color: "475569" })] }),
          new Paragraph({ children: [new TextRun({ text: "Developed a computer-vision tool using Python and OpenCV to recognize real-time hand movements and translate gestures into mouse cursor interactions with sub-20ms response latency.", size: 19 })] }),
          blank(),
          new Paragraph({ children: [new TextRun({ text: "Personal Portfolio", bold: true, size: 20, color: "111827" }), new TextRun({ text: "  |  Web Development", italics: true, size: 18, color: "475569" })] }),
          new Paragraph({ children: [new TextRun({ text: "Designed and built a fully responsive personal portfolio website to showcase technical skills, live project demonstrations, and professional experience. Available at: alanz.netlify.app", size: 19 })] }),
          blank(),

          // Education
          heading("Education"),
          new Paragraph({ children: [new TextRun({ text: "Bachelor of Technology in Information Technology — Anna University", bold: true, size: 21, color: "111827" }), new TextRun({ text: "   2020 – 2024", bold: true, size: 19, color: "1E293B" })] }),
          new Paragraph({ children: [new TextRun({ text: "Mohamed Sathak A J College of Engineering  |  Chennai, India  |  80% (First Class with Distinction)", size: 19, color: "475569" })] }),
          blank(),

          // Achievements & Certifications
          heading("Achievements & Certifications"),
          bullet("The Versatile Rockstar Award, Megam Technologies, 2026: certificate and trophy for excellent performance."),
          bullet("Full Stack Development: Infosys Springboard (Frontend Frameworks, Backend APIs, and Databases)"),
          bullet("Python Programming: Spoken Tutorial Project, IIT Bombay")
        ]
      }]
    });

    Packer.toBlob(doc).then(blob => {
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement("a");
      a.href     = url;
      a.download = "Boobalan_V_Resume.docx";
      a.click();
      URL.revokeObjectURL(url);
      showToast("✅ Word document saved!");
    });

  } catch (err) {
    showToast("❌ Word export error: " + err.message);
    console.error(err);
  }
}

// ── docx Helpers ──────────────────────────────
function blank()       { return new docx.Paragraph({ children: [] }); }
function heading(text) {
  return new docx.Paragraph({
    children: [new docx.TextRun({ text: text.toUpperCase(), bold: true, size: 20, color: "1B365D" })],
    border:   { bottom: { style: "single", size: 8, color: "1B365D", space: 2 } }
  });
}
function bullet(text) {
  return new docx.Paragraph({ bullet: { level: 0 }, children: [new docx.TextRun({ text, size: 19 })] });
}

// ── Toast ─────────────────────────────────────
function showToast(msg, ms = 3000) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(t._t);
  t._t = setTimeout(() => t.classList.remove("show"), ms);
}
