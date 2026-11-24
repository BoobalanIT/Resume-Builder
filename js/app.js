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
          bullet("Web Servers & Deployment: Nginx, Apache, Application Deployment"),
          bullet("Tools & CI/CD: Git, GitLab CI/CD, Postman, pgAdmin"),
          bullet("AI & Integration: Open-source LLM integration, Model Context Protocol (MCP), AI-powered chatbot integration"),
          blank(),

          // Professional Experience
          heading("Professional Experience"),
          new Paragraph({ children: [new TextRun({ text: "Megam Technologies", bold: true, size: 21, color: "111827" }), new TextRun({ text: "   December 2024 – Present", bold: true, size: 19, color: "1E293B" })] }),
          new Paragraph({ children: [new TextRun({ text: "Software Developer | Chennai, India", italics: true, size: 19, color: "475569" })] }),
          new Paragraph({ children: [new TextRun({ text: "Datalytics — AI-Powered Business Intelligence Platform", bold: true, size: 19, color: "1B365D" })] }),
          bullet("Developed responsive frontend components using Angular and integrated REST APIs built with Node.js and NestJS."),
          bullet("Integrated an open-source Large Language Model (LLM) with Model Context Protocol (MCP), enabling users to query business data using natural language and receive AI-generated insights."),
          bullet("Developed interactive dashboards and reporting features to monitor key business metrics, including employee work hours, project progress, and issue tracking."),
          bullet("Configured Nginx as a reverse proxy to route frontend requests and backend API traffic, and managed GitLab CI/CD pipelines for application deployment."),
          new Paragraph({ children: [new TextRun({ text: "Datalytics — UI Component Library & Embed SDK", bold: true, size: 19, color: "1B365D" })] }),
          bullet("UI Component Library: Developed an independent, reusable Angular component library featuring interactive charts (bar, line, area, and pie), KPI widgets, data tables with sorting, filtering, and pagination, and report configuration settings. Maintained independent versioning and published the package to npm."),
          bullet("Embed SDK: Developed a reusable Embed SDK within the main application repository, with a dedicated package.json and independent npm publishing configuration. Enabled third-party applications to embed live analytics dashboards using token-based authentication."),
          bullet("Package Management & Integration: Managed independent versioning and npm releases for both packages, and integrated updated library versions into the main Datalytics SaaS platform."),
          new Paragraph({ children: [new TextRun({ text: "LeatherLink TanPro 11.0 — Tannery Management System (TMS)", bold: true, size: 19, color: "1B365D" })] }),
          bullet("Migrated PHP applications and servers, configured deployment environments, and performed application maintenance, bug fixes, and feature enhancements."),
          bullet("Developed and maintained core application modules using PHP and PostgreSQL, supporting tannery business operations across 300+ database tables."),
          bullet("Implemented PostgreSQL schema-per-tenant architecture to maintain separate database schemas for different enterprise clients."),
          bullet("Implemented and maintained MongoDB-based audit logging and multi-level approval workflows to track business transactions."),
          bullet("Developed and optimized SQL queries, views, materialized views, functions, and triggers to support application workflows and database operations."),
          bullet("Implemented automated PDF generation using mPDF for delivery challans, GST invoices, and job cards."),
          bullet("Configured and maintained Apache and PHP-FPM environments, including server setup, deployment, and environment configuration for multiple client applications."),
          bullet("Configured and maintained cron jobs to automate scheduled application tasks and background processes."),
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
          new Paragraph({ children: [new TextRun({ text: "Bachelor of Technology (B.Tech) — Information Technology", bold: true, size: 21, color: "111827" }), new TextRun({ text: "   2020 – 2024", bold: true, size: 19, color: "1E293B" })] }),
          new Paragraph({ children: [new TextRun({ text: "Mohamed Sathak A J College of Engineering  |  Chennai, India  |  80% (First Class)", size: 19, color: "475569" })] }),
          blank(),

          // Achievements & Certifications
          heading("Achievements & Certifications"),
          bullet("The Versatile Rockstar Award, Megam Technologies, 2026 — certificate and trophy for excellent performance."),
          bullet("Full Stack Development — Infosys Springboard (Frontend Frameworks, Backend APIs, & Databases)"),
          bullet("Python Programming — Spoken Tutorial Project at IIT Bombay")
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
