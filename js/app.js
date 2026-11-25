/* ═══════════════════════════════════════════════
   Resume Builder — app.js
   Edit Mode · Theme · PDF Export · Word Export
═══════════════════════════════════════════════ */

let isEditMode = false;

// ── Toggle Edit Mode ──────────────────────────
function toggleEdit() {
  isEditMode = !isEditMode;
  const body = document.body;
  const panel = document.getElementById("editPanel");
  const btn = document.getElementById("toggleEditBtn");

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

    const t = id => { const el = document.getElementById(id); return el ? el.innerText.trim() : ""; };
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
          new Paragraph({ children: [new TextRun({ text: "Software Developer with nearly 2 years of experience building and maintaining enterprise applications and SaaS solutions. Hands-on experience across frontend development, backend API development, database management, and AI-powered integrations using Angular, Node.js, NestJS, PHP, PostgreSQL, and MongoDB. Contributed to enterprise ERP systems, reusable UI component libraries, embedded analytics solutions, and AI-driven reporting features. Familiar with application deployment, server configuration, and CI/CD workflows.", size: 19 })] }),
          blank(),

          // Technical Skills
          heading("Technical Skills"),
          bullet("Frontend: Angular, TypeScript, JavaScript, HTML5, CSS3, Bootstrap"),
          bullet("Backend: Node.js, NestJS, Express.js, PHP, REST API Development, Token-Based Authentication, TOTP-Based Two-Factor Authentication (2FA)"),
          bullet("Databases: PostgreSQL (Views, Materialized Views, Functions, Triggers, Schema-per-Tenant Architecture), MongoDB"),
          bullet("AI & Integration: Open-Source LLM Integration, Model Context Protocol (MCP), AI-Powered Chatbot Integration, Reusable Angular Component Library Development, Embed SDK Development"),
          bullet("DevOps & Deployment: Nginx (Reverse Proxy), Apache HTTP Server, GitLab CI/CD, DigitalOcean, Ubuntu Server, Cron Jobs"),
          bullet("Tools & Platforms: Git, Self-Hosted GitLab (Project and Repository Management), Postman, pgAdmin"),
          blank(),

          // Professional Experience
          heading("Professional Experience"),
          new Paragraph({ children: [new TextRun({ text: "Megam Technologies", bold: true, size: 21, color: "111827" }), new TextRun({ text: "   December 2024 – Present", bold: true, size: 19, color: "1E293B" })] }),
          new Paragraph({ children: [new TextRun({ text: "Software Developer | Chennai, India", italics: true, size: 19, color: "475569" })] }),
          new Paragraph({ children: [new TextRun({ text: "Datalytics — AI-Powered Business Intelligence SaaS Platform", bold: true, size: 19, color: "1B365D" })] }),
          bullet("Developed and enhanced SaaS platform features, including interactive dashboards, reporting modules, and business data visualization for enterprise users."),
          bullet("Built interactive dashboards to track employee working hours, time spent, project progress, and issue metrics, enabling users to monitor operational performance."),
          bullet("Integrated a self-hosted open-source LLM with MCP-based tools to build a natural-language chatbot capable of retrieving company-specific data, answering business queries, and generating reports."),
          bullet("Implemented TOTP-based Two-Factor Authentication (2FA) and developed responsive Angular features integrated with Node.js and NestJS REST APIs."),
          bullet("Created and configured 150+ business reports within the Datalytics platform for client-facing and internal use cases, covering project tasks, employee working hours, time tracking, issue management, and operational metrics, leveraging MCP-based integrations where applicable."),
          new Paragraph({ children: [new TextRun({ text: "Datalytics — Angular Component Library Modernization", bold: true, size: 19, color: "1B365D" })] }),
          bullet("Modernized an existing Angular component library by migrating 30+ interactive charts from the legacy charting implementation to Apache ECharts while preserving existing functionality."),
          bullet("Published and maintained the Angular component library, managing version releases and package updates for integration across multiple applications."),
          bullet("Integrated the published library into consuming applications, enabling reuse of standardized chart components and visualizations."),
          new Paragraph({ children: [new TextRun({ text: "Datalytics — Embedded Analytics SDK", bold: true, size: 19, color: "1B365D" })] }),
          bullet("Developed a standalone JavaScript Embed SDK with token-based authentication, enabling external applications to embed Datalytics dashboards and reports."),
          bullet("Integrated the SDK into Datalytics Docs, with the reusable SDK subsequently adopted by other internal applications, including Covalens."),
          new Paragraph({ children: [new TextRun({ text: "LeatherLink TanPro 11.0 — Tannery Management System (TMS)", bold: true, size: 19, color: "1B365D" })] }),
          bullet("Contributed to the end-to-end development, implementation, and maintenance of a multi-tenant tannery ERP comprising 300+ database tables and approximately 200K+ lines of code, supporting production workflows, batch tracking, and stock assortment."),
          bullet("Implemented and extended PostgreSQL database functionality using schema-per-tenant architecture, including SQL queries, views, materialized views, functions, and triggers."),
          bullet("Improved database operations through SQL query optimization and implemented MongoDB-based audit logging and multi-level approval workflows."),
          bullet("Managed application setup, configuration, and deployment across client environments, including migrating PHP applications to Ubuntu servers hosted on DigitalOcean."),
          bullet("Configured Apache web servers and automated scheduled operations using cron jobs; implemented PDF generation using mPDF for GST invoices, delivery challans, and job cards."),
          bullet("Collaborated directly with clients to understand business requirements, translate workflows into application enhancements, resolve production issues, and deliver requested features."),
          blank(),

          // Personal Projects
          heading("Personal Projects"),
          new Paragraph({ children: [new TextRun({ text: "KYAAT – Gesture-Controlled Mouse", bold: true, size: 20, color: "111827" }), new TextRun({ text: "  |  Python, OpenCV", italics: true, size: 18, color: "475569" })] }),
          new Paragraph({ children: [new TextRun({ text: "Developed a computer vision application using OpenCV to recognize real-time hand gestures and translate them into mouse cursor movements and actions, achieving sub-20ms response latency.", size: 19 })] }),
          blank(),
          new Paragraph({ children: [new TextRun({ text: "Personal Portfolio", bold: true, size: 20, color: "111827" }), new TextRun({ text: "  |  HTML, CSS, JavaScript", italics: true, size: 18, color: "475569" })] }),
          new Paragraph({ children: [new TextRun({ text: "Designed and developed a responsive portfolio website to showcase technical skills, professional experience, and project demonstrations. Portfolio: https://alanz.netlify.app/", size: 19 })] }),
          blank(),

          // Education
          heading("Education"),
          new Paragraph({ children: [new TextRun({ text: "Bachelor of Technology in Information Technology — Anna University", bold: true, size: 21, color: "111827" })] }),
          new Paragraph({ children: [new TextRun({ text: "2020 – 2024  |  80% First class", size: 19, color: "475569" })] }),
          blank(),

          // Achievements & Certifications
          heading("Achievements & Certifications"),
          bullet("The Versatile Rockstar Award (Megam Technologies, 2026)"),
          bullet("Full Stack Development (Infosys Springboard — Frontend Frameworks, Backend APIs, & Databases)"),
          bullet("Python Programming (Spoken Tutorial Project, IIT Bombay)")
        ]
      }]
    });

    Packer.toBlob(doc).then(blob => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
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
function blank() { return new docx.Paragraph({ children: [] }); }
function heading(text) {
  return new docx.Paragraph({
    children: [new docx.TextRun({ text: text.toUpperCase(), bold: true, size: 20, color: "1B365D" })],
    border: { bottom: { style: "single", size: 8, color: "1B365D", space: 2 } }
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
