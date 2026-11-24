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

// ── PDF Export ────────────────────────────────
function exportPDF() {
  const wasEdit = isEditMode;
  if (wasEdit) toggleEdit();

  const toolbar = document.getElementById("toolbar");
  toolbar.style.display = "none";

  const el  = document.getElementById("resume");
  const opt = {
    margin:      [8, 8, 8, 8],
    filename:    "Boobalan_V_Resume.pdf",
    image:       { type: "jpeg", quality: 0.98 },
    html2canvas: { scale: 2, useCORS: true, logging: false },
    jsPDF:       { unit: "mm", format: "a4", orientation: "portrait" },
    pagebreak:   { mode: ["avoid-all", "css", "legacy"] }
  };

  showToast("Generating PDF…");

  html2pdf().set(opt).from(el).save().then(() => {
    toolbar.style.display = "";
    showToast("✅ PDF saved!");
    if (wasEdit) toggleEdit();
  }).catch(err => {
    toolbar.style.display = "";
    showToast("❌ " + err.message);
    if (wasEdit) toggleEdit();
  });
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
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "LinkedIn: Boobalan V.  |  GitHub: boobalanIT  |  Portfolio: alanz.netlify.app", size: 18, color: "475569" })] }),
          blank(),

          // Professional Summary
          heading("Professional Summary"),
          new Paragraph({ children: [new TextRun({ text: t("summary"), size: 19 })] }),
          blank(),

          // Technical Skills
          heading("Technical Skills"),
          bullet("Frontend: Angular, HTML5, CSS3, JavaScript, Bootstrap"),
          bullet("Backend: Node.js, NestJS, Express.js, PHP, REST APIs"),
          bullet("Databases: PostgreSQL, MongoDB, MySQL"),
          bullet("Web Servers & Deployment: Nginx, Apache, Application Deployment"),
          bullet("Tools & CI/CD: Git, GitLab CI/CD, Postman, pgAdmin"),
          bullet("AI & Integration: Open-source LLM integration, Model Context Protocol (MCP), AI-powered chatbot integration"),
          blank(),

          // Professional Experience
          heading("Professional Experience"),
          new Paragraph({ children: [new TextRun({ text: "Megam Technologies", bold: true, size: 21, color: "111827" }), new TextRun({ text: "   December 2024 – Present", bold: true, size: 19, color: "1E293B" })] }),
          new Paragraph({ children: [new TextRun({ text: "Software Developer | Chennai, India", italics: true, size: 19, color: "475569" })] }),
          new Paragraph({ children: [new TextRun({ text: "Datalytics — Business Intelligence & Analytics Platform", bold: true, size: 19, color: "1B365D" })] }),
          bullet("Developed and maintained features for a business intelligence platform, including MCP server tools that enable AI-assisted interaction with platform data and functionality."),
          bullet("Integrated an open-source large language model (LLM) to support natural-language questions about business data and provide responses based on analytics information."),
          bullet("Implemented functionality to generate reports and dashboards from user requests, covering operational metrics such as working hours, time spent, and issues created or closed."),
          bullet("Supported comparative analysis of employee work activity and time allocation, alongside bug fixing, feature enhancements, and application maintenance."),
          new Paragraph({ children: [new TextRun({ text: "ERP — Enterprise Resource Planning Application", bold: true, size: 19, color: "1B365D" })] }),
          bullet("Developed and maintained ERP application features using PHP and PostgreSQL; implemented CRUD operations and integrated application functionality with existing database and system workflows."),
          bullet("Created and modified SQL queries, views, materialized views, functions, and triggers to support business requirements; troubleshot application and database issues and implemented fixes."),
          new Paragraph({ children: [new TextRun({ text: "Deployment & Development Operations", bold: true, size: 19, color: "1B365D" })] }),
          bullet("Configured Nginx as a reverse proxy for Datalytics (Angular & Node.js/NestJS) and Apache HTTP Server for LeatherLink ERP (.htaccess modules, virtual hosts); supported server configuration and deployments."),
          bullet("Worked with GitLab CI/CD pipelines to support application build and deployment workflows; used Git for version control and Postman for API testing and troubleshooting."),
          blank(),

          new Paragraph({ children: [new TextRun({ text: "Shiash Info Solutions Private Limited", bold: true, size: 21, color: "111827" }), new TextRun({ text: "   March 2023 – May 2023", bold: true, size: 19, color: "1E293B" })] }),
          new Paragraph({ children: [new TextRun({ text: "Junior Full-Stack Developer | Chennai, India", italics: true, size: 19, color: "475569" })] }),
          bullet("Contributed to full-stack web development tasks using HTML, CSS, and JavaScript; assisted with feature implementation and debugging within an established software development team and workflow."),
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
          bullet("Python Programming — Spoken Tutorial / CodeTantra"),
          bullet("Full Stack Development — Infosys Springboard"),
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
