const ExcelJS = require('exceljs');
const path = require('path');

async function createResume() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Boobalan V';
  workbook.lastModifiedBy = 'Boobalan V';
  workbook.created = new Date();
  workbook.modified = new Date();

  // Color Palette matching reference image
  const NAVY = 'FF1B365D';         // Deep navy blue for headings & title
  const DARK_TEXT = 'FF1E293B';    // Slate-800 for body text
  const MUTED_TEXT = 'FF475569';   // Slate-600 for subheadings/metadata
  const LINK_BLUE = 'FF0969DA';    // Professional link blue
  const BORDER_COLOR = 'FF1B365D'; // Navy border for section dividers
  const LIGHT_ROW_BG = 'FFF8FAFC'; // Very subtle zebra/accent
  const BULLET_CHAR = '•';

  // Helper function to build the Executive Resume worksheet
  function buildExecutiveWorksheet(sheetName, showGridLines) {
    const ws = workbook.addWorksheet(sheetName, {
      views: [{ showGridLines: showGridLines }],
      pageSetup: {
        paperSize: 9, // A4
        orientation: 'portrait',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        margins: {
          left: 0.4,
          right: 0.4,
          top: 0.4,
          bottom: 0.4,
          header: 0.2,
          footer: 0.2
        }
      }
    });

    // Column Layout: Total printable width ~ 88-90 units
    ws.columns = [
      { key: 'colA', width: 2.5 },   // Left margin / indent
      { key: 'colB', width: 3.5 },   // Bullet symbol
      { key: 'colC', width: 17.0 },  // Content 1
      { key: 'colD', width: 20.0 },  // Content 2
      { key: 'colE', width: 20.0 },  // Content 3
      { key: 'colF', width: 15.0 },  // Content 4
      { key: 'colG', width: 18.0 }   // Date / Right alignment
    ];

    let currentRow = 1;

    function addSpacer(height = 6) {
      const row = ws.getRow(currentRow++);
      row.height = height;
      return row;
    }

    function addSectionHeader(title) {
      addSpacer(8);
      const row = ws.getRow(currentRow++);
      row.height = 20;
      ws.mergeCells(`B${row.number}:G${row.number}`);
      const cell = ws.getCell(`B${row.number}`);
      cell.value = title.toUpperCase();
      cell.font = {
        name: 'Calibri',
        size: 11,
        bold: true,
        color: { argb: NAVY }
      };
      cell.alignment = { vertical: 'bottom', horizontal: 'left' };
      
      // Bottom border across entire width from B to G
      for (let col = 2; col <= 7; col++) {
        const c = row.getCell(col);
        c.border = {
          bottom: { style: 'medium', color: { argb: BORDER_COLOR } }
        };
      }
      addSpacer(4);
    }

    // 1. TOP HEADER - NAME
    addSpacer(4);
    const nameRow = ws.getRow(currentRow++);
    nameRow.height = 28;
    ws.mergeCells(`B${nameRow.number}:G${nameRow.number}`);
    const nameCell = ws.getCell(`B${nameRow.number}`);
    nameCell.value = 'BOOBALAN V';
    nameCell.font = {
      name: 'Calibri',
      size: 20,
      bold: true,
      color: { argb: NAVY }
    };
    nameCell.alignment = { vertical: 'middle', horizontal: 'center' };

    // Subtitle / Professional Title
    const titleRow = ws.getRow(currentRow++);
    titleRow.height = 16;
    ws.mergeCells(`B${titleRow.number}:G${titleRow.number}`);
    const titleCell = ws.getCell(`B${titleRow.number}`);
    titleCell.value = 'Software Developer | Full-Stack Web Applications';
    titleCell.font = {
      name: 'Calibri',
      size: 10.5,
      bold: true,
      color: { argb: DARK_TEXT }
    };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

    // Contact Info Line 1 (Location, Phone, Email)
    const contactRow1 = ws.getRow(currentRow++);
    contactRow1.height = 15;
    ws.mergeCells(`B${contactRow1.number}:G${contactRow1.number}`);
    const contactCell1 = ws.getCell(`B${contactRow1.number}`);
    contactCell1.value = {
      richText: [
        { text: 'Mayiladuthurai, Tamil Nadu, India  |  +91 9025006329  |  ', font: { name: 'Calibri', size: 9.5, color: { argb: MUTED_TEXT } } },
        { 
          text: 'boobalanvaithiyanathan@gmail.com',
          font: { name: 'Calibri', size: 9.5, color: { argb: LINK_BLUE }, underline: true }
        }
      ]
    };
    contactCell1.alignment = { vertical: 'middle', horizontal: 'center' };

    // Contact Info Line 2 (LinkedIn, GitHub, Portfolio)
    const contactRow2 = ws.getRow(currentRow++);
    contactRow2.height = 15;
    ws.mergeCells(`B${contactRow2.number}:G${contactRow2.number}`);
    const contactCell2 = ws.getCell(`B${contactRow2.number}`);
    contactCell2.value = {
      richText: [
        { text: 'LinkedIn: ', font: { name: 'Calibri', size: 9.5, color: { argb: MUTED_TEXT } } },
        { text: 'linkedin.com/in/boobalan', font: { name: 'Calibri', size: 9.5, color: { argb: LINK_BLUE }, underline: true } },
        { text: '  |  GitHub: ', font: { name: 'Calibri', size: 9.5, color: { argb: MUTED_TEXT } } },
        { text: 'github.com/boobalan', font: { name: 'Calibri', size: 9.5, color: { argb: LINK_BLUE }, underline: true } },
        { text: '  |  Portfolio: ', font: { name: 'Calibri', size: 9.5, color: { argb: MUTED_TEXT } } },
        { text: 'alanz-tech.netlify.app', font: { name: 'Calibri', size: 9.5, color: { argb: LINK_BLUE }, underline: true } }
      ]
    };
    contactCell2.alignment = { vertical: 'middle', horizontal: 'center' };

    // -------------------------------------------------------------
    // 2. PROFESSIONAL SUMMARY
    // -------------------------------------------------------------
    addSectionHeader('PROFESSIONAL SUMMARY');
    
    const summaryRow = ws.getRow(currentRow++);
    summaryRow.height = 58;
    ws.mergeCells(`B${summaryRow.number}:G${summaryRow.number}`);
    const summaryCell = ws.getCell(`B${summaryRow.number}`);
    summaryCell.value = 'Software Developer with nearly 2 years of professional experience building and maintaining web applications. Experienced in Angular, Node.js/NestJS, PHP, MongoDB, and PostgreSQL, with hands-on work across frontend features, backend functionality, and database operations. Contributed to a business intelligence platform by implementing features, integrating AI capabilities with an open-source LLM using MCP, and resolving bugs. Interested in building reliable, maintainable applications and continuing to grow in full-stack development.';
    summaryCell.font = { name: 'Calibri', size: 9.5, color: { argb: DARK_TEXT } };
    summaryCell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };

    // -------------------------------------------------------------
    // 3. PROFESSIONAL EXPERIENCE
    // -------------------------------------------------------------
    addSectionHeader('PROFESSIONAL EXPERIENCE');

    // Experience 1: Megam Technologies
    const exp1Header = ws.getRow(currentRow++);
    exp1Header.height = 18;
    ws.mergeCells(`B${exp1Header.number}:E${exp1Header.number}`);
    const exp1CompCell = ws.getCell(`B${exp1Header.number}`);
    exp1CompCell.value = 'Megam Technologies';
    exp1CompCell.font = { name: 'Calibri', size: 10.5, bold: true, color: { argb: DARK_TEXT } };
    exp1CompCell.alignment = { vertical: 'middle', horizontal: 'left' };

    ws.mergeCells(`F${exp1Header.number}:G${exp1Header.number}`);
    const exp1DateCell = ws.getCell(`F${exp1Header.number}`);
    exp1DateCell.value = 'Dec 2024 – Present';
    exp1DateCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: DARK_TEXT } };
    exp1DateCell.alignment = { vertical: 'middle', horizontal: 'right' };

    // Subtitle
    const exp1Sub = ws.getRow(currentRow++);
    exp1Sub.height = 15;
    ws.mergeCells(`B${exp1Sub.number}:G${exp1Sub.number}`);
    const exp1SubCell = ws.getCell(`B${exp1Sub.number}`);
    exp1SubCell.value = 'Software Developer | Chennai / Tamil Nadu, India';
    exp1SubCell.font = { name: 'Calibri', size: 9.5, italic: true, color: { argb: MUTED_TEXT } };
    exp1SubCell.alignment = { vertical: 'middle', horizontal: 'left' };

    function addBullet(bulletText, estimatedLines = 1) {
      const row = ws.getRow(currentRow++);
      row.height = estimatedLines === 1 ? 17 : (estimatedLines === 2 ? 30 : 44);
      
      const bCell = ws.getCell(`B${row.number}`);
      bCell.value = BULLET_CHAR;
      bCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: NAVY } };
      bCell.alignment = { vertical: 'top', horizontal: 'center' };

      ws.mergeCells(`C${row.number}:G${row.number}`);
      const textCell = ws.getCell(`C${row.number}`);
      textCell.value = bulletText;
      textCell.font = { name: 'Calibri', size: 9.5, color: { argb: DARK_TEXT } };
      textCell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    }

    addBullet('Develop and maintain features for Datalytics, a business intelligence and analytics platform, using the technologies relevant to assigned application components.', 2);
    addBullet('Integrate AI capabilities with an open-source large language model using MCP to support AI-enabled product functionality.', 2);
    addBullet('Investigate and fix application bugs, maintain existing features, and support ongoing product improvements.', 1);
    addBullet('Work with Angular, Node.js/NestJS, and MongoDB across application development tasks.', 1);
    addBullet('Develop and maintain ERP functionality using PHP and PostgreSQL.', 1);
    addBullet('Write and work with PostgreSQL queries, joins, views, materialized views, triggers, functions, and CRUD operations.', 2);
    addBullet('Build responsive user interfaces using HTML, CSS, and JavaScript, and manage source code and deployments through GitLab and CI/CD workflows.', 2);

    addSpacer(5);

    // Experience 2: Shiash Info Solutions
    const exp2Header = ws.getRow(currentRow++);
    exp2Header.height = 18;
    ws.mergeCells(`B${exp2Header.number}:E${exp2Header.number}`);
    const exp2CompCell = ws.getCell(`B${exp2Header.number}`);
    exp2CompCell.value = 'Shiash Info Solutions Private Limited';
    exp2CompCell.font = { name: 'Calibri', size: 10.5, bold: true, color: { argb: DARK_TEXT } };
    exp2CompCell.alignment = { vertical: 'middle', horizontal: 'left' };

    ws.mergeCells(`F${exp2Header.number}:G${exp2Header.number}`);
    const exp2DateCell = ws.getCell(`F${exp2Header.number}`);
    exp2DateCell.value = 'Mar 2023 – May 2023';
    exp2DateCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: DARK_TEXT } };
    exp2DateCell.alignment = { vertical: 'middle', horizontal: 'right' };

    const exp2Sub = ws.getRow(currentRow++);
    exp2Sub.height = 15;
    ws.mergeCells(`B${exp2Sub.number}:G${exp2Sub.number}`);
    const exp2SubCell = ws.getCell(`B${exp2Sub.number}`);
    exp2SubCell.value = 'Junior Full-Stack Developer | Chennai, Tamil Nadu, India';
    exp2SubCell.font = { name: 'Calibri', size: 9.5, italic: true, color: { argb: MUTED_TEXT } };
    exp2SubCell.alignment = { vertical: 'middle', horizontal: 'left' };

    addBullet('Assisted in developing web application modules using HTML, CSS, and JavaScript.', 1);
    addBullet('Gained hands-on experience with full-stack development workflows and collaborative software development practices.', 1);

    // -------------------------------------------------------------
    // 4. TECHNICAL SKILLS
    // -------------------------------------------------------------
    addSectionHeader('TECHNICAL SKILLS');

    function addSkillCategory(categoryName, skillsList) {
      const row = ws.getRow(currentRow++);
      row.height = 18;
      ws.mergeCells(`B${row.number}:G${row.number}`);
      const cell = ws.getCell(`B${row.number}`);
      cell.value = {
        richText: [
          { 
            text: categoryName + ': ', 
            font: { name: 'Calibri', size: 9.5, bold: true, color: { argb: DARK_TEXT } } 
          },
          { 
            text: skillsList, 
            font: { name: 'Calibri', size: 9.5, bold: false, color: { argb: DARK_TEXT } } 
          }
        ]
      };
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
    }

    addSkillCategory('Frontend', 'Angular, HTML, CSS, JavaScript, Bootstrap');
    addSkillCategory('Backend', 'Node.js, NestJS, PHP');
    addSkillCategory('Databases', 'MongoDB, PostgreSQL, MySQL');
    addSkillCategory('AI / Integration', 'Open-source LLM integration, Model Context Protocol (MCP)');
    addSkillCategory('Tools', 'Git, GitLab, Postman, pgAdmin, CI/CD pipelines');

    // -------------------------------------------------------------
    // 5. SELECTED PERSONAL PROJECTS
    // -------------------------------------------------------------
    addSectionHeader('SELECTED PERSONAL PROJECTS');

    function addProject(title, description, lines = 2) {
      const row = ws.getRow(currentRow++);
      row.height = lines === 1 ? 18 : 28;
      
      const bCell = ws.getCell(`B${row.number}`);
      bCell.value = BULLET_CHAR;
      bCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: NAVY } };
      bCell.alignment = { vertical: 'top', horizontal: 'center' };

      ws.mergeCells(`C${row.number}:G${row.number}`);
      const textCell = ws.getCell(`C${row.number}`);
      textCell.value = {
        richText: [
          { 
            text: `${title} | `, 
            font: { name: 'Calibri', size: 9.5, bold: true, color: { argb: DARK_TEXT } } 
          },
          { 
            text: description, 
            font: { name: 'Calibri', size: 9.5, color: { argb: DARK_TEXT } } 
          }
        ]
      };
      textCell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    }

    addProject('KYAAT – Gesture-Controlled Mouse', 'Built a hand-gesture-controlled mouse using Python and OpenCV to support touch-free cursor control.', 1);
    addProject('MSAJCE Canteen App', 'Developed a college canteen food-ordering app using XML and Firebase to support meal pre-ordering.', 1);
    addProject('Personal Portfolio', 'Created a responsive portfolio website (alanz-tech.netlify.app) to showcase technical skills and web projects.', 1);

    // -------------------------------------------------------------
    // 6. EDUCATION
    // -------------------------------------------------------------
    addSectionHeader('EDUCATION');

    const eduRow = ws.getRow(currentRow++);
    eduRow.height = 18;
    ws.mergeCells(`B${eduRow.number}:E${eduRow.number}`);
    const eduTitleCell = ws.getCell(`B${eduRow.number}`);
    eduTitleCell.value = 'B.Tech in Information Technology — Mohamed Sathak A J College of Engineering';
    eduTitleCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: DARK_TEXT } };
    eduTitleCell.alignment = { vertical: 'middle', horizontal: 'left' };

    ws.mergeCells(`F${eduRow.number}:G${eduRow.number}`);
    const eduDateCell = ws.getCell(`F${eduRow.number}`);
    eduDateCell.value = '2020 – 2024';
    eduDateCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: DARK_TEXT } };
    eduDateCell.alignment = { vertical: 'middle', horizontal: 'right' };

    const eduSub = ws.getRow(currentRow++);
    eduSub.height = 15;
    ws.mergeCells(`B${eduSub.number}:G${eduSub.number}`);
    const eduSubCell = ws.getCell(`B${eduSub.number}`);
    eduSubCell.value = 'Academic Performance: 80% (First Class with Distinction)';
    eduSubCell.font = { name: 'Calibri', size: 9.5, italic: true, color: { argb: MUTED_TEXT } };
    eduSubCell.alignment = { vertical: 'middle', horizontal: 'left' };

    // -------------------------------------------------------------
    // 7. CERTIFICATIONS & ACHIEVEMENTS
    // -------------------------------------------------------------
    addSectionHeader('CERTIFICATIONS & ACHIEVEMENTS');

    function addCert(certText) {
      const row = ws.getRow(currentRow++);
      row.height = 17;
      const bCell = ws.getCell(`B${row.number}`);
      bCell.value = BULLET_CHAR;
      bCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: NAVY } };
      bCell.alignment = { vertical: 'middle', horizontal: 'center' };

      ws.mergeCells(`C${row.number}:G${row.number}`);
      const textCell = ws.getCell(`C${row.number}`);
      textCell.value = certText;
      textCell.font = { name: 'Calibri', size: 9.5, color: { argb: DARK_TEXT } };
      textCell.alignment = { vertical: 'middle', horizontal: 'left' };
    }

    addCert('Good Performer Award — 2026 (Awarded for outstanding technical contribution & performance)');
    addCert('Full Stack Development — Infosys Springboard');
    addCert('Python Certification — Spoken Tutorial / CodeTantra');

    addSpacer(10);
  }

  // Build Sheet 1: Clean Executive View (Without gridlines, looks like document)
  buildExecutiveWorksheet('Executive Resume', false);

  // Build Sheet 2: Executive View With Gridlines (Easy for editing in Excel)
  buildExecutiveWorksheet('Executive (Gridlines On)', true);

  // -------------------------------------------------------------
  // SHEET 3: ATS & Tabular Recruiter View (Structured Data)
  // -------------------------------------------------------------
  const ws3 = workbook.addWorksheet('Tabular Data (ATS-Friendly)', {
    views: [{ showGridLines: true }]
  });

  ws3.columns = [
    { header: 'Section', key: 'section', width: 20 },
    { header: 'Category / Organization', key: 'cat', width: 32 },
    { header: 'Role / Degree / Project', key: 'item', width: 32 },
    { header: 'Duration / Year', key: 'period', width: 18 },
    { header: 'Location / Platform', key: 'loc', width: 25 },
    { header: 'Key Details / Technologies / Scope', key: 'details', width: 60 }
  ];

  // Header styling
  const headerRow3 = ws3.getRow(1);
  headerRow3.height = 24;
  headerRow3.eachCell((cell) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: NAVY }
    };
    cell.font = {
      name: 'Calibri',
      size: 10.5,
      bold: true,
      color: { argb: 'FFFFFFFF' }
    };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
  });

  const atsData = [
    { section: 'Personal Information', cat: 'Full Name', item: 'Boobalan V', period: '-', loc: 'Mayiladuthurai, Tamil Nadu, India', details: 'Phone: +91 9025006329 | Email: boobalanvaithiyanathan@gmail.com' },
    { section: 'Personal Information', cat: 'Links', item: 'Online Profiles', period: '-', loc: 'Web', details: 'LinkedIn: linkedin.com/in/boobalan | GitHub: github.com/boobalan | Portfolio: alanz-tech.netlify.app' },
    { section: 'Summary', cat: 'Professional Summary', item: 'Software Developer', period: '~2 Years Exp', loc: 'India', details: 'Full-stack developer experienced in Angular, Node.js/NestJS, PHP, PostgreSQL, MongoDB, and LLM MCP integration.' },
    { section: 'Experience', cat: 'Megam Technologies', item: 'Software Developer', period: 'Dec 2024 – Present', loc: 'Tamil Nadu, India', details: 'Develop features for Datalytics BI platform, LLM AI integration via MCP, Angular, Node.js/NestJS, MongoDB, ERP modules in PHP & PostgreSQL.' },
    { section: 'Experience', cat: 'Shiash Info Solutions', item: 'Junior Full-Stack Developer', period: 'Mar 2023 – May 2023', loc: 'Chennai, India', details: 'Assisted in developing web application modules using HTML, CSS, JavaScript; engaged in collaborative agile workflows.' },
    { section: 'Technical Skills', cat: 'Frontend', item: 'Client-side Engineering', period: 'Active', loc: '-', details: 'Angular, HTML, CSS, JavaScript, Bootstrap' },
    { section: 'Technical Skills', cat: 'Backend', item: 'Server & Microservices', period: 'Active', loc: '-', details: 'Node.js, NestJS, PHP' },
    { section: 'Technical Skills', cat: 'Databases', item: 'Relational & NoSQL', period: 'Active', loc: '-', details: 'MongoDB, PostgreSQL, MySQL' },
    { section: 'Technical Skills', cat: 'AI / Integration', item: 'Generative AI & LLMs', period: 'Active', loc: '-', details: 'Open-source LLM integration, Model Context Protocol (MCP)' },
    { section: 'Technical Skills', cat: 'Tools', item: 'Developer Tooling & DevOps', period: 'Active', loc: '-', details: 'Git, GitLab, Postman, pgAdmin, CI/CD pipelines' },
    { section: 'Personal Projects', cat: 'KYAAT', item: 'Gesture-Controlled Mouse', period: 'Project', loc: 'Desktop', details: 'Python, OpenCV hand-gesture computer mouse control supporting touch-free cursor navigation.' },
    { section: 'Personal Projects', cat: 'MSAJCE Canteen App', item: 'Food-Ordering App', period: 'Project', loc: 'Mobile', details: 'XML, Firebase real-time meal pre-ordering application for campus cafeteria.' },
    { section: 'Personal Projects', cat: 'Personal Portfolio', item: 'Web Portfolio', period: 'Active', loc: 'alanz-tech.netlify.app', details: 'Created responsive portfolio website to showcase technical capabilities and live project demos.' },
    { section: 'Education', cat: 'Mohamed Sathak A J College', item: 'B.Tech in Information Technology', period: '2020 – 2024', loc: 'Chennai, India', details: 'Graduated with 80% aggregate (First Class with Distinction).' },
    { section: 'Certifications', cat: 'Megam Technologies', item: 'Good Performer Award', period: '2026', loc: 'Workplace', details: 'Recognized for outstanding technical contribution & project performance.' },
    { section: 'Certifications', cat: 'Infosys Springboard', item: 'Full Stack Development', period: 'Certified', loc: 'Infosys', details: 'Comprehensive enterprise full-stack development certification.' },
    { section: 'Certifications', cat: 'Spoken Tutorial / CodeTantra', item: 'Python Programming', period: 'Certified', loc: 'National Mission', details: 'Foundational and advanced Python language and scripting certification.' }
  ];

  atsData.forEach((item, idx) => {
    const row = ws3.addRow(item);
    row.height = 20;
    const isEven = idx % 2 === 0;
    row.eachCell((cell) => {
      cell.font = { name: 'Calibri', size: 9.5 };
      cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      if (isEven) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: LIGHT_ROW_BG }
        };
      }
      cell.border = {
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
      };
    });
  });

  const outputPath = path.join(__dirname, 'Boobalan_V_Resume.xlsx');
  await workbook.xlsx.writeFile(outputPath);
  console.log('Successfully generated Excel resume at:', outputPath);
}

createResume().catch(err => {
  console.error('Error generating Excel resume:', err);
  process.exit(1);
});
