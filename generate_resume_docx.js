const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  BorderStyle,
  Table,
  TableRow,
  TableCell,
  WidthType,
  VerticalAlign,
  TableLayoutType,
  convertInchesToTwip,
  UnderlineType,
  ExternalHyperlink
} = require('docx');
const fs = require('fs');
const path = require('path');

// --- Color Palette ---
const NAVY  = '1B365D';
const DARK  = '1E293B';
const MUTED = '475569';
const LINK  = '0969DA';
const BLACK = '111827';

// --- Typography ---
const FONT       = 'Calibri';
const SIZE_NAME  = 46;   // 23pt
const SIZE_TITLE = 22;   // 11pt
const SIZE_BODY  = 18;   // 9pt (perfect 1-page fit in Word)
const SIZE_SECT  = 20;   // 10pt bold uppercase
const SIZE_COMP  = 20;   // 10pt bold

// Spacing tokens (in twips: 20 twips = 1 pt)
const SP_BEFORE_SECTION = 45;
const SP_AFTER_SECTION  = 14;
const SP_BULLET         = 8;

function sectionHeading(title) {
  return new Paragraph({
    spacing: { before: SP_BEFORE_SECTION, after: SP_AFTER_SECTION },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 2 }
    },
    children: [
      new TextRun({
        text: title.toUpperCase(),
        font: FONT,
        size: SIZE_SECT,
        bold: true,
        color: NAVY
      })
    ]
  });
}

function companyDateRow(company, date) {
  return new Table({
    layout: TableLayoutType.FIXED,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE, size: 0 },
      bottom: { style: BorderStyle.NONE, size: 0 },
      left: { style: BorderStyle.NONE, size: 0 },
      right: { style: BorderStyle.NONE, size: 0 },
      insideHorizontal: { style: BorderStyle.NONE, size: 0 },
      insideVertical: { style: BorderStyle.NONE, size: 0 }
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            borders: {
              top: { style: BorderStyle.NONE, size: 0 },
              bottom: { style: BorderStyle.NONE, size: 0 },
              left: { style: BorderStyle.NONE, size: 0 },
              right: { style: BorderStyle.NONE, size: 0 }
            },
            children: [
              new Paragraph({
                spacing: { before: 40, after: 15 },
                children: [
                  new TextRun({
                    text: company,
                    font: FONT,
                    size: SIZE_COMP,
                    bold: true,
                    color: BLACK
                  })
                ]
              })
            ]
          }),
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            borders: {
              top: { style: BorderStyle.NONE, size: 0 },
              bottom: { style: BorderStyle.NONE, size: 0 },
              left: { style: BorderStyle.NONE, size: 0 },
              right: { style: BorderStyle.NONE, size: 0 }
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                spacing: { before: 40, after: 15 },
                children: [
                  new TextRun({
                    text: date,
                    font: FONT,
                    size: SIZE_BODY,
                    bold: true,
                    color: DARK
                  })
                ]
              })
            ]
          })
        ]
      })
    ]
  });
}

function roleSubtitle(text) {
  return new Paragraph({
    spacing: { before: 10, after: 30 },
    children: [
      new TextRun({
        text,
        font: FONT,
        size: SIZE_BODY,
        italics: true,
        color: MUTED
      })
    ]
  });
}

function expSubHeading(text) {
  return new Paragraph({
    spacing: { before: 32, after: 10 },
    children: [
      new TextRun({
        text,
        font: FONT,
        size: SIZE_BODY,
        bold: true,
        color: NAVY
      })
    ]
  });
}

function bullet(text) {
  return new Paragraph({
    spacing: { before: SP_BULLET, after: SP_BULLET },
    indent: { left: convertInchesToTwip(0.2), hanging: convertInchesToTwip(0.14) },
    children: [
      new TextRun({ text: '\u2022\t', font: FONT, size: SIZE_BODY, color: NAVY }),
      new TextRun({ text, font: FONT, size: SIZE_BODY, color: BLACK })
    ]
  });
}

function bulletWithBold(boldPrefix, text) {
  return new Paragraph({
    spacing: { before: SP_BULLET, after: SP_BULLET },
    indent: { left: convertInchesToTwip(0.2), hanging: convertInchesToTwip(0.14) },
    children: [
      new TextRun({ text: '\u2022\t', font: FONT, size: SIZE_BODY, color: NAVY }),
      new TextRun({ text: boldPrefix, font: FONT, size: SIZE_BODY, bold: true, color: BLACK }),
      new TextRun({ text, font: FONT, size: SIZE_BODY, color: BLACK })
    ]
  });
}

function link(label, url) {
  return new ExternalHyperlink({
    link: url,
    children: [
      new TextRun({
        text: label,
        font: FONT,
        size: SIZE_BODY,
        color: LINK,
        underline: { type: UnderlineType.SINGLE }
      })
    ]
  });
}

function sep() {
  return new TextRun({ text: '  |  ', font: FONT, size: SIZE_BODY, color: MUTED });
}

async function buildDocument() {
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: FONT, size: SIZE_BODY, color: BLACK }
        }
      }
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: 11906,  // A4 width in twips (210mm)
              height: 16838  // A4 height in twips (297mm)
            },
            margin: {
              top: convertInchesToTwip(0.32),
              bottom: convertInchesToTwip(0.32),
              left: convertInchesToTwip(0.48),
              right: convertInchesToTwip(0.48)
            }
          }
        },
        children: [
          // Header
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 20 },
            children: [
              new TextRun({
                text: 'BOOBALAN V',
                font: FONT,
                size: SIZE_NAME,
                bold: true,
                color: NAVY
              })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 25 },
            children: [
              new TextRun({
                text: 'Software Developer | Full-Stack Developer',
                font: FONT,
                size: SIZE_TITLE,
                bold: true,
                color: DARK
              })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 20 },
            children: [
              new TextRun({ text: 'Chennai, Tamil Nadu, India', font: FONT, size: SIZE_BODY, color: MUTED }),
              sep(),
              new TextRun({ text: '+91 9025006329', font: FONT, size: SIZE_BODY, color: MUTED }),
              sep(),
              link('boobalanvaithiyanathan1@gmail.com', 'mailto:boobalanvaithiyanathan1@gmail.com')
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 50 },
            children: [
              new TextRun({ text: 'LinkedIn: ', font: FONT, size: SIZE_BODY, color: MUTED }),
              link('Boobalan V', 'https://www.linkedin.com/in/boobalan001/'),
              sep(),
              new TextRun({ text: 'GitHub: ', font: FONT, size: SIZE_BODY, color: MUTED }),
              link('boobalanIT', 'https://github.com/boobalanIT'),
              sep(),
              new TextRun({ text: 'Portfolio: ', font: FONT, size: SIZE_BODY, color: MUTED }),
              link('alanz.netlify.app', 'https://alanz.netlify.app/')
            ]
          }),

          // Professional Summary
          sectionHeading('Professional Summary'),
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 20, after: 25 },
            children: [
              new TextRun({
                text: 'Software Developer with nearly 2 years of professional experience developing responsive web applications, robust REST APIs, and database-driven enterprise solutions. Skilled in Angular, Node.js, NestJS, PHP, PostgreSQL, and MongoDB. Hands-on experience integrating open-source LLMs via Model Context Protocol (MCP) for conversational business intelligence, developing automated PDF reporting systems, and managing deployments on Nginx and Apache servers. Dedicated to writing clean, maintainable code, solving complex business requirements, and optimizing application performance.',
                font: FONT,
                size: SIZE_BODY,
                color: BLACK
              })
            ]
          }),

          // Technical Skills
          sectionHeading('Technical Skills'),
          new Table({
            layout: TableLayoutType.FIXED,
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE, size: 0 },
              bottom: { style: BorderStyle.NONE, size: 0 },
              left: { style: BorderStyle.NONE, size: 0 },
              right: { style: BorderStyle.NONE, size: 0 },
              insideHorizontal: { style: BorderStyle.NONE, size: 0 },
              insideVertical: { style: BorderStyle.NONE, size: 0 }
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } },
                    children: [
                      new Paragraph({
                        spacing: { before: 15, after: 15 },
                        children: [
                          new TextRun({ text: 'Frontend: ', font: FONT, size: SIZE_BODY, bold: true, color: BLACK }),
                          new TextRun({ text: 'Angular, TypeScript, JavaScript, HTML5, CSS3, Bootstrap', font: FONT, size: SIZE_BODY, color: BLACK })
                        ]
                      })
                    ]
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } },
                    children: [
                      new Paragraph({
                        spacing: { before: 15, after: 15 },
                        children: [
                          new TextRun({ text: 'Backend: ', font: FONT, size: SIZE_BODY, bold: true, color: BLACK }),
                          new TextRun({ text: 'Node.js, NestJS, PHP, REST APIs', font: FONT, size: SIZE_BODY, color: BLACK })
                        ]
                      })
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } },
                    children: [
                      new Paragraph({
                        spacing: { before: 15, after: 15 },
                        children: [
                          new TextRun({ text: 'Databases: ', font: FONT, size: SIZE_BODY, bold: true, color: BLACK }),
                          new TextRun({ text: 'PostgreSQL, MongoDB', font: FONT, size: SIZE_BODY, color: BLACK })
                        ]
                      })
                    ]
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } },
                    children: [
                      new Paragraph({
                        spacing: { before: 15, after: 15 },
                        children: [
                          new TextRun({ text: 'Web Servers & Deployment: ', font: FONT, size: SIZE_BODY, bold: true, color: BLACK }),
                          new TextRun({ text: 'Nginx, Apache, Application Deployment', font: FONT, size: SIZE_BODY, color: BLACK })
                        ]
                      })
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    columnSpan: 2,
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } },
                    children: [
                      new Paragraph({
                        spacing: { before: 15, after: 15 },
                        children: [
                          new TextRun({ text: 'Tools & CI/CD: ', font: FONT, size: SIZE_BODY, bold: true, color: BLACK }),
                          new TextRun({ text: 'Git, GitLab CI/CD, Postman, pgAdmin', font: FONT, size: SIZE_BODY, color: BLACK })
                        ]
                      })
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    columnSpan: 2,
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } },
                    children: [
                      new Paragraph({
                        spacing: { before: 15, after: 25 },
                        children: [
                          new TextRun({ text: 'AI & Integration: ', font: FONT, size: SIZE_BODY, bold: true, color: BLACK }),
                          new TextRun({ text: 'Open-source LLM integration, Model Context Protocol (MCP), AI-powered chatbot integration', font: FONT, size: SIZE_BODY, color: BLACK })
                        ]
                      })
                    ]
                  })
                ]
              })
            ]
          }),

          // Professional Experience
          sectionHeading('Professional Experience'),

          companyDateRow('Megam Technologies', 'December 2024 \u2013 Present'),
          roleSubtitle('Software Developer  |  Chennai, India'),

          expSubHeading('Datalytics \u2014 AI-Powered Business Intelligence Platform'),
          bullet('Developed responsive frontend components using Angular and integrated REST APIs built with Node.js and NestJS.'),
          bullet('Integrated an open-source Large Language Model (LLM) with Model Context Protocol (MCP), enabling users to query business data using natural language and receive AI-generated insights.'),
          bullet('Developed interactive dashboards and reporting features to monitor key business metrics, including employee work hours, project progress, and issue tracking.'),
          bullet('Configured Nginx as a reverse proxy to route frontend requests and backend API traffic, and managed GitLab CI/CD pipelines for application deployment.'),

          expSubHeading('Datalytics \u2014 UI Component Library & Embed SDK'),
          bulletWithBold('UI Component Library: ', 'Developed an independent, reusable Angular component library featuring interactive charts (bar, line, area, and pie), KPI widgets, data tables with sorting, filtering, and pagination, and report configuration settings. Maintained independent versioning and published the package to npm.'),
          bulletWithBold('Embed SDK: ', 'Developed a reusable Embed SDK within the main application repository, with a dedicated package.json and independent npm publishing configuration. Enabled third-party applications to embed live analytics dashboards using token-based authentication.'),
          bulletWithBold('Package Management & Integration: ', 'Managed independent versioning and npm releases for both packages, and integrated updated library versions into the main Datalytics SaaS platform.'),

          expSubHeading('LeatherLink TanPro 11.0 \u2014 Tannery Management System (TMS)'),
          bullet('Migrated PHP applications and servers, configured deployment environments, and performed application maintenance, bug fixes, and feature enhancements.'),
          bullet('Developed and maintained core application modules using PHP and PostgreSQL, supporting tannery business operations across 300+ database tables.'),
          bullet('Implemented PostgreSQL schema-per-tenant architecture to maintain separate database schemas for different enterprise clients.'),
          bullet('Implemented and maintained MongoDB-based audit logging and multi-level approval workflows to track business transactions.'),
          bullet('Developed and optimized SQL queries, views, materialized views, functions, and triggers to support application workflows and database operations.'),
          bullet('Implemented automated PDF generation using mPDF for delivery challans, GST invoices, and job cards.'),
          bullet('Configured and maintained Apache and PHP-FPM environments, including server setup, deployment, and environment configuration for multiple client applications.'),
          bullet('Configured and maintained cron jobs to automate scheduled application tasks and background processes.'),


          // Personal Projects
          sectionHeading('Personal Projects'),

          // KYAAT
          new Paragraph({
            spacing: { before: 20, after: 6 },
            children: [
              new TextRun({ text: 'KYAAT \u2014 Gesture-Controlled Mouse ', font: FONT, size: SIZE_BODY, bold: true, color: BLACK }),
              new TextRun({ text: '(Python, OpenCV)', font: FONT, size: SIZE_BODY - 1, italics: true, color: MUTED })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 0, after: 12 },
            children: [
              new TextRun({
                text: 'Developed a computer-vision tool using Python and OpenCV to recognize real-time hand movements and translate gestures into mouse cursor interactions with sub-20ms response latency.',
                font: FONT,
                size: SIZE_BODY,
                color: BLACK
              })
            ]
          }),

          // Personal Portfolio
          new Paragraph({
            spacing: { before: 12, after: 6 },
            children: [
              new TextRun({ text: 'Personal Portfolio ', font: FONT, size: SIZE_BODY, bold: true, color: BLACK }),
              new TextRun({ text: '(Web Development)', font: FONT, size: SIZE_BODY - 1, italics: true, color: MUTED })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 0, after: 8 },
            children: [
              new TextRun({
                text: 'Designed and built a fully responsive personal portfolio website to showcase technical skills, live project demonstrations, and professional experience.',
                font: FONT,
                size: SIZE_BODY,
                color: BLACK
              })
            ]
          }),
          new Paragraph({
            spacing: { before: 0, after: 16 },
            children: [
              new TextRun({ text: 'Available at: ', font: FONT, size: SIZE_BODY, color: MUTED }),
              link('alanz.netlify.app', 'https://alanz.netlify.app/')
            ]
          }),

          // Education
          sectionHeading('Education'),
          new Table({
            layout: TableLayoutType.FIXED,
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE, size: 0 },
              bottom: { style: BorderStyle.NONE, size: 0 },
              left: { style: BorderStyle.NONE, size: 0 },
              right: { style: BorderStyle.NONE, size: 0 },
              insideHorizontal: { style: BorderStyle.NONE, size: 0 },
              insideVertical: { style: BorderStyle.NONE, size: 0 }
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 75, type: WidthType.PERCENTAGE },
                    borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } },
                    children: [
                      new Paragraph({
                        spacing: { before: 15, after: 5 },
                        children: [
                          new TextRun({ text: 'Bachelor of Technology (B.Tech) \u2014 Information Technology', font: FONT, size: SIZE_BODY, bold: true, color: BLACK })
                        ]
                      })
                    ]
                  }),
                  new TableCell({
                    width: { size: 25, type: WidthType.PERCENTAGE },
                    borders: { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 } },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        spacing: { before: 15, after: 5 },
                        children: [
                          new TextRun({ text: '2020 \u2013 2024', font: FONT, size: SIZE_BODY, bold: true, color: DARK })
                        ]
                      })
                    ]
                  })
                ]
              })
            ]
          }),
          new Paragraph({
            spacing: { before: 0, after: 20 },
            children: [
              new TextRun({ text: 'Mohamed Sathak A J College of Engineering', font: FONT, size: SIZE_BODY, color: MUTED }),
              sep(),
              new TextRun({ text: 'Chennai, India', font: FONT, size: SIZE_BODY, color: MUTED }),
              sep(),
              new TextRun({ text: '80% (First Class)', font: FONT, size: SIZE_BODY, color: MUTED })
            ]
          }),

          // Achievements & Certifications
          sectionHeading('Achievements & Certifications'),
          bulletWithBold('The Versatile Rockstar Award,', ' Megam Technologies, 2026 \u2014 certificate and trophy for excellent performance.'),
          bullet('Full Stack Development \u2014 Infosys Springboard (Frontend Frameworks, Backend APIs, & Databases)'),
          bullet('Python Programming \u2014 Spoken Tutorial Project at IIT Bombay')
        ]
      }
    ]
  });

  const buffer = await Packer.toBuffer(doc);
  const outputPath = path.join(__dirname, 'Boobalan_V_Resume.docx');
  fs.writeFileSync(outputPath, buffer);
  console.log('Successfully generated Word resume at:', outputPath);
}

buildDocument().catch(err => {
  console.error('Error generating Word resume:', err);
  process.exit(1);
});
