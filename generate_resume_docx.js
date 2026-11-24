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
const SP_BEFORE_SECTION = 100;
const SP_AFTER_SECTION  = 40;
const SP_BULLET         = 25;

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
    spacing: { before: 60, after: 20 },
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
              top: convertInchesToTwip(0.4),
              bottom: convertInchesToTwip(0.4),
              left: convertInchesToTwip(0.55),
              right: convertInchesToTwip(0.55)
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
                text: 'BOOBALAN V.',
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
              link('Boobalan V.', 'https://www.linkedin.com/in/boobalan001/'),
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
            spacing: { before: 25, after: 30 },
            children: [
              new TextRun({
                text: 'Software Developer with nearly 2 years of professional experience developing and maintaining web applications, working with APIs and databases, and supporting application deployments. Experienced in Angular, Node.js, NestJS, PHP, PostgreSQL, and MongoDB. Contributed to an AI-powered business intelligence platform featuring conversational data analysis, Model Context Protocol (MCP) tools, and user-requested report and dashboard generation. Hands-on experience with SQL database objects, Nginx, Apache, and GitLab CI/CD.',
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
                          new TextRun({ text: 'Angular, HTML5, CSS3, JavaScript, Bootstrap', font: FONT, size: SIZE_BODY, color: BLACK })
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
                          new TextRun({ text: 'Node.js, NestJS, Express.js, PHP, REST APIs', font: FONT, size: SIZE_BODY, color: BLACK })
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
                          new TextRun({ text: 'PostgreSQL, MongoDB, MySQL', font: FONT, size: SIZE_BODY, color: BLACK })
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

          expSubHeading('Datalytics \u2014 Business Intelligence & Analytics Platform'),
          bullet('Developed and maintained features for a business intelligence platform, including MCP server tools that enable AI-assisted interaction with platform data and functionality.'),
          bullet('Integrated an open-source large language model (LLM) to support natural-language questions about business data and provide responses based on analytics information.'),
          bullet('Implemented functionality to generate reports and dashboards from user requests, covering operational metrics such as working hours, time spent, and issues created or closed.'),
          bullet('Supported comparative analysis of employee work activity and time allocation, alongside bug fixing, feature enhancements, and application maintenance.'),

          expSubHeading('LeatherLink TanPro 11.0 (TMS) \u2014 Multi-Tenant Industrial ERP'),
          bullet('Developed and maintained features for an enterprise industrial ERP (300+ tables, ~200k+ LOC) using PHP, PostgreSQL (schema-per-tenant isolation), and MongoDB for audit logs and approval workflows.'),
          bullet('Created and modified SQL queries, views, materialized views, functions, and triggers to support business requirements; troubleshot application and database issues and implemented fixes.'),

          expSubHeading('Deployment & Development Operations'),
          bullet('Configured Nginx as a reverse proxy for Datalytics (Angular & Node.js/NestJS) and Apache HTTP Server for LeatherLink ERP (.htaccess modules, virtual hosts); supported server configuration and deployments.'),
          bullet('Worked with GitLab CI/CD pipelines to support application build and deployment workflows; used Git for version control and Postman for API testing and troubleshooting.'),

          companyDateRow('Shiash Info Solutions Private Limited', 'March 2023 \u2013 May 2023'),
          roleSubtitle('Junior Full-Stack Developer  |  Chennai, India'),
          bullet('Contributed to full-stack web development tasks using HTML, CSS, and JavaScript; assisted with feature implementation and debugging within an established software development team and workflow.'),

          // Personal Projects
          sectionHeading('Personal Projects'),

          // KYAAT
          new Paragraph({
            spacing: { before: 25, after: 8 },
            children: [
              new TextRun({ text: 'KYAAT \u2014 Gesture-Controlled Mouse ', font: FONT, size: SIZE_BODY, bold: true, color: BLACK }),
              new TextRun({ text: '(Python, OpenCV)', font: FONT, size: SIZE_BODY - 1, italics: true, color: MUTED })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 0, after: 15 },
            children: [
              new TextRun({
                text: 'Developed a computer-vision tool using OpenCV to recognize real-time hand movements and translate gestures into mouse cursor interactions with sub-20ms response latency.',
                font: FONT,
                size: SIZE_BODY,
                color: BLACK
              })
            ]
          }),

          // Personal Portfolio
          new Paragraph({
            spacing: { before: 15, after: 8 },
            children: [
              new TextRun({ text: 'Personal Portfolio ', font: FONT, size: SIZE_BODY, bold: true, color: BLACK }),
              new TextRun({ text: '(Web Development)', font: FONT, size: SIZE_BODY - 1, italics: true, color: MUTED })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 0, after: 12 },
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
            spacing: { before: 0, after: 20 },
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
          bullet('Python Programming \u2014 Spoken Tutorial / CodeTantra'),
          bullet('Full Stack Development \u2014 Infosys Springboard')
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
