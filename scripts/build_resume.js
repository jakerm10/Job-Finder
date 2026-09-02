const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  BorderStyle, PositionalTab, PositionalTabAlignment, PositionalTabLeader,
  PositionalTabPosition, LevelFormat, TabStopType, TabStopPosition,
} = require('docx');

const NAME = 'Jake Moore';
const CONTACT = 'jake.moore-2@colorado.edu | 720-629-2123 | https://www.linkedin.com/in/jakermoore/';
const FONT = 'Times New Roman';
const BODY_SIZE = 21; // 10.5pt
const NAME_SIZE = 32; // 16pt
const CONTACT_SIZE = 18; // 9pt
const HEADING_SIZE = 21; // 10.5pt bold

const EDU = {
  institution: 'University of Colorado Boulder',
  gradDate: 'May 2028',
  degree: 'Bachelor of Science in Computer Science, Minor in Statistics, Engineering Leadership Certificate',
  gpa: 'GPA: 3.578',
};

const TRACKS = {
  development: {
    label: 'Software Engineering / Full-Stack Development',
    coursework: 'Software Development: Full Stack, Computer Algorithms, Computer Systems, Design and Analysis of Database Systems',
    skills: {
      Languages: 'C++, Python, React.js, JavaScript, Node.js, C, HTML, SQL, R-Studio, Arduino',
      Tools: 'Azure DevOps, Docker, Postman, Salesforce, Microsoft Excel, MacOS, Windows OS, Adobe Suite, Google Suite, Asana',
      'Other Skills': 'AI Prompt Engineering, Analytical Problem Solving, Project Management, Data Analysis',
    },
    projects: [
      { name: 'Personal Portfolio Website', when: 'Summer 2026', desc: 'Developed a full-stack software and photography portfolio using React.js, EmailJS, Cloudflare, and Firebase. Owned the project solo end-to-end — native CSS, deployment pipelines, and full account creation/modification, contact-via-email, and image submission functionality. Stack: JavaScript, CSS, HTML, SQL.' },
      { name: 'Gameday Stats', when: 'In Progress', desc: 'Owning a database-backed tool end-to-end, from requirements through schema design and query logic, that lets users selectively pull statistics for individual games, players, and teams. Implementing and testing SQL-driven backend query logic in Python to eliminate manual data-gathering for end users. Stack: Python, SQL.' },
      { name: 'Photo Sorter', when: 'Winter 2025-26', desc: 'Built a file-processing pipeline for camera image files that scans an input folder, parses MacOS "locked" metadata, duplicates matching files, strips the metadata, and moves copies into a target folder.' },
    ],
    employment: [
      { title: 'AI Assisted Researcher', org: 'Brian Hewlett LLC', dates: 'Summer 2026', bullets: [
        'Designed and owned AI-assisted prompt workflows end-to-end, enforcing consistent, reliable output across employees and cutting review time from 5 hours to 1',
        'Leveraged AI to synthesize marketing techniques and data points into content plans',
        'Structured and designed reports to efficiently guide new content strategies',
      ]},
      { title: 'Gold Level Leader – Creative Media Intern', org: 'University of Colorado Boulder', dates: 'August 2024 – Present', bullets: [
        'Captured and edited photo/video for 100+ university athletic competitions, practices, and media days',
        'Delivered images in a fast-paced gametime environment generating over 1 million views',
        'Mentored newer interns on game-day operations and efficient editing techniques',
      ]},
    ],
    leadership: [
      { title: 'Junior Counselor (JC)', org: 'Young Rotary Youth Leadership Awards Conference', dates: 'October 2022 – June 2024', bullets: [
        'Facilitated activities/discussions with middle school students, promoting leadership skill development',
        'Managed conflicts and unforeseen circumstances with poise',
      ]},
      { title: 'Patrol Leader', org: 'Troop 329', dates: 'February 2019 – August 2019', bullets: [
        'Led a patrol of 30 Scouts through weekly meetings, rank progression, and event planning',
      ]},
    ],
    honors: ["Dean's List, University of Colorado Boulder, May 2025", 'Salutatorian, Golden High School', 'Eagle Scout, Scouting of America, November 2021'],
  },
  database: {
    label: 'Data / Database Engineering',
    coursework: 'Design and Analysis of Database Systems, Computer Algorithms, Software Development: Full Stack',
    skills: {
      Languages: 'SQL, Python, R-Studio, C++, JavaScript, Node.js, C, HTML, Arduino',
      Tools: 'Microsoft Excel, Postman, Docker, Azure DevOps, Salesforce, MacOS, Windows OS, Google Suite',
      'Other Skills': 'Data Analysis, Statistical Analysis, Analytical Problem Solving, AI Prompt Engineering',
    },
    projects: [
      { name: 'Gameday Stats', when: 'In Progress', desc: 'Designing and documenting the schema and query requirements for a database-oriented tool allowing users to selectively query statistics from individual games, players, and teams. Removes the need for users to manually acquire data by leveraging SQL techniques to select and filter stats across a relational database. Built with Python and SQL.' },
      { name: 'Personal Portfolio Website', when: 'Summer 2026', desc: 'Developed a full-stack site using React.js, EmailJS, Cloudflare, and Firebase, including SQL-backed account creation/modification and data persistence. Handled deployment pipelines and user event/data flows end-to-end.' },
      { name: 'Photo Sorter', when: 'Winter 2025-26', desc: 'Built a file-processing script that parses and rewrites MacOS metadata, applying structured data transformations across a batch file set.' },
    ],
    employment: [
      { title: 'AI Assisted Researcher', org: 'Brian Hewlett LLC', dates: 'Summer 2026', bullets: [
        'Cleaned and synthesized marketing data points into structured, digestible content plans using AI-assisted analysis',
        'Structured and designed reports to guide data-informed content strategy',
        'Designed prompts to automate consistency across employees, cutting review time from 5 hours to 1',
      ]},
      { title: 'Gold Level Leader – Creative Media Intern', org: 'University of Colorado Boulder', dates: 'August 2024 – Present', bullets: [
        'Captured and edited photo/video for 100+ university athletic competitions, practices, and media days',
        'Delivered content generating over 1 million views in a fast-paced, deadline-driven environment',
      ]},
    ],
    leadership: [
      { title: 'Junior Counselor (JC)', org: 'Young Rotary Youth Leadership Awards Conference', dates: 'October 2022 – June 2024', bullets: [
        'Measured learning progress to adapt and adjust training methods',
      ]},
      { title: 'Patrol Leader', org: 'Troop 329', dates: 'February 2019 – August 2019', bullets: [
        'Facilitated and oversaw troop communication and planning with senior leadership',
      ]},
    ],
    honors: ["Dean's List, University of Colorado Boulder, May 2025", 'Salutatorian, Golden High School', 'Eagle Scout, Scouting of America, November 2021'],
  },
  'sports-tech': {
    label: 'Sports Technology / Sports Data & Analytics',
    coursework: 'Design and Analysis of Database Systems, Computer Algorithms, Software Development: Full Stack',
    skills: {
      Languages: 'Python, SQL, R-Studio, JavaScript, React.js, Node.js, C++',
      Tools: 'MacOS, Windows OS, Docker, Google Suite, Asana',
      'Other Skills': 'Sports Analytics & Data Visualization, Statistical Analysis, Data Analysis, AI Prompt Engineering',
    },
    projects: [
      { name: 'Gameday Stats', when: 'In Progress', desc: 'Designing a database-oriented sports statistics tool allowing users to selectively find and pull accurate statistics for individual games, players, and teams via SQL-driven data pipelines. Built with Python and SQL to remove the need for manual sports data gathering.' },
      { name: 'Personal Portfolio Website', when: 'Summer 2026', desc: 'Developed a full-stack site using React.js, EmailJS, Cloudflare, and Firebase, with a SQL-backed account system. JavaScript, CSS, HTML, SQL.' },
      { name: 'Photo Sorter', when: 'Winter 2025-26', desc: 'Developed a file-processing script that scans an input folder, parses MacOS "locked" metadata, creates copies, strips metadata, and organizes output.' },
    ],
    employment: [
      { title: 'AI Assisted Researcher', org: 'Brian Hewlett LLC', dates: 'Summer 2026', bullets: [
        'Designed prompts to automate consistency across employees, cutting review time from 5 hours to 1',
        'Leveraged AI to synthesize marketing techniques and data points into content plans',
        'Structured and designed reports to efficiently guide new content strategies',
      ]},
      { title: 'Gold Level Leader – Creative Media Intern', org: 'University of Colorado Boulder', dates: 'August 2024 – Present', bullets: [
        'Embedded within CU Boulder Athletics\' gameday operations for two years, gaining first-hand exposure to how a college athletics department runs its live sports and stats operations',
      ]},
    ],
    leadership: [
      { title: 'Junior Counselor (JC)', org: 'Young Rotary Youth Leadership Awards Conference', dates: 'October 2022 – June 2024', bullets: [
        'Facilitated activities/discussions with middle school students, promoting leadership skill development',
      ]},
      { title: 'Patrol Leader', org: 'Troop 329', dates: 'February 2019 – August 2019', bullets: [
        'Led a patrol of 30 Scouts through weekly meetings, rank progression, merit badges, and event planning',
      ]},
    ],
    honors: ["Dean's List, University of Colorado Boulder, May 2025", 'Salutatorian, Golden High School', 'Eagle Scout, Scouting of America, November 2021'],
    otherInterests: 'Sports analytics and data visualization, personal health and fitness (CU Club Track and Field), technology development and integration',
  },
  leadership: {
    label: 'Leadership-Track Internship',
    coursework: 'Engineering Leadership: Complex Leadership Challenges, Software Development: Full Stack, Computer Systems',
    skills: {
      Languages: 'SQL, JavaScript, React.js',
      Tools: 'Azure DevOps, Salesforce, Asana, Google Suite, Microsoft Excel',
      'Other Skills': 'Project Management, Leadership, Networking, AI Prompt Engineering, Analytical Problem Solving',
    },
    projects: [
      { name: 'Gameday Stats', when: 'In Progress', desc: 'Owning a real-world, self-directed database-oriented tool end-to-end — from requirements through schema design — allowing users to selectively find statistics from individual games, players, and teams. Built with Python and SQL.' },
      { name: 'Personal Portfolio Website', when: 'Summer 2026', desc: 'Developed a full-stack software and photography portfolio using React.js, EmailJS, Cloudflare, and Firebase.' },
    ],
    employment: [
      { title: 'AI Assisted Researcher', org: 'Brian Hewlett LLC', dates: 'Summer 2026', bullets: [
        'Structured and designed reports and a repeatable practical/technical process to efficiently guide new content strategies',
        'Designed prompts to automate consistency across employees, cutting review time from 5 hours to 1',
        'Leveraged AI to synthesize marketing techniques and data points into content plans',
      ]},
      { title: 'Gold Level Leader – Creative Media Intern', org: 'University of Colorado Boulder', dates: 'August 2024 – Present', bullets: [
        'Mentored newer interns on game-day operations, efficient editing techniques, and shooting skills',
        'Captured and edited photo/video for over 100 university athletic competitions, practices, and media days',
        'Edited and delivered images in a fast-paced gametime environment generating over 1 million views',
      ]},
    ],
    leadership: [
      { title: 'Junior Counselor (JC)', org: 'Young Rotary Youth Leadership Awards Conference', dates: 'October 2022 – June 2024', bullets: [
        'Facilitated activities/discussions with middle school students, promoting leadership skill development',
        'Managed conflicts and unforeseen circumstances with poise to preserve a positive learning environment',
        'Measured learning progress to adapt and adjust training methods',
        'Supervised living conditions of conferees to prioritize safety and cleanliness',
      ]},
      { title: 'Patrol Leader', org: 'Troop 329', dates: 'February 2019 – August 2019', bullets: [
        'Led a patrol of 30 Scouts through weekly meetings, rank progression, merit badges, and event planning',
        'Mentored patrol members in leadership and character development',
        'Facilitated and oversaw cross-team communication and planning with senior troop leadership',
      ]},
    ],
    honors: ["Dean's List, University of Colorado Boulder, May 2025", 'Salutatorian, Golden High School', 'Eagle Scout, Scouting of America, November 2021'],
  },
};

function bodyRun(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: BODY_SIZE, ...opts });
}

function sectionHeading(text) {
  return new Paragraph({
    spacing: { before: 180, after: 60 },
    border: { bottom: { style: BorderStyle.NONE } },
    children: [new TextRun({ text, font: FONT, size: HEADING_SIZE, bold: true, underline: {}, allCaps: true })],
  });
}

function bulletParagraph(text) {
  return new Paragraph({
    spacing: { after: 20 },
    indent: { left: 270, hanging: 180 },
    children: [bodyRun('- '), bodyRun(text)],
  });
}

function indentedLine(label, value) {
  return new Paragraph({
    indent: { left: 270, hanging: 180 },
    spacing: { after: 20 },
    children: [bodyRun('- '), bodyRun(`${label}: `, { bold: true }), bodyRun(value)],
  });
}

function titleDateLine(title, dates) {
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: { after: 0 },
    children: [
      bodyRun(title, { bold: true }),
      new TextRun({ text: `\t${dates}`, font: FONT, size: BODY_SIZE }),
    ],
  });
}

function orgLine(org) {
  return new Paragraph({ spacing: { after: 40 }, children: [bodyRun(org)] });
}

function buildDoc(track) {
  const t = TRACKS[track];
  const children = [];

  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 40 },
    children: [new TextRun({ text: NAME, font: FONT, size: NAME_SIZE, bold: true })],
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, space: 6, color: '000000' } },
    children: [new TextRun({ text: CONTACT, font: FONT, size: CONTACT_SIZE })],
  }));

  // EDUCATION
  children.push(sectionHeading('EDUCATION'));
  children.push(new Paragraph({
    spacing: { after: 20 },
    children: [
      bodyRun(EDU.institution, { bold: true }),
      bodyRun(`, ${EDU.gradDate}, `),
      bodyRun(EDU.degree, { bold: true }),
    ],
  }));
  children.push(indentedLine('GPA', '3.578'));
  children.push(indentedLine('Relevant Coursework', t.coursework));

  // TECHNICAL SKILLS
  children.push(sectionHeading('TECHNICAL SKILLS'));
  for (const [label, value] of Object.entries(t.skills)) {
    children.push(new Paragraph({
      spacing: { after: 20 },
      children: [bodyRun(`${label}: `, { bold: true }), bodyRun(value)],
    }));
  }

  // PROJECTS
  children.push(sectionHeading('PROJECTS'));
  for (const p of t.projects) {
    children.push(new Paragraph({
      spacing: { after: 60 },
      children: [
        bodyRun(`${p.name}`, { bold: true }),
        bodyRun(`, ${p.when}: `),
        bodyRun(p.desc),
      ],
    }));
  }

  // EMPLOYMENT HISTORY
  if (t.employment && t.employment.length) {
    children.push(sectionHeading('EMPLOYMENT HISTORY'));
    for (const e of t.employment) {
      children.push(titleDateLine(e.title, e.dates));
      children.push(orgLine(e.org));
      for (const b of e.bullets) children.push(bulletParagraph(b));
    }
  }

  // LEADERSHIP EXPERIENCE
  if (t.leadership && t.leadership.length) {
    children.push(sectionHeading('LEADERSHIP EXPERIENCE'));
    for (const e of t.leadership) {
      children.push(titleDateLine(e.title, e.dates));
      children.push(orgLine(e.org));
      for (const b of e.bullets) children.push(bulletParagraph(b));
    }
  }

  // HONORS & AWARDS
  if (t.honors && t.honors.length) {
    children.push(sectionHeading('HONORS & AWARDS'));
    for (const h of t.honors) children.push(bulletParagraph(h));
  }

  // OTHER INTERESTS AND INVOLVEMENTS
  if (t.otherInterests) {
    children.push(sectionHeading('OTHER INTERESTS AND INVOLVEMENTS'));
    children.push(new Paragraph({ spacing: { after: 20 }, children: [bodyRun(t.otherInterests)] }));
  }

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 720, bottom: 720, left: 720, right: 720 },
        },
      },
      children,
    }],
  });
}

async function main() {
  const args = process.argv.slice(2);
  const tracks = args[0] === 'all' || args.length === 0 ? Object.keys(TRACKS) : args;
  const outDir = path.join(__dirname, '..', 'resumes');
  for (const track of tracks) {
    if (!TRACKS[track]) { console.error('Unknown track', track); continue; }
    const doc = buildDoc(track);
    const buf = await Packer.toBuffer(doc);
    const outPath = path.join(outDir, `${track}.docx`);
    fs.writeFileSync(outPath, buf);
    console.log('Wrote', outPath);
  }
}

main();
