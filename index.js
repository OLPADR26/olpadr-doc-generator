const express = require('express');
const https = require('https');
const { exec } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const app = express();

app.use(express.json({ limit: '50mb' }));
app.use(require('cors')());

function fetchBuffer(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetchBuffer(res.headers.location).then(resolve).catch(reject);
      }
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

// Force EVERY single line break to become a real paragraph break,
// EXCEPT inside table blocks (where single line breaks between rows are required)
// and EXCEPT between consecutive list items.
function forceParagraphBreaks(text) {
  const lines = text.split('\n');
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const nextLine = lines[i + 1] || '';
    out.push(line);

    const thisIsTableRow = /^\s*\|/.test(line);
    const nextIsTableRow = /^\s*\|/.test(nextLine);
    const thisIsListItem = /^\s*([-*]|\d+\.)\s/.test(line);
    const nextIsListItem = /^\s*([-*]|\d+\.)\s/.test(nextLine);
    const lineIsBlank = line.trim() === '';
    const nextIsBlank = nextLine.trim() === '';

    // Keep table rows tight against each other (no blank line inserted between them)
    if (thisIsTableRow && nextIsTableRow) continue;
    // Keep list items tight against each other
    if (thisIsListItem && nextIsListItem) continue;
    // Don't double up existing blank lines
    if (lineIsBlank || nextIsBlank) continue;
    // Don't break right before a table starts (that blank line is inserted separately below)
    if (nextIsTableRow) { out.push(''); continue; }

    // Otherwise: force a blank line so this becomes its own paragraph
    out.push('');
  }
  return out.join('\n');
}

// Turn "Section N: Title" and "Artifact Name: Title" into forced BOLD PLAIN TEXT
// (not a heading style) so it never inherits the template's colored Heading style.
function boldTitles(text) {
  let out = text;
  out = out.replace(/^Section \d+:\s*(.+)$/gm, (match, title) => `**Section: ${title}**`);
  out = out.replace(/^Artifact Name:\s*(.+)$/gm, (match, title) => `**${title}**`);

  const subTitles = [
    'Primary Accountability Context', 'Leading Asset Statement', 'Terminal Gap', 'Causal Anchor',
    'Active Constraints', 'Assumptions', 'Executive Directive for Causal Mapping',
    'Causal Spine Description', 'Logic Matrix', 'Visual Placeholder', 'Triage Brief',
    'Full Minimum Viable Data Field List', 'Executive Recommendation', 'Programme Strengths',
    'The Leakage Constraint', 'Baseline Gate Declaration', 'What-If Scenario Matrix',
    'What-If Simulation Matrix', 'Scenario Notes', 'Simulation Pivot', 'What to Stop Doing',
    'What to Double Down On', 'Constraint Workaround', 'Cost Per Beneficiary',
    'Executive Architecture Directive', 'Sprint Pulse Report', 'Evidence Gap Alert',
    'Active Constraint Check', 'Evidence Sensor', 'Sprint Gate Decision',
    'Identified Chaos', 'Failure Type', 'Clinical Diagnostics', 'Evidence Ledger',
    'Artifact Readiness', 'Causal Logic', 'Constraints', 'Auditor\u2019s Verdict',
    'Reconciliation Note', 'Singular Architectural Move', 'Asset Statement',
    'Terminal Constraint', 'Strategic Pivot', 'Current Evidence Position',
    'Outcome Commitments', 'Evidence Gap', 'Data System Status', 'Minimum Viable Evidence Architecture',
    'Indicator Architecture', 'Claim Control', 'Cost Position', 'Protocol Pathway',
    'Post-Protocol Position', 'Decision Artifact', 'Conditions', 'Next-Season Plan',
    'Grant-Ready Decision', 'Singular Strategic Recommendation'
  ];
  subTitles.forEach(title => {
    const escaped = title.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    out = out.replace(new RegExp(`^${escaped}$`, 'gm'), `**${title}**`);
  });
  return out;
}

function cleanMarkdown(text) {
  let cleaned = text;

  cleaned = cleaned.replace(/(?:=\s){5,}=?/g, '\n\n---\n\n');
  cleaned = cleaned.replace(/(?:_\s){5,}_?/g, '\n\n---\n\n');

  cleaned = boldTitles(cleaned);
  cleaned = forceParagraphBreaks(cleaned);

  // Final safety cleanup: collapse any 3+ blank lines down to exactly 2 line breaks
  cleaned = cleaned.replace(/\n{3,}/g, '\n\n');

  return cleaned;
}

app.post('/generate', async (req, res) => {
  try {
    const { templateUrl, markdownContent, fileName, headerText } = req.body;

    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'doc-'));
    const referenceDocPath = path.join(tempDir, 'reference.docx');
    const markdownPath = path.join(tempDir, 'content.md');
    const docxPath = path.join(tempDir, 'output.docx');

    const templateBuffer = await fetchBuffer(templateUrl);
    fs.writeFileSync(referenceDocPath, templateBuffer);

    const fullMarkdown = (headerText ? `${headerText}\n\n---\n\n` : '') + cleanMarkdown(markdownContent);
    fs.writeFileSync(markdownPath, fullMarkdown, 'utf8');

    await new Promise((resolve, reject) => {
      exec(
        `pandoc "${markdownPath}" -o "${docxPath}" --reference-doc="${referenceDocPath}" -f markdown+pipe_tables --standalone`,
        (error) => error ? reject(error) : resolve()
      );
    });

    await new Promise((resolve, reject) => {
      exec(
        `libreoffice --headless -env:UserInstallation=file://${tempDir}/loconfig --convert-to pdf --outdir "${tempDir}" "${docxPath}"`,
        { timeout: 60000 },
        (error, stdout, stderr) => {
          if (error) {
            console.error('LibreOffice error:', error);
            console.error('LibreOffice stderr:', stderr);
            return reject(error);
          }
          resolve();
        }
      );
    });

    const pdfPath = path.join(tempDir, 'output.pdf');
    if (!fs.existsSync(pdfPath)) {
      throw new Error('PDF conversion failed — no output file was created');
    }
    const pdfBuffer = fs.readFileSync(pdfPath);
    if (pdfBuffer.length < 100) {
      throw new Error('PDF conversion produced an empty or invalid file');
    }
    const pdfFileName = (fileName || 'report.docx').replace(/\.docx$/i, '.pdf');
    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="${pdfFileName}"`,
      'Content-Length': pdfBuffer.length
    });
    res.send(pdfBuffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Running on port ${PORT}`));
