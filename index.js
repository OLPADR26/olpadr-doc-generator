const express = require('express');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const HTMLModule = require('docxtemplater-html-module');
const { marked } = require('marked');
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

// Turn "Section N: Title" and "Artifact Name: Title" lines into real markdown headings
// (## so they become bold via the HTML <h2> tag, with plain black text forced by CSS below)
function normalizeMarkdown(text) {
  let out = text;

  out = out.replace(/(?:=\s){5,}=?/g, '\n\n---\n\n');
  out = out.replace(/(?:_\s){5,}_?/g, '\n\n---\n\n');

  out = out.replace(/^Section \d+:\s*(.+)$/gm, '## Section: $1');
  out = out.replace(/^Artifact Name:\s*(.+)$/gm, '## $1');

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
    out = out.replace(new RegExp(`^${escaped}$`, 'gm'), `### ${title}`);
  });

  return out;
}

app.post('/generate', async (req, res) => {
  try {
    const { templateUrl, markdownContent, fileName, headerText, dateIssued } = req.body;

    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'doc-'));
    const docxPath = path.join(tempDir, 'output.docx');

    // 1. Convert the raw agent markdown into real HTML (proper tables, headings, bold, lists)
    const fullMarkdown = (headerText ? `${headerText}\n\n---\n\n` : '') + normalizeMarkdown(markdownContent);
    const htmlBody = marked.parse(fullMarkdown);

    // 2. Download the letterhead template
    const templateBuffer = await fetchBuffer(templateUrl);

    // 3. Inject the HTML into the template's {~content} placeholder using the HTML module.
    //    This creates REAL Word tables, real bold headings, real paragraphs — not text guessing.
    const zip = new PizZip(templateBuffer);
    const htmlModule = new HTMLModule({
      // Force headings to render as bold, black, normal-size text (no colored heading style)
      styleSets: {
        h2: { bold: true, color: '000000' },
        h3: { bold: true, color: '000000' }
      }
    });
    const doc = new Docxtemplater(zip, {
      modules: [htmlModule],
      paragraphLoop: true,
      linebreaks: true
    });

    doc.render({ content: htmlBody, date_issued: dateIssued || '' });

    const outputBuffer = doc.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(docxPath, outputBuffer);

    // 4. Convert the finished docx to PDF
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
