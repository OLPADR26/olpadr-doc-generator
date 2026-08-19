const express = require('express');
const DocxMerger = require('docx-merger');
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

function execPromise(cmd, opts) {
  return new Promise((resolve, reject) => {
    exec(cmd, opts, (error, stdout, stderr) => {
      if (error) {
        console.error('Command failed:', cmd);
        console.error(stderr);
        return reject(error);
      }
      resolve(stdout);
    });
  });
}

// Turn "Section N: Title" and "Artifact Name: Title" lines into real markdown headings
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
  let tempDir;
  try {
    const { templateUrl, markdownContent, fileName, headerText, dateIssued } = req.body;
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'doc-'));

    // 1. Markdown -> HTML
    const fullMarkdown = (headerText ? `${headerText}\n\n---\n\n` : '') + normalizeMarkdown(markdownContent);
    const htmlBody = marked.parse(fullMarkdown);
    const fullHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
      body { font-family: Calibri, Arial, sans-serif; color: #000000; font-size: 11pt; }
      h2, h3 { color: #000000; font-weight: bold; }
      table { border-collapse: collapse; width: 100%; margin: 8px 0; }
      table, th, td { border: 1px solid #444444; padding: 4px 8px; }
    </style></head><body>${htmlBody}</body></html>`;

    const htmlPath = path.join(tempDir, 'content.html');
    fs.writeFileSync(htmlPath, fullHtml, 'utf8');

    // 2. HTML -> DOCX via LibreOffice (real Word tables/bold/lists, no paid module needed)
    await execPromise(
      `libreoffice --headless -env:UserInstallation=file://${tempDir}/loconfig1 --convert-to docx --outdir "${tempDir}" "${htmlPath}"`,
      { timeout: 60000 }
    );
    const contentDocxPath = path.join(tempDir, 'content.docx');
    if (!fs.existsSync(contentDocxPath)) {
      throw new Error('HTML to DOCX conversion failed — no content.docx was produced');
    }

    // 3. Download the letterhead template.
    //    NOTE: the template no longer needs a {~content} placeholder — its body
    //    can be empty (or just a cover section). The generated content is appended
    //    after it, and the template's headers/footers/styles/margins are kept.
    const templateBuffer = await fetchBuffer(templateUrl);
    const templatePath = path.join(tempDir, 'template.docx');
    fs.writeFileSync(templatePath, templateBuffer);

    // 4. Merge template + generated content
    const templateBin = fs.readFileSync(templatePath, 'binary');
    const contentBin = fs.readFileSync(contentDocxPath, 'binary');
    const merger = new DocxMerger({}, [templateBin, contentBin]);

    const mergedPath = path.join(tempDir, 'output.docx');
    await new Promise((resolve, reject) => {
      try {
        merger.save('nodebuffer', (data) => {
          fs.writeFileSync(mergedPath, data);
          resolve();
        });
      } catch (e) {
        reject(e);
      }
    });

    // 5. Convert merged docx -> PDF
    await execPromise(
      `libreoffice --headless -env:UserInstallation=file://${tempDir}/loconfig2 --convert-to pdf --outdir "${tempDir}" "${mergedPath}"`,
      { timeout: 60000 }
    );

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
  } finally {
    if (tempDir) {
      fs.rm(tempDir, { recursive: true, force: true }, () => {});
    }
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Running on port ${PORT}`));
