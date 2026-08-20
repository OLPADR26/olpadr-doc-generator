const express = require('express');
const https = require('https');
const { execFile } = require('child_process');
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

function runCommand(cmd, args, opts) {
  return new Promise((resolve) => {
    execFile(cmd, args, { timeout: 60000, ...opts }, (error, stdout, stderr) => {
      resolve({ error, stdout: stdout || '', stderr: stderr || '' });
    });
  });
}

// Force EVERY line to be its own paragraph, except inside table rows and list items
// (which must stay tight together for Pandoc to recognise them as tables/lists).
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

    if (thisIsTableRow && nextIsTableRow) continue;
    if (thisIsListItem && nextIsListItem) continue;
    if (lineIsBlank || nextIsBlank) continue;
    if (nextIsTableRow) { out.push(''); continue; }

    out.push('');
  }
  return out.join('\n');
}

// Turn "Section N: Title" and "Artifact Name: Title" into forced PLAIN BOLD text
// (never a heading style) so it never inherits the template's colored Heading style.
function boldTitles(text) {
  let out = text;
  out = out.replace(/^Section \d+:\s*(.+)$/gm, (m, title) => `**Section: ${title}**`);
  out = out.replace(/^Artifact Name:\s*(.+)$/gm, (m, title) => `**${title}**`);

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
  cleaned = cleaned.replace(/\n{3,}/g, '\n\n');
  return cleaned;
}

app.post('/generate', async (req, res) => {
  let tempDir;
  try {
    const { templateUrl, markdownContent, fileName, headerText } = req.body;

    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'doc-'));
    const referenceDocPath = path.join(tempDir, 'reference.docx');
    const markdownPath = path.join(tempDir, 'content.md');
    const docxPath = path.join(tempDir, 'output.docx');

    const templateBuffer = await fetchBuffer(templateUrl);
    fs.writeFileSync(referenceDocPath, templateBuffer);

    const fullMarkdown = (headerText ? `${headerText}\n\n---\n\n` : '') + cleanMarkdown(markdownContent);
    fs.writeFileSync(markdownPath, fullMarkdown, 'utf8');

    // Single command: markdown -> docx, styled using your letterhead (logo, header, footer)
    const step1 = await runCommand('pandoc', [
      markdownPath,
      '-o', docxPath,
      '--reference-doc', referenceDocPath,
      '-f', 'markdown+pipe_tables',
      '--standalone'
    ]);
    console.log('[pandoc] stdout:', step1.stdout);
    console.log('[pandoc] stderr:', step1.stderr);

    if (step1.error || !fs.existsSync(docxPath)) {
      throw new Error(
        `Pandoc conversion failed. exec_error=${step1.error ? step1.error.message : 'none'} | stderr=${step1.stderr.slice(0, 500)}`
      );
    }

    // docx -> pdf
    const step2 = await runCommand('libreoffice', [
      '--headless',
      `-env:UserInstallation=file://${tempDir}/loconfig`,
      '--convert-to', 'pdf',
      '--outdir', tempDir,
      docxPath
    ]);
    console.log('[docx->pdf] stdout:', step2.stdout);
    console.log('[docx->pdf] stderr:', step2.stderr);

    const pdfPath = path.join(tempDir, 'output.pdf');
    if (step2.error || !fs.existsSync(pdfPath)) {
      throw new Error(
        `PDF conversion failed. exec_error=${step2.error ? step2.error.message : 'none'} | stderr=${step2.stderr.slice(0, 500)}`
      );
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
