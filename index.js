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

// Turn "Section N: Title" and "Artifact Name: Title" lines into real bold headings
function boldTitles(text) {
  let out = text;
  // Section headers -> Heading 2
  out = out.replace(/^Section \d+:\s*(.+)$/gm, '## Section: $1');
  // Artifact block titles -> Heading 2
  out = out.replace(/^Artifact Name:\s*(.+)$/gm, '## $1');
  // Common sub-section titles used across artifacts -> Heading 3
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

// Calculate proportional column widths for a markdown table block
function balanceTableColumns(tableBlock) {
  const lines = tableBlock.split('\n').filter(l => l.trim().startsWith('|'));
  if (lines.length < 2) return tableBlock;

  const headerCells = lines[0].split('|').slice(1, -1).map(c => c.trim());
  const colCount = headerCells.length;
  const maxLen = new Array(colCount).fill(3);

  lines.forEach((line, idx) => {
    if (idx === 1) return; // skip separator row
    const cells = line.split('|').slice(1, -1);
    cells.forEach((cell, i) => {
      if (i < colCount) {
        const len = cell.trim().length;
        if (len > maxLen[i]) maxLen[i] = len;
      }
    });
  });

  // Cap each column's proportional weight so no single column dominates completely
  const capped = maxLen.map(l => Math.min(Math.max(l, 6), 40));
  const newSeparator = '|' + capped.map(l => '-'.repeat(l)).join('|') + '|';

  lines[1] = newSeparator;
  return lines.join('\n');
}

function balanceAllTables(text) {
  const blocks = text.split(/\n\n(?=\|)/);
  return blocks.map(block => {
    if (block.trim().startsWith('|')) {
      return balanceTableColumns(block);
    }
    return block;
  }).join('\n\n');
}

function cleanMarkdown(text) {
  let cleaned = text;
  cleaned = cleaned.replace(/(?:=\s){5,}=?/g, '\n\n---\n\n');
  cleaned = cleaned.replace(/(?:_\s){5,}_?/g, '\n\n---\n\n');
  cleaned = cleaned.replace(/([^\n|])\s\|(?=[^|]*\|)/g, '$1\n|');
  cleaned = cleaned.replace(/([^\n])\n(\|)/g, '$1\n\n$2');
  cleaned = cleaned.replace(/([^\n])\s(\d+\.\s)/g, '$1\n\n$2');
  cleaned = cleaned.replace(/([^\n])\s([-*]\s)/g, '$1\n\n$2');
  cleaned = cleaned.replace(/(\|.*\|)\n([^\n|])/g, '$1\n\n$2');
  cleaned = cleaned.replace(/(^#{1,6}\s.*$)\n([^\n#])/gm, '$1\n\n$2');
  cleaned = cleaned.replace(/\n{3,}/g, '\n\n');

  cleaned = boldTitles(cleaned);
  cleaned = balanceAllTables(cleaned);

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
