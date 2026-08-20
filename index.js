const express = require('express');
const PizZip = require('pizzip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const { marked } = require('marked');
// Preserve single line breaks (e.g. "Source Step: X\nProtocol: Y") as <br> instead
// of collapsing them into one run-on line — confirmed this doesn't affect table
// parsing (tables still need a real blank line before/after, which they have).
marked.setOptions({ breaks: true });
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

// Runs a command via execFile (array args, no shell string-interpolation/quoting
// bugs) and ALWAYS returns stdout/stderr, even on failure, so callers can log or
// surface the real LibreOffice error instead of a generic message.
function runCommand(cmd, args, opts) {
  return new Promise((resolve) => {
    execFile(cmd, args, { timeout: 60000, ...opts }, (error, stdout, stderr) => {
      resolve({ error, stdout: stdout || '', stderr: stderr || '' });
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

// ---------- docx XML merge helpers (non-namespace-aware: match literal "w:" prefixed tags/attrs) ----------

function parseXml(str) {
  return new DOMParser().parseFromString(str, 'text/xml');
}
function serializeXml(doc) {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + new XMLSerializer().serializeToString(doc);
}
function directChildren(el, tagName) {
  const out = [];
  for (let i = 0; i < el.childNodes.length; i++) {
    const c = el.childNodes[i];
    if (c.nodeType === 1 && (!tagName || c.tagName === tagName)) out.push(c);
  }
  return out;
}
function allText(doc) {
  return Array.from(doc.getElementsByTagName('w:t'));
}

// Merges generated `contentDocxBuffer` (body only) into `templateDocxBuffer`
// (which supplies headers/footers/styles/sectPr), remapping numbering IDs so
// bullet/numbered lists from the content don't collide with the template's own,
// and copying over any paragraph/character styles the content uses that the
// template doesn't already define. Returns a Buffer of the merged .docx.
function mergeDocx(templateDocxBuffer, contentDocxBuffer, dateIssuedText) {
  const templateZip = new PizZip(templateDocxBuffer);
  const contentZip = new PizZip(contentDocxBuffer);

  const templateDoc = parseXml(templateZip.file('word/document.xml').asText());
  const contentDoc = parseXml(contentZip.file('word/document.xml').asText());

  const templateBody = templateDoc.getElementsByTagName('w:body')[0];
  const contentBody = contentDoc.getElementsByTagName('w:body')[0];
  if (!templateBody) throw new Error('Template docx has no <w:body> — is it a valid Word file?');
  if (!contentBody) throw new Error('Generated content docx has no <w:body>');

  const templateSectPr = directChildren(templateBody, 'w:sectPr')[0] || null;
  const contentChildren = directChildren(contentBody).filter(c => c.tagName !== 'w:sectPr');

  // 1. Replace the letterhead's date placeholder text, if present, with the real date.
  if (dateIssuedText) {
    const tNodes = allText(templateDoc);
    for (const node of tNodes) {
      if (node.textContent && node.textContent.includes('[Month DD, YYYY]')) {
        node.textContent = node.textContent.replace('[Month DD, YYYY]', dateIssuedText);
      }
    }
  }

  // 2. Merge numbering.xml (bullet/numbered list defs), remapping IDs so the
  //    content's lists don't collide with anything the template already defines.
  const templateNumberingFile = templateZip.file('word/numbering.xml');
  const contentNumberingFile = contentZip.file('word/numbering.xml');
  let newNumberingXmlStr = null;

  if (contentNumberingFile) {
    const contentNumberingDoc = parseXml(contentNumberingFile.asText());
    const contentAbstractNums = Array.from(contentNumberingDoc.getElementsByTagName('w:abstractNum'));
    const contentNums = Array.from(contentNumberingDoc.getElementsByTagName('w:num'));

    if (templateNumberingFile) {
      const templateNumberingDoc = parseXml(templateNumberingFile.asText());
      const templateNumberingRoot = templateNumberingDoc.getElementsByTagName('w:numbering')[0];

      const existingNumIds = Array.from(templateNumberingDoc.getElementsByTagName('w:num'))
        .map(n => parseInt(n.getAttribute('w:numId'), 10))
        .filter(n => !isNaN(n));
      const existingAbstractIds = Array.from(templateNumberingDoc.getElementsByTagName('w:abstractNum'))
        .map(a => parseInt(a.getAttribute('w:abstractNumId'), 10))
        .filter(n => !isNaN(n));

      const numOffset = existingNumIds.length ? Math.max(...existingNumIds) + 1 : 0;
      const abstractOffset = existingAbstractIds.length ? Math.max(...existingAbstractIds) + 1 : 0;

      const abstractIdMap = {};
      contentAbstractNums.forEach(absNum => {
        const oldId = parseInt(absNum.getAttribute('w:abstractNumId'), 10);
        const newId = oldId + abstractOffset;
        abstractIdMap[oldId] = newId;
        absNum.setAttribute('w:abstractNumId', String(newId));
        templateNumberingRoot.appendChild(absNum);
      });

      const numIdMap = {};
      contentNums.forEach(num => {
        const oldId = parseInt(num.getAttribute('w:numId'), 10);
        const newId = oldId + numOffset;
        numIdMap[oldId] = newId;
        num.setAttribute('w:numId', String(newId));
        const absRef = directChildren(num, 'w:abstractNumId')[0];
        if (absRef) {
          const oldAbs = parseInt(absRef.getAttribute('w:val'), 10);
          if (abstractIdMap[oldAbs] !== undefined) {
            absRef.setAttribute('w:val', String(abstractIdMap[oldAbs]));
          }
        }
        templateNumberingRoot.appendChild(num);
      });

      // remap numId references inside the content body's paragraphs to match
      const contentNumIdRefs = Array.from(contentBody.getElementsByTagName('w:numId'));
      for (const el of contentNumIdRefs) {
        const oldVal = parseInt(el.getAttribute('w:val'), 10);
        if (numIdMap[oldVal] !== undefined) {
          el.setAttribute('w:val', String(numIdMap[oldVal]));
        }
      }

      newNumberingXmlStr = serializeXml(templateNumberingDoc);
    } else {
      newNumberingXmlStr = serializeXml(contentNumberingDoc);
    }
  }

  // 3. Copy over any styles (e.g. Heading2, ListParagraph, TableGrid) the content
  //    uses that the template doesn't already define, so they render correctly
  //    instead of silently falling back to Word defaults.
  const templateStylesFile = templateZip.file('word/styles.xml');
  const contentStylesFile = contentZip.file('word/styles.xml');
  let newStylesXmlStr = null;

  if (templateStylesFile && contentStylesFile) {
    const templateStylesDoc = parseXml(templateStylesFile.asText());
    const contentStylesDoc = parseXml(contentStylesFile.asText());
    const templateStylesRoot = templateStylesDoc.getElementsByTagName('w:styles')[0];

    const existingStyleIds = new Set(
      Array.from(templateStylesDoc.getElementsByTagName('w:style')).map(s => s.getAttribute('w:styleId'))
    );
    const contentStyles = Array.from(contentStylesDoc.getElementsByTagName('w:style'));
    for (const style of contentStyles) {
      const sid = style.getAttribute('w:styleId');
      if (sid && !existingStyleIds.has(sid)) {
        templateStylesRoot.appendChild(style);
      }
    }
    newStylesXmlStr = serializeXml(templateStylesDoc);
  }

  // 4. Splice the content's paragraphs/tables into the template body, right
  //    before its final sectPr, so the template's headers/footers/margins govern.
  if (templateSectPr) {
    for (const child of contentChildren) {
      templateBody.insertBefore(child, templateSectPr);
    }
  } else {
    for (const child of contentChildren) {
      templateBody.appendChild(child);
    }
  }

  const newDocumentXmlStr = serializeXml(templateDoc);

  // 5. Repackage as a new docx: same zip as the template, with document.xml
  //    (and numbering.xml/styles.xml if touched) swapped for the merged versions.
  templateZip.file('word/document.xml', newDocumentXmlStr);
  if (newNumberingXmlStr) templateZip.file('word/numbering.xml', newNumberingXmlStr);
  if (newStylesXmlStr) templateZip.file('word/styles.xml', newStylesXmlStr);

  return templateZip.generate({ type: 'nodebuffer' });
}

app.post('/generate', async (req, res) => {
  let tempDir;
  try {
    const { templateUrl, markdownContent, fileName, headerText, dateIssued } = req.body;
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'doc-'));

    // 1. Markdown -> HTML (no headerText prefix here — the letterhead template
    //    itself already displays "Report date: ..." via its own placeholder,
    //    which we fill in below. Prefixing it into the body too would duplicate it.)
    const fullMarkdown = normalizeMarkdown(markdownContent);
    const htmlBody = marked.parse(fullMarkdown);
    const fullHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
      body { font-family: Calibri, Arial, sans-serif; color: #000000; font-size: 11pt; }
      h2, h3 { color: #000000; font-weight: bold; }
      table { border-collapse: collapse; width: 100%; margin: 8px 0; }
      table, th, td { border: 1px solid #444444; padding: 4px 8px; }
    </style></head><body>${htmlBody}</body></html>`;

    const htmlPath = path.join(tempDir, 'content.html');
    fs.writeFileSync(htmlPath, fullHtml, 'utf8');

    // 2. HTML -> ODT -> DOCX via LibreOffice.
    //    Going straight from HTML to DOCX is a known weak spot for LibreOffice's
    //    filters (especially with tables). ODT is LibreOffice's native format, so
    //    HTML->ODT is a much more faithful conversion, and ODT->DOCX is a
    //    well-trodden, reliable hop.
    const step1 = await runCommand('libreoffice', [
      '--headless',
      `-env:UserInstallation=file://${tempDir}/loconfig1`,
      '--convert-to', 'odt:writer8',
      '--outdir', tempDir,
      htmlPath
    ]);
    console.log('[html->odt] stdout:', step1.stdout);
    console.log('[html->odt] stderr:', step1.stderr);

    const contentOdtPath = path.join(tempDir, 'content.odt');
    if (step1.error || !fs.existsSync(contentOdtPath)) {
      throw new Error(
        `HTML to ODT conversion failed. exec_error=${step1.error ? step1.error.message : 'none'} | ` +
        `stderr=${step1.stderr.slice(0, 500)}`
      );
    }

    const step1b = await runCommand('libreoffice', [
      '--headless',
      `-env:UserInstallation=file://${tempDir}/loconfig1b`,
      '--convert-to', 'docx:MS Word 2007 XML',
      '--outdir', tempDir,
      contentOdtPath
    ]);
    console.log('[odt->docx] stdout:', step1b.stdout);
    console.log('[odt->docx] stderr:', step1b.stderr);

    const contentDocxPath = path.join(tempDir, 'content.docx');
    if (step1b.error || !fs.existsSync(contentDocxPath)) {
      throw new Error(
        `ODT to DOCX conversion failed. exec_error=${step1b.error ? step1b.error.message : 'none'} | ` +
        `stderr=${step1b.stderr.slice(0, 500)}`
      );
    }

    // 3. Download the letterhead template.
    const templateBuffer = await fetchBuffer(templateUrl);

    // 4. Work out the date text to inject into the letterhead's own placeholder.
    //    Accepts either a dedicated `dateIssued` field, or extracts it from a
    //    `headerText` like "Report date: 19 Aug 2026" (strips the label).
    let dateIssuedText = (dateIssued || '').trim();
    if (!dateIssuedText && headerText) {
      dateIssuedText = headerText.replace(/^\s*report date\s*:\s*/i, '').trim();
    }

    // 5. Merge template + generated content (real XML surgery, not a third-party
    //    "merge two docx files" library — those are a known source of silently
    //    corrupt output, which is what caused the previous failure).
    const contentDocxBuffer = fs.readFileSync(contentDocxPath);
    const mergedBuffer = mergeDocx(templateBuffer, contentDocxBuffer, dateIssuedText);
    const mergedPath = path.join(tempDir, 'output.docx');
    fs.writeFileSync(mergedPath, mergedBuffer);

    // 6. Convert merged docx -> PDF
    const step2 = await runCommand('libreoffice', [
      '--headless',
      `-env:UserInstallation=file://${tempDir}/loconfig2`,
      '--convert-to', 'pdf',
      '--outdir', tempDir,
      mergedPath
    ]);
    console.log('[docx->pdf] stdout:', step2.stdout);
    console.log('[docx->pdf] stderr:', step2.stderr);

    const pdfPath = path.join(tempDir, 'output.pdf');
    if (step2.error || !fs.existsSync(pdfPath)) {
      throw new Error(
        `PDF conversion failed. exec_error=${step2.error ? step2.error.message : 'none'} | ` +
        `stderr=${step2.stderr.slice(0, 500)}`
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
