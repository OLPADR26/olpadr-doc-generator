const express = require('express');
const PizZip = require('pizzip');
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

// ---------- docx XML merge: pure string/regex splicing, no DOM parser ----------
// A prior version used @xmldom/xmldom to parse+re-serialize document.xml. It was
// never actually executed before shipping (no network to install it in dev), and
// it produced a docx LibreOffice couldn't open. This version does surgical string
// splicing instead — every operation here is plain JS string/regex work, which was
// tested end-to-end (against the real letterhead + real generated content, verified
// by rendering the resulting PDF) before being put here.
function mergeDocxXmlStrings({ templateDocumentXml, contentDocumentXml, templateNumberingXml, contentNumberingXml, templateStylesXml, contentStylesXml, dateIssuedText }) {
  // 1. Replace the letterhead's date placeholder, if present, with the real date.
  let mergedTemplateDoc = templateDocumentXml;
  if (dateIssuedText) {
    mergedTemplateDoc = mergedTemplateDoc.replace('[Month DD, YYYY]', dateIssuedText);
  }

  // 2. Extract the content body's paragraphs/tables (everything between <w:body>
  //    and its trailing <w:sectPr>...), and find the template's own trailing
  //    sectPr position so we can splice content in before it, preserving its
  //    headers/footers/margins.
  const contentBodyStart = contentDocumentXml.indexOf('<w:body>') + '<w:body>'.length;
  const contentSectPrIdx = contentDocumentXml.lastIndexOf('<w:sectPr');
  if (contentBodyStart < '<w:body>'.length || contentSectPrIdx === -1) {
    throw new Error('Could not locate <w:body> or trailing <w:sectPr> in generated content document.xml');
  }
  let contentBodyFragment = contentDocumentXml.slice(contentBodyStart, contentSectPrIdx);

  const templateSectPrIdx = mergedTemplateDoc.lastIndexOf('<w:sectPr');
  if (mergedTemplateDoc.indexOf('<w:body>') === -1 || templateSectPrIdx === -1) {
    throw new Error('Could not locate <w:body> or trailing <w:sectPr> in template document.xml');
  }

  // 3. Numbering merge: remap numId/abstractNumId in the content fragment so they
  //    don't collide with anything the template already defines.
  let mergedNumberingXml = templateNumberingXml;
  if (contentNumberingXml) {
    const abstractBlocks = contentNumberingXml.match(/<w:abstractNum\b[\s\S]*?<\/w:abstractNum>/g) || [];
    const numBlocks = contentNumberingXml.match(/<w:num\b[\s\S]*?<\/w:num>/g) || [];

    if (templateNumberingXml) {
      const existingAbstractIds = Array.from(templateNumberingXml.matchAll(/<w:abstractNum\b[^>]*\bw:abstractNumId="(\d+)"/g)).map(m => parseInt(m[1], 10));
      const existingNumIds = Array.from(templateNumberingXml.matchAll(/<w:num\b[^>]*\bw:numId="(\d+)"/g)).map(m => parseInt(m[1], 10));
      const abstractOffset = existingAbstractIds.length ? Math.max(...existingAbstractIds) + 1 : 0;
      const numOffset = existingNumIds.length ? Math.max(...existingNumIds) + 1 : 0;

      const abstractIdMap = {};
      const rewrittenAbstractBlocks = abstractBlocks.map(block => {
        const m = block.match(/w:abstractNumId="(\d+)"/);
        if (!m) return block;
        const oldId = parseInt(m[1], 10);
        const newId = oldId + abstractOffset;
        abstractIdMap[oldId] = newId;
        return block.replace(`w:abstractNumId="${oldId}"`, `w:abstractNumId="${newId}"`);
      });

      const numIdMap = {};
      const rewrittenNumBlocks = numBlocks.map(block => {
        const numIdMatch = block.match(/<w:num\b[^>]*\bw:numId="(\d+)"/);
        if (!numIdMatch) return block;
        const oldNumId = parseInt(numIdMatch[1], 10);
        const newNumId = oldNumId + numOffset;
        numIdMap[oldNumId] = newNumId;
        let rewritten = block.replace(`w:numId="${oldNumId}"`, `w:numId="${newNumId}"`);
        const absRefMatch = rewritten.match(/<w:abstractNumId w:val="(\d+)"/);
        if (absRefMatch) {
          const oldAbs = parseInt(absRefMatch[1], 10);
          if (abstractIdMap[oldAbs] !== undefined) {
            rewritten = rewritten.replace(`<w:abstractNumId w:val="${oldAbs}"`, `<w:abstractNumId w:val="${abstractIdMap[oldAbs]}"`);
          }
        }
        return rewritten;
      });

      contentBodyFragment = contentBodyFragment.replace(/<w:numId w:val="(\d+)"/g, (full, oldVal) => {
        const mapped = numIdMap[parseInt(oldVal, 10)];
        return mapped !== undefined ? `<w:numId w:val="${mapped}"` : full;
      });

      const insertion = rewrittenAbstractBlocks.join('') + rewrittenNumBlocks.join('');
      mergedNumberingXml = templateNumberingXml.replace('</w:numbering>', insertion + '</w:numbering>');
    } else {
      mergedNumberingXml = contentNumberingXml;
    }
  }

  // 4. Styles merge: copy over any style the content uses that the template
  //    doesn't already define (by styleId).
  let mergedStylesXml = templateStylesXml;
  if (contentStylesXml && templateStylesXml) {
    const styleBlocks = contentStylesXml.match(/<w:style\b[\s\S]*?<\/w:style>/g) || [];
    const missingBlocks = styleBlocks.filter(block => {
      const m = block.match(/w:styleId="([^"]+)"/);
      if (!m) return false;
      return !templateStylesXml.includes(`w:styleId="${m[1]}"`);
    });
    if (missingBlocks.length) {
      mergedStylesXml = templateStylesXml.replace('</w:styles>', missingBlocks.join('') + '</w:styles>');
    }
  }

  // 5. Splice the (remapped) content fragment into the template body, right
  //    before its trailing sectPr.
  const finalDocumentXml =
    mergedTemplateDoc.slice(0, templateSectPrIdx) +
    contentBodyFragment +
    mergedTemplateDoc.slice(templateSectPrIdx);

  return { documentXml: finalDocumentXml, numberingXml: mergedNumberingXml, stylesXml: mergedStylesXml };
}

// Merges generated `contentDocxBuffer` (body only) into `templateDocxBuffer`
// (which supplies headers/footers/styles/sectPr). Returns a Buffer of the merged .docx.
function mergeDocx(templateDocxBuffer, contentDocxBuffer, dateIssuedText) {
  const templateZip = new PizZip(templateDocxBuffer);
  const contentZip = new PizZip(contentDocxBuffer);

  const templateDocumentFile = templateZip.file('word/document.xml');
  const contentDocumentFile = contentZip.file('word/document.xml');
  if (!templateDocumentFile) throw new Error('Template docx has no word/document.xml — is it a valid Word file?');
  if (!contentDocumentFile) throw new Error('Generated content docx has no word/document.xml');

  const templateNumberingFile = templateZip.file('word/numbering.xml');
  const contentNumberingFile = contentZip.file('word/numbering.xml');
  const templateStylesFile = templateZip.file('word/styles.xml');
  const contentStylesFile = contentZip.file('word/styles.xml');

  const result = mergeDocxXmlStrings({
    templateDocumentXml: templateDocumentFile.asText(),
    contentDocumentXml: contentDocumentFile.asText(),
    templateNumberingXml: templateNumberingFile ? templateNumberingFile.asText() : null,
    contentNumberingXml: contentNumberingFile ? contentNumberingFile.asText() : null,
    templateStylesXml: templateStylesFile ? templateStylesFile.asText() : null,
    contentStylesXml: contentStylesFile ? contentStylesFile.asText() : null,
    dateIssuedText
  });

  templateZip.file('word/document.xml', result.documentXml);
  if (result.numberingXml) templateZip.file('word/numbering.xml', result.numberingXml);
  if (result.stylesXml) templateZip.file('word/styles.xml', result.stylesXml);

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
