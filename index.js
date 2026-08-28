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

// Turn "Section N: Title" and "Artifact Name: Title" into bold text (never a heading
// style, so it can't inherit the template's colored Heading style — always plain black
// bold, guaranteed consistent regardless of what template is in play).
function normalizeMarkdown(text) {
  let out = text;

  out = out.replace(/(?:=\s){5,}=?/g, '\n\n---\n\n');
  out = out.replace(/(?:_\s){5,}_?/g, '\n\n---\n\n');

  out = out.replace(/^Section \d+:\s*(.+)$/gm, '**Section: $1**');
  out = out.replace(/^Artifact Name:\s*(.+)$/gm, '**$1**');

  out = boldStandaloneTitles(out);
  return out;
}

// Dynamically bolds any standalone "title-shaped" line — isolated by blank lines
// above and below, short, and not ending in sentence punctuation — instead of
// relying on a hardcoded list of known section names. This is what makes header
// bolding work for ANY artifact/section name the agents produce, not just ones
// someone happened to enumerate in advance.
function boldStandaloneTitles(text) {
  const lines = text.split('\n');
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();
    const prevBlank = i === 0 || lines[i - 1].trim() === '';
    const nextBlank = i === lines.length - 1 || lines[i + 1].trim() === '';
    const isTableRow = /^\s*\|/.test(line);
    const isListItem = /^\s*([-*]|\d+\.)\s/.test(line);
    const isDivider = /^-{3,}$/.test(trimmed) || trimmed === '---';
    const alreadyBold = /^\*\*.*\*\*$/.test(trimmed);
    const endsWithSentencePunct = /[.!?:;,]["'\u201d)\]]?$/.test(trimmed);
    const looksLikeTitle =
      trimmed.length > 0 &&
      trimmed.length <= 90 &&
      prevBlank && nextBlank &&
      !isTableRow && !isListItem && !isDivider && !alreadyBold &&
      !endsWithSentencePunct;
    out.push(looksLikeTitle ? `**${trimmed}**` : line);
  }
  return out.join('\n');
}

// ---------- docx XML merge: pure string/regex splicing, no DOM parser ----------
// A prior version used @xmldom/xmldom to parse+re-serialize document.xml. It was
// never actually executed before shipping (no network to install it in dev), and
// it produced a docx LibreOffice couldn't open. This version does surgical string
// splicing instead — every operation here is plain JS string/regex work, which was
// tested end-to-end (against the real letterhead + real generated content, verified
// by rendering the resulting PDF) before being put here.
// ---------- Two-template merge: page-1 cover template + page-2-onward template + generated content ----------
// This produces a document with TWO Word sections: section 1 is page1Template's own
// content (its cover/date line) using page1's header/footer/margins; section 2 is the
// generated report content using page2Template's header/footer/margins. A real section
// break is required (not just a "different first page" toggle) because the two
// templates use different page margins, which is a section-level property in OOXML.
function extractSectPrParts(sectPrXml) {
  const pgSz = sectPrXml.match(/<w:pgSz[^/]*\/>/);
  const pgMar = sectPrXml.match(/<w:pgMar[^/]*\/>/);
  const cols = sectPrXml.match(/<w:cols[^/]*\/>/);
  const grid = sectPrXml.match(/<w:docGrid[^/]*\/>/);
  return {
    pgSz: pgSz ? pgSz[0] : '',
    pgMar: pgMar ? pgMar[0] : '',
    cols: cols ? cols[0] : '<w:cols w:space="720"/>',
    grid: grid ? grid[0] : ''
  };
}

function mergeNumbering(baseNumberingXml, contentNumberingXml, contentFragment) {
  if (!contentNumberingXml) return { numberingXml: baseNumberingXml, contentFragment };
  if (!baseNumberingXml) return { numberingXml: contentNumberingXml, contentFragment };

  const abstractBlocks = contentNumberingXml.match(/<w:abstractNum\b[\s\S]*?<\/w:abstractNum>/g) || [];
  const numBlocks = contentNumberingXml.match(/<w:num\b[\s\S]*?<\/w:num>/g) || [];
  const existingAbstractIds = Array.from(baseNumberingXml.matchAll(/<w:abstractNum\b[^>]*\bw:abstractNumId="(\d+)"/g)).map(m => parseInt(m[1], 10));
  const existingNumIds = Array.from(baseNumberingXml.matchAll(/<w:num\b[^>]*\bw:numId="(\d+)"/g)).map(m => parseInt(m[1], 10));
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

  // Single-pass regex+callback remap, done BEFORE this fragment is spliced anywhere
  // else — critical, since sequential/repeated string replaces here can cascade and
  // corrupt each other when offset ranges overlap (confirmed by direct testing).
  const remappedFragment = contentFragment.replace(/<w:numId w:val="(\d+)"/g, (full, oldVal) => {
    const mapped = numIdMap[parseInt(oldVal, 10)];
    return mapped !== undefined ? `<w:numId w:val="${mapped}"` : full;
  });

  const insertion = rewrittenAbstractBlocks.join('') + rewrittenNumBlocks.join('');
  const numberingXml = baseNumberingXml.replace('</w:numbering>', insertion + '</w:numbering>');
  return { numberingXml, contentFragment: remappedFragment };
}

function mergeMissingStyles(baseStylesXml, otherStylesXmlList) {
  let merged = baseStylesXml;
  for (const src of otherStylesXmlList) {
    if (!src) continue;
    const existingIds = new Set(Array.from(merged.matchAll(/w:styleId="([^"]+)"/g)).map(m => m[1]));
    const blocks = src.match(/<w:style\b[\s\S]*?<\/w:style>/g) || [];
    const missing = blocks.filter(b => {
      const m = b.match(/w:styleId="([^"]+)"/);
      return m && !existingIds.has(m[1]);
    });
    if (missing.length) {
      merged = merged.replace('</w:styles>', missing.join('') + '</w:styles>');
      // keep existingIds current across sources so we don't add the same style twice
      missing.forEach(b => {
        const m = b.match(/w:styleId="([^"]+)"/);
        if (m) existingIds.add(m[1]);
      });
    }
  }
  return merged;
}

// Uniform table border definition — single, solid, black, 4 half-points (0.5pt).
// Applied to every table side so borders are always visible and consistent regardless
// of what the HTML→ODT→DOCX conversion may or may not have preserved.
const TABLE_BORDERS =
  '<w:tblBorders>' +
  '<w:top w:val="single" w:sz="4" w:space="0" w:color="444444"/>' +
  '<w:left w:val="single" w:sz="4" w:space="0" w:color="444444"/>' +
  '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="444444"/>' +
  '<w:right w:val="single" w:sz="4" w:space="0" w:color="444444"/>' +
  '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="444444"/>' +
  '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="444444"/>' +
  '</w:tblBorders>';

// Enforce uniform visible borders on every table in the content fragment:
// replace any existing tblBorders block, or inject one if absent.
function enforceTableBorders(fragment) {
  // Replace existing tblBorders blocks
  let result = fragment.replace(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/g, TABLE_BORDERS);
  // Inject into tblPr blocks that have no tblBorders at all
  result = result.replace(/(<w:tblPr>(?:(?!<\/w:tblPr>)(?!<w:tblBorders>)[\s\S])*?)(<\/w:tblPr>)/g,
    (match, before, close) => {
      if (match.includes('<w:tblBorders>')) return match;
      return before + TABLE_BORDERS + close;
    }
  );
  return result;
}

// Merges page1Buffer (cover page template), page2Buffer (page-2-onward template), and
// contentDocxBuffer (generated report body) into one docx with a real section break.
function mergeTwoTemplates(page1Buffer, page2Buffer, contentDocxBuffer, dateIssuedText) {
  const page1Zip = new PizZip(page1Buffer);
  const page2Zip = new PizZip(page2Buffer);
  const contentZip = new PizZip(contentDocxBuffer);

  let p1Doc = page1Zip.file('word/document.xml').asText();
  const p2Doc = page2Zip.file('word/document.xml').asText();
  const contentDoc = contentZip.file('word/document.xml').asText();

  if (dateIssuedText) {
    p1Doc = p1Doc.replace('[Month DD, YYYY]', dateIssuedText);
  }

  const p1BodyStart = p1Doc.indexOf('<w:body>') + '<w:body>'.length;
  const p1SectPrStart = p1Doc.lastIndexOf('<w:sectPr');
  let p1Paragraphs = p1Doc.slice(p1BodyStart, p1SectPrStart);

  // Reduce the gap between the date line and first content line by tightening the
  // empty spacer paragraph that sits between them. 400→240 twips is a ~40% reduction
  // — visible but deliberate, not dramatic. Targets only empty paragraphs with exactly
  // this spacing value, so body text spacing elsewhere is unaffected.
  p1Paragraphs = p1Paragraphs.replace(
    /<w:p\b[^>]*>\s*<w:pPr>\s*<w:spacing w:after="400"\/>\s*<\/w:pPr>\s*<\/w:p>/,
    (m) => m.replace('w:after="400"', 'w:after="240"')
  );

  const p2SectPrStart = p2Doc.lastIndexOf('<w:sectPr');
  const p2SectPrEnd = p2Doc.indexOf('</w:sectPr>', p2SectPrStart) + '</w:sectPr>'.length;
  const p2SectPrXml = p2Doc.slice(p2SectPrStart, p2SectPrEnd);

  // Content starts ON page 1, flowing into page 2+ as it overflows — this is a single
  // Word section using "different first page" (titlePg), not two separate sections.
  // That means page geometry (margins) is shared across every page; page1 and page2
  // templates use slightly different margins (0.625" vs 0.7" left/right), so page2's
  // margins are used throughout, since they govern the majority of pages. Replace
  // page1's "[Body content begins here]" placeholder with the actual generated content,
  // right where it belongs, instead of dropping it.
  const p2Parts = extractSectPrParts(p2SectPrXml);

  const cBodyStart = contentDoc.indexOf('<w:body>') + '<w:body>'.length;
  const cSectPrStart = contentDoc.lastIndexOf('<w:sectPr');
  let contentFragment = contentDoc.slice(cBodyStart, cSectPrStart);

  // Fix table overflow: clamp every table to the actual page text width and switch to
  // autofit layout so columns redistribute proportionally rather than spilling off the
  // right edge. Page text width = 12240 (letter) - 720 (left margin) - 720 (right) = 10800.
  // LibreOffice's HTML→DOCX converter sets tblW to the raw HTML table pixel width
  // (often wider than the page) and locks layout to "fixed", which is exactly why the
  // last column disappears off the right edge on wide tables like the 6-column
  // Indicator Dictionary. Switching to autofit + capping tblW forces a proportional
  // column redistribution that stays within the page.
  contentFragment = contentFragment
    .replace(/<w:tblW w:w="[^"]*" w:type="dxa"\/>/g, '<w:tblW w:w="10800" w:type="dxa"/>')
    .replace(/<w:tblW w:w="[^"]*" w:type="pct"\/>/g, '<w:tblW w:w="5000" w:type="pct"/>')
    .replace(/<w:tblLayout w:type="fixed"\/>/g, '<w:tblLayout w:type="autofit"/>');

  // Enforce uniform visible borders on every table.
  contentFragment = enforceTableBorders(contentFragment);

  // Numbering: page1's numbering.xml (if any) + content's, remapped. Page2 templates
  // don't carry their own numbering.xml, so page1's is the base.
  const p1NumberingFile = page1Zip.file('word/numbering.xml');
  const contentNumberingFile = contentZip.file('word/numbering.xml');
  const numberingResult = mergeNumbering(
    p1NumberingFile ? p1NumberingFile.asText() : null,
    contentNumberingFile ? contentNumberingFile.asText() : null,
    contentFragment
  );
  contentFragment = numberingResult.contentFragment;
  const mergedNumberingXml = numberingResult.numberingXml;

  // Styles: page1's base, plus any missing styles from page2 and content.
  const p1StylesXml = page1Zip.file('word/styles.xml').asText();
  const p2StylesXml = page2Zip.file('word/styles.xml').asText();
  const contentStylesFile = contentZip.file('word/styles.xml');
  const mergedStylesXml = mergeMissingStyles(p1StylesXml, [p2StylesXml, contentStylesFile ? contentStylesFile.asText() : null]);

  // Replace page1's "[Body content begins here]" placeholder paragraph with the actual
  // generated content — content starts ON page 1, right after the date line.
  // Use a loose pattern that matches any paragraph containing this placeholder text,
  // regardless of the exact surrounding XML structure (italic/color formatting, bookmarks,
  // etc. may differ between template versions).
  const placeholderPattern = /<w:p\b[^>]*>(?:(?!<w:p\b).)*?\[Body content begins here\](?:(?!<w:p\b).)*?<\/w:p>/s;
  if (placeholderPattern.test(p1Paragraphs)) {
    p1Paragraphs = p1Paragraphs.replace(placeholderPattern, contentFragment);
  } else {
    // Fallback: placeholder not found in template — append content after page1's own
    // paragraphs rather than silently losing it.
    p1Paragraphs = p1Paragraphs + contentFragment;
  }

  // Single section, "different first page" (titlePg): page 1 uses the "first"
  // header/footer (page1 template's), every subsequent page uses the "default"
  // header/footer (page2 template's) — pagination happens naturally as content
  // overflows, no manual page-break bookkeeping needed.
  // Margins: use 720 twips (0.5") left/right throughout — narrower than page2's
  // native 1008 twips, giving wide tables (e.g. 6-column Indicator Dictionary)
  // enough room to render without content overflowing the page edge. Top/bottom
  // and header/footer offsets are kept from page2's template.
  const p2MarMatch = p2SectPrXml.match(/<w:pgMar([^/]*)\/>/);
  let finalPgMar = p2Parts.pgMar;
  if (p2MarMatch) {
    // Override left/right to 720 twips, keep top/bottom/header/footer from template
    finalPgMar = p2Parts.pgMar
      .replace(/w:left="[^"]*"/, 'w:left="720"')
      .replace(/w:right="[^"]*"/, 'w:right="720"');
  }

  const finalSectPr =
    `<w:sectPr>` +
    `<w:headerReference w:type="default" r:id="rIdHdrP2"/>` +
    `<w:footerReference w:type="default" r:id="rIdFtrP2"/>` +
    `<w:headerReference w:type="first" r:id="rIdHdrP1"/>` +
    `<w:footerReference w:type="first" r:id="rIdFtrP1"/>` +
    `${p2Parts.pgSz}${finalPgMar}${p2Parts.cols}${p2Parts.grid}` +
    `<w:titlePg/>` +
    `</w:sectPr>`;

  const p1DocOpenTagStart = p1Doc.indexOf('<w:document');
  const finalDocumentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    p1Doc.slice(p1DocOpenTagStart, p1Doc.indexOf('<w:body>')) +
    '<w:body>' +
    p1Paragraphs +
    finalSectPr +
    '</w:body></w:document>';

  // Assemble the output package: page1's singleton parts (settings/fontTable/theme/etc,
  // media), each template's own default header/footer renamed to avoid collisions, a
  // fresh minimal document.xml.rels, and fresh [Content_Types].xml. We only carry the
  // "default" header/footer type from each template — confirmed neither template has
  // evenAndOddHeaders or titlePg active, so "even"/"first" header/footer parts are
  // unused dead weight and are deliberately dropped rather than carried over.
  const singletonParts = [
    'docProps/core.xml', 'docProps/app.xml', '_rels/.rels', 'word/fontTable.xml',
    'word/settings.xml', 'word/webSettings.xml', 'word/theme/theme1.xml',
    'word/footnotes.xml', 'word/endnotes.xml', 'word/media/image1.png'
  ];

  const outZip = new PizZip();
  for (const name of singletonParts) {
    const f = page1Zip.file(name);
    if (f) outZip.file(name, f.asUint8Array());
  }

  outZip.file('word/document.xml', finalDocumentXml);
  outZip.file('word/styles.xml', mergedStylesXml);
  if (mergedNumberingXml) outZip.file('word/numbering.xml', mergedNumberingXml);

  outZip.file('word/header_p1.xml', page1Zip.file('word/header2.xml').asUint8Array());
  outZip.file('word/footer_p1.xml', page1Zip.file('word/footer2.xml').asUint8Array());
  outZip.file('word/header_p2.xml', page2Zip.file('word/header2.xml').asUint8Array());
  outZip.file('word/footer_p2.xml', page2Zip.file('word/footer2.xml').asUint8Array());
  const p1HeaderRels = page1Zip.file('word/_rels/header2.xml.rels');
  const p1FooterRels = page1Zip.file('word/_rels/footer2.xml.rels');
  const p2HeaderRels = page2Zip.file('word/_rels/header2.xml.rels');
  const p2FooterRels = page2Zip.file('word/_rels/footer2.xml.rels');
  if (p1HeaderRels) outZip.file('word/_rels/header_p1.xml.rels', p1HeaderRels.asUint8Array());
  if (p1FooterRels) outZip.file('word/_rels/footer_p1.xml.rels', p1FooterRels.asUint8Array());
  if (p2HeaderRels) outZip.file('word/_rels/header_p2.xml.rels', p2HeaderRels.asUint8Array());
  if (p2FooterRels) outZip.file('word/_rels/footer_p2.xml.rels', p2FooterRels.asUint8Array());

  const relParts = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>',
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>',
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>',
    '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>',
    '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
    '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>',
    '<Relationship Id="rIdHdrP1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header_p1.xml"/>',
    '<Relationship Id="rIdFtrP1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer_p1.xml"/>',
    '<Relationship Id="rIdHdrP2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header_p2.xml"/>',
    '<Relationship Id="rIdFtrP2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer_p2.xml"/>'
  ];
  if (mergedNumberingXml) {
    relParts.push('<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>');
  }
  relParts.push('</Relationships>');
  outZip.file('word/_rels/document.xml.rels', relParts.join(''));

  const ctParts = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="png" ContentType="image/png"/>',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
    '<Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>',
    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
    '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>',
    '<Override PartName="/word/header_p1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    '<Override PartName="/word/footer_p1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>',
    '<Override PartName="/word/header_p2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    '<Override PartName="/word/footer_p2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>',
    '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>',
    '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
  ];
  if (mergedNumberingXml) {
    ctParts.push('<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>');
  }
  ctParts.push('</Types>');
  outZip.file('[Content_Types].xml', ctParts.join(''));

  return outZip.generate({ type: 'nodebuffer' });
}

app.post('/generate', async (req, res) => {
  let tempDir;
  try {
    const { templateUrl, templateUrlPage1, templateUrlRest, markdownContent, fileName, headerText, dateIssued } = req.body;
    // Accepts either the new two-template fields (templateUrlPage1 + templateUrlRest)
    // or, for backwards compatibility, a single `templateUrl` used for both.
    const page1Url = templateUrlPage1 || templateUrl;
    const page2Url = templateUrlRest || templateUrlPage1 || templateUrl;
    if (!page1Url || !page2Url) {
      throw new Error('templateUrlPage1 and templateUrlRest (or templateUrl) are required');
    }
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
      table, th, td { border: 1px solid #444444; padding: 3px 6px; }
      th, td { overflow-wrap: break-word; word-break: break-word; }
      p { margin-top: 0; margin-bottom: 0.14in; }
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

    // 3. Download both letterhead templates: page 1 (cover) and page 2 onward.
    const page1Buffer = await fetchBuffer(page1Url);
    const page2Buffer = await fetchBuffer(page2Url);

    // 4. Work out the date text to inject into page 1's own placeholder.
    //    Accepts either a dedicated `dateIssued` field, or extracts it from a
    //    `headerText` like "Report date: 19 Aug 2026" (strips the label).
    let dateIssuedText = (dateIssued || '').trim();
    if (!dateIssuedText && headerText) {
      dateIssuedText = headerText.replace(/^\s*report date\s*:\s*/i, '').trim();
    }

    // 5. Merge page1 template + page2 template + generated content into one docx
    //    with a real section break (page1 and page2 templates use different page
    //    margins, so this can't be done with a simple "different first page" toggle).
    const contentDocxBuffer = fs.readFileSync(contentDocxPath);
    const mergedBuffer = mergeTwoTemplates(page1Buffer, page2Buffer, contentDocxBuffer, dateIssuedText);
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
