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

function cleanMarkdown(text) {
  let cleaned = text;
  cleaned = cleaned.replace(/(?:=\s){5,}=?/g, '\n\n---\n\n');
  cleaned = cleaned.replace(/(?:_\s){5,}_?/g, '\n\n---\n\n');
  cleaned = cleaned.replace(/([^\n|])\s\|(?=[^|]*\|)/g, '$1\n|');
  cleaned = cleaned.replace(/([^\n])\n(\|)/g, '$1\n\n$2');
  cleaned = cleaned.replace(/([^\n])\s(\d+\.\s)/g, '$1\n\n$2');
  cleaned = cleaned.replace(/([^\n])\s([-*]\s)/g, '$1\n\n$2');

  // Ensure a blank line AFTER every table (a table row followed by a non-pipe line)
  cleaned = cleaned.replace(/(\|.*\|)\n([^\n|])/g, '$1\n\n$2');

  // Ensure a blank line after any heading line (lines starting with #)
  cleaned = cleaned.replace(/(^#{1,6}\s.*$)\n([^\n#])/gm, '$1\n\n$2');

  // Ensure double line breaks between paragraphs generally (collapse 3+ newlines to exactly 2)
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

    // Step 1: Markdown -> docx (using letterhead as style reference)
    await new Promise((resolve, reject) => {
      exec(
        
        `pandoc "${markdownPath}" -o "${docxPath}" --reference-doc="${referenceDocPath}" -f markdown+pipe_tables+grid_tables --standalone`,
        (error) => error ? reject(error) : resolve()
      );
    });

    
      // Step 2: docx -> pdf (using LibreOffice)
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
