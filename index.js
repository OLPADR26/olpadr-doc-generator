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

app.post('/generate', async (req, res) => {
  try {
    const { templateUrl, markdownContent, fileName, headerText } = req.body;

    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'doc-'));
    const referenceDocPath = path.join(tempDir, 'reference.docx');
    const markdownPath = path.join(tempDir, 'content.md');
    const outputPath = path.join(tempDir, 'output.docx');

    // Download the letterhead template to use as the style reference
    const templateBuffer = await fetchBuffer(templateUrl);
    fs.writeFileSync(referenceDocPath, templateBuffer);

    // Build the markdown content (optional header line + the real report body)
    const fullMarkdown = (headerText ? `${headerText}\n\n---\n\n` : '') + markdownContent;
    fs.writeFileSync(markdownPath, fullMarkdown, 'utf8');

    // Run Pandoc: convert markdown into a docx, styled using our letterhead as reference
    await new Promise((resolve, reject) => {
      exec(
        `pandoc "${markdownPath}" -o "${outputPath}" --reference-doc="${referenceDocPath}"`,
        (error) => error ? reject(error) : resolve()
      );
    });

    const outputBuffer = fs.readFileSync(outputPath);

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="${fileName || 'report.docx'}"`,
      'Content-Length': outputBuffer.length
    });
    res.send(outputBuffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Running on port ${PORT}`));
