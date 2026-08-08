const express = require('express');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const https = require('https');
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
    const { templateUrl, data, fileName } = req.body;
    const templateBuffer = await fetchBuffer(templateUrl);

    console.log('Downloaded size:', templateBuffer.length, 'bytes');
    console.log('First 100 characters:', templateBuffer.toString('utf8', 0, 100));

    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    doc.render(data);
    const outputBuffer = doc.getZip().generate({ type: 'nodebuffer' });
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
