const fs = require('fs');
const path = require('path');
const PDFDocument = require('pdfkit');

// Paths
const manualMdPath = path.join(__dirname, '..', 'USER_MANUAL.md');
const pdfOutputDir = path.join(__dirname, '..', 'src');
const pdfOutputPath = path.join(pdfOutputDir, 'USER_MANUAL.pdf');

// Ensure output directory exists
if (!fs.existsSync(pdfOutputDir)) {
  fs.mkdirSync(pdfOutputDir, { recursive: true });
}

// Read markdown
const mdContent = fs.readFileSync(manualMdPath, 'utf8');

// Generate PDF
const doc = new PDFDocument({ 
  margin: 50,
  info: {
    Title: 'SAP User License Analyzer User Manual',
    Author: 'SAPID Team',
  }
});

const writeStream = fs.createWriteStream(pdfOutputPath);
doc.pipe(writeStream);

// Cover Page / Header
doc.font('Helvetica-Bold').fontSize(26).fillColor('#0ea5e9').text('SAP User License Analyzer', { align: 'center' });
doc.font('Helvetica').fontSize(14).fillColor('#64748b').text('User Manual & Operations Guide', { align: 'center' });
doc.moveDown(1.5);
doc.strokeColor('#38bdf8').lineWidth(2).moveTo(50, doc.y).lineTo(560, doc.y).stroke();
doc.moveDown(2);

// Split markdown into lines
const lines = mdContent.split('\n');

let inCodeBlock = false;

lines.forEach(line => {
  const trimmed = line.trim();
  
  // Code block toggle
  if (trimmed.startsWith('```')) {
    inCodeBlock = !inCodeBlock;
    return;
  }
  
  if (inCodeBlock) {
    doc.font('Courier').fontSize(8.5).fillColor('#0f172a').text(line, { indent: 15 });
    doc.moveDown(0.15);
    return;
  }

  // Set standard font back
  doc.font('Helvetica');

  if (line.startsWith('# ')) {
    // Title / Main header (skipping top one as we drew a custom header)
    const titleText = line.replace('# ', '');
    if (titleText !== 'SAP User License Analyzer User Manual') {
      doc.addPage();
      doc.font('Helvetica-Bold').fontSize(20).fillColor('#0b0f19').text(titleText);
      doc.moveDown(1);
    }
  } else if (line.startsWith('## ')) {
    // Section Header
    doc.moveDown(1);
    doc.font('Helvetica-Bold').fontSize(14).fillColor('#0ea5e9').text(line.replace('## ', ''));
    doc.moveDown(0.5);
  } else if (line.startsWith('### ')) {
    // Sub Section Header
    doc.moveDown(0.75);
    doc.font('Helvetica-Bold').fontSize(11).fillColor('#475569').text(line.replace('### ', ''));
    doc.moveDown(0.4);
  } else if (line.startsWith('- ') || line.startsWith('* ')) {
    // Bullet point
    doc.font('Helvetica').fontSize(9.5).fillColor('#334155').text('•  ' + line.substring(2), { indent: 10 });
    doc.moveDown(0.3);
  } else if (line.match(/^\d+\.\s/)) {
    // Numbered list
    doc.font('Helvetica').fontSize(9.5).fillColor('#334155').text(line, { indent: 10 });
    doc.moveDown(0.3);
  } else if (trimmed === '---') {
    // Divider line
    doc.moveDown(0.5);
    doc.strokeColor('#e2e8f0').lineWidth(0.5).moveTo(50, doc.y).lineTo(560, doc.y).stroke();
    doc.moveDown(0.8);
  } else if (trimmed.length > 0) {
    // Normal paragraph text
    // Replace markdown inline bold **text** with clean bold for simplicity
    const cleanLine = line.replace(/\*\*/g, '');
    doc.font('Helvetica').fontSize(9.5).fillColor('#475569').text(cleanLine, { lineGap: 2 });
    doc.moveDown(0.4);
  } else {
    // Empty line
    doc.moveDown(0.15);
  }
});

doc.end();

writeStream.on('finish', () => {
  console.log('Successfully generated USER_MANUAL.pdf in src/ directory.');
});
