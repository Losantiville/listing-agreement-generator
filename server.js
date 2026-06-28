const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const PDFDocument = require('pdfkit');
const { Readable } = require('stream');
const { generateDocx } = require('./files/generate.js');
const crypto = require('crypto');
const { execSync } = require('child_process');

const app = express();
const PORT = process.env.PORT || 3737;

// Load sections configuration from JSON file
let sectionsData = {};
const sectionsPath = path.join(__dirname, 'sections.json');
const syncFolderPath = path.join(process.env.HOME || '/tmp', 'Library/CloudStorage/GoogleDrive-mbergman@astonishcommercial.com/My Drive/Listing Agreement Sync');
const syncFilePath = path.join(syncFolderPath, 'sections.json');

try {
  if (fs.existsSync(sectionsPath)) {
    sectionsData = JSON.parse(fs.readFileSync(sectionsPath, 'utf8'));
  }
} catch (error) {
  console.warn('Warning: Could not load sections.json', error.message);
}

// Ensure sync folder exists
try {
  if (!fs.existsSync(syncFolderPath)) {
    fs.mkdirSync(syncFolderPath, { recursive: true });
  }
} catch (error) {
  console.warn('Warning: Could not create sync folder', error.message);
}

// Password authentication setup
const APP_PASSWORD = process.env.LISTING_APP_PASSWORD || 'astonish123';
const sessions = {}; // In-memory session store (use database in production)

// Middleware to check authentication
function authMiddleware(req, res, next) {
  const sessionId = req.headers['x-session-id'] || req.query.sessionId;
  if (sessionId && sessions[sessionId] && Date.now() < sessions[sessionId].expires) {
    next();
  } else {
    res.status(401).json({ error: 'Unauthorized' });
  }
}

// Helper function to generate PDF from form data
function generatePdfBuffer(data) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ margin: 50 });
    const chunks = [];

    doc.on('data', chunk => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    // Title
    doc.fontSize(18).font('Helvetica-Bold').text(data.contractLabel || 'Listing Agreement', { align: 'center' });
    doc.moveDown();

    // Property Info
    doc.fontSize(12).font('Helvetica-Bold').text('PROPERTY INFORMATION');
    doc.fontSize(10).font('Helvetica');
    doc.text(`Address: ${data.propAddr || ''}`);
    doc.text(`City: ${data.propCity || ''}, State: ${data.propState || ''}, ZIP: ${data.propZip || ''}`);
    if (data.propApn) doc.text(`APN: ${data.propApn}`);
    doc.moveDown();

    // Agent Info
    doc.fontSize(12).font('Helvetica-Bold').text('AGENT INFORMATION');
    doc.fontSize(10).font('Helvetica');
    doc.text(`Name: ${data.agentName || ''}`);
    doc.text(`License: ${data.agentLicense || ''}`);
    if (data.agentEmail) doc.text(`Email: ${data.agentEmail}`);
    if (data.agentPhone) doc.text(`Phone: ${data.agentPhone}`);
    doc.moveDown();

    // Owner Info
    doc.fontSize(12).font('Helvetica-Bold').text('OWNER INFORMATION');
    doc.fontSize(10).font('Helvetica');
    doc.text(`Name: ${data.ownerName || ''}`);
    doc.text(`Address: ${data.ownerAddress || ''}`);
    doc.text(`City: ${data.ownerCity || ''}, State: ${data.ownerState || ''}, ZIP: ${data.ownerZip || ''}`);
    if (data.ownerEmail) doc.text(`Email: ${data.ownerEmail}`);
    doc.moveDown();

    // Additional details
    if (data.listPrice) {
      doc.fontSize(12).font('Helvetica-Bold').text('LISTING PRICE');
      doc.fontSize(10).font('Helvetica').text(`$${data.listPrice}`);
      doc.moveDown();
    }

    // Sale/Lease Commission Details
    doc.fontSize(12).font('Helvetica-Bold').text('COMMISSION DETAILS');
    doc.fontSize(9).font('Helvetica');

    if (data.saleComm) doc.text(`• Sale Commission: ${data.saleComm}% of gross sales price`);
    if (data.leaseComm) {
      const basis = data.commBasis === 'net' ? 'net' : 'gross';
      doc.text(`• Lease Commission: ${data.leaseComm}% of ${basis} rent`);
    }
    if (data.renewComm) doc.text(`• Renewal Commission: ${data.renewComm}%`);
    if (data.listPrice) doc.text(`• Listing Price: $${data.listPrice}`);
    doc.moveDown(0.5);

    // Listing Term
    doc.fontSize(12).font('Helvetica-Bold').text('LISTING TERM');
    doc.fontSize(9).font('Helvetica');
    if (data.tStart) doc.text(`• Start Date: ${data.tStart}`);
    if (data.tEnd) doc.text(`• End Date: ${data.tEnd}`);
    if (data.type === 'auction' && data.auctionDate) doc.text(`• Auction Date: ${data.auctionDate}`);
    doc.moveDown(0.5);

    // Agreement Type
    doc.fontSize(12).font('Helvetica-Bold').text('AGREEMENT TYPE');
    const typeLabels = { lease: 'Exclusive Right to Lease', sell: 'Exclusive Right to Sell', sell_lease: 'Exclusive Right to Sell and Lease', auction: 'Exclusive Right to Sell (Auction)' };
    doc.fontSize(9).font('Helvetica').text(typeLabels[data.type] || data.type);
    doc.moveDown(0.5);

    // Signage Authorization
    if (data.signBldg || data.signAsph || data.signFnce || data.signWndw || data.signYard) {
      doc.fontSize(12).font('Helvetica-Bold').text('SIGNAGE AUTHORIZED');
      doc.fontSize(9).font('Helvetica');
      if (data.signBldg) doc.text('✓ Building Sign(s)');
      if (data.signAsph) doc.text('✓ Asphalt/Rebar Spike Sign(s)');
      if (data.signFnce) doc.text('✓ Fence Sign(s)');
      if (data.signWndw) doc.text('✓ Window Sign(s)');
      if (data.signYard) doc.text('✓ Yard Sign(s)');
      doc.moveDown(0.5);
    }

    // Special Terms for Auction
    if (data.type === 'auction') {
      doc.fontSize(12).font('Helvetica-Bold').text('AUCTION DETAILS');
      doc.fontSize(9).font('Helvetica');
      if (data.res1) doc.text(`• Reserve Price (Property 1): ${data.res1}`);
      if (data.prop2 && data.res2) doc.text(`• Reserve Price (Property 2): ${data.res2}`);
      if (data.prop3 && data.res3) doc.text(`• Reserve Price (Property 3): ${data.res3}`);
      if (data.termFee) doc.text(`• Termination Fee: Greater of ${data.termFee}% or $20,000`);
      doc.moveDown(0.5);
    }

    // Footer
    doc.moveDown(0.5);
    doc.moveTo(40, doc.y).lineTo(doc.page.width - 40, doc.y).stroke();
    doc.moveDown(0.3);
    doc.fontSize(8).font('Helvetica').text('Astonish LLC • 9918 Carver Rd., Suite 101, Cincinnati OH 45242', { align: 'center' });
    doc.fontSize(7).font('Helvetica').text(`Generated: ${new Date().toLocaleDateString()} | For the full formatted agreement, download the Word document (.docx)`, { align: 'center' });

    doc.end();
  });
}

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ limit: '10mb', extended: true }));

// Serve static files from /files directory
app.use(express.static(path.join(__dirname, 'files')));

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'ok', message: 'Server is running' });
});

// Login endpoint
app.post('/api/login', (req, res) => {
  const { password } = req.body;
  if (password === APP_PASSWORD) {
    const sessionId = crypto.randomBytes(32).toString('hex');
    sessions[sessionId] = {
      createdAt: Date.now(),
      expires: Date.now() + (24 * 60 * 60 * 1000) // 24 hour session
    };
    res.json({ success: true, sessionId });
  } else {
    res.status(401).json({ error: 'Invalid password' });
  }
});

// Logout endpoint
app.post('/api/logout', (req, res) => {
  const sessionId = req.body.sessionId;
  if (sessionId && sessions[sessionId]) {
    delete sessions[sessionId];
  }
  res.json({ success: true });
});

// Get sections endpoint (no auth required for local use)
app.get('/api/sections', (req, res) => {
  res.json(sectionsData);
});

// Save sections endpoint (no auth required for local use)
app.post('/api/sections', (req, res) => {
  try {
    const updatedSections = req.body;
    fs.writeFileSync(sectionsPath, JSON.stringify(updatedSections, null, 2), 'utf8');
    sectionsData = updatedSections;
    res.json({ success: true, message: 'Sections saved successfully' });
  } catch (error) {
    console.error('Error saving sections:', error);
    res.status(500).json({ error: 'Failed to save sections', details: error.message });
  }
});

// Upload sections to team sync folder
app.post('/api/sync/upload', (req, res) => {
  try {
    fs.writeFileSync(syncFilePath, JSON.stringify(sectionsData, null, 2), 'utf8');
    res.json({ success: true, message: 'Sections uploaded to team folder', timestamp: new Date().toISOString() });
  } catch (error) {
    console.error('Error uploading to sync:', error);
    res.status(500).json({ error: 'Failed to upload sections', details: error.message });
  }
});

// Download sections from team sync folder
app.post('/api/sync/download', (req, res) => {
  try {
    if (!fs.existsSync(syncFilePath)) {
      return res.status(404).json({ error: 'No shared sections found yet' });
    }
    const teamSections = JSON.parse(fs.readFileSync(syncFilePath, 'utf8'));
    const stats = fs.statSync(syncFilePath);
    res.json({
      success: true,
      sections: teamSections,
      timestamp: stats.mtime.toISOString(),
      message: 'Sections downloaded from team folder'
    });
  } catch (error) {
    console.error('Error downloading from sync:', error);
    res.status(500).json({ error: 'Failed to download sections', details: error.message });
  }
});

// Check sync status
app.get('/api/sync/status', (req, res) => {
  try {
    let status = {
      syncFolderExists: fs.existsSync(syncFolderPath),
      teamSectionsExist: fs.existsSync(syncFilePath),
      localTimestamp: fs.statSync(sectionsPath).mtime.toISOString()
    };
    if (status.teamSectionsExist) {
      status.teamTimestamp = fs.statSync(syncFilePath).mtime.toISOString();
    }
    res.json(status);
  } catch (error) {
    console.error('Error checking sync status:', error);
    res.status(500).json({ error: 'Failed to check sync status', details: error.message });
  }
});

// Generate document endpoint (no auth required for local use)
app.post('/generate', async (req, res) => {
  try {
    const data = req.body;
    const format = req.query.format || 'docx'; // Default to docx

    // Validate required fields
    if (!data.propAddr) {
      return res.status(400).json({ error: 'Property address is required' });
    }

    if (!data.type) {
      return res.status(400).json({ error: 'Contract type is required' });
    }

    const filename = (data.propAddr || 'Agreement').replace(/[^a-zA-Z0-9]/g, '_').slice(0, 30).replace(/_+$/, '');
    let buffer, contentType;

    if (format === 'pdf') {
      // Try to generate PDF from DOCX using LibreOffice, fall back to basic PDF
      try {
        const docxBuffer = await generateDocx(data, sectionsData);
        const tempDocx = path.join('/tmp', `temp_${Date.now()}.docx`);
        const tempPdf = path.join('/tmp', `temp_${Date.now()}.pdf`);

        fs.writeFileSync(tempDocx, docxBuffer);

        try {
          // Try LibreOffice conversion
          execSync(`libreoffice --headless --convert-to pdf --outdir /tmp "${tempDocx}"`, { timeout: 30000 });
          buffer = fs.readFileSync(tempPdf);

          // Clean up temp files
          fs.unlinkSync(tempDocx);
          fs.unlinkSync(tempPdf);
        } catch (e) {
          // LibreOffice not available, use basic PDF
          console.warn('LibreOffice conversion failed, using basic PDF generator:', e.message);
          buffer = await generatePdfBuffer(data);
          fs.unlinkSync(tempDocx);
        }
      } catch (e) {
        console.error('PDF generation error:', e);
        buffer = await generatePdfBuffer(data);
      }

      contentType = 'application/pdf';
      res.set({
        'Content-Type': contentType,
        'Content-Disposition': `attachment; filename="${filename}.pdf"`
      });
    } else {
      // Generate DOCX (default)
      buffer = await generateDocx(data, sectionsData);
      contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
      res.set({
        'Content-Type': contentType,
        'Content-Disposition': `attachment; filename="${filename}.docx"`
      });
    }

    res.send(buffer);
  } catch (error) {
    console.error('Error generating document:', error);
    res.status(500).json({ error: 'Failed to generate document', details: error.message });
  }
});

// Serve index.html for root path
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'files', 'index.html'));
});

// 404 fallback - serve index.html (for client-side routing if needed)
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'files', 'index.html'));
});

// Start server
app.listen(PORT, () => {
  console.log(`✅ Astonish Agreement Generator`);
  console.log(`🚀 Server running at http://localhost:${PORT}`);
  console.log(`📝 API endpoint: POST /generate`);
});

// Graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM received, shutting down gracefully');
  process.exit(0);
});
