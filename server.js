const express = require('express');
const cors = require('cors');
const path = require('path');
const { generateDocx } = require('./files/generate.js');

const app = express();
const PORT = process.env.PORT || 3737;

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

// Generate document endpoint
app.post('/generate', async (req, res) => {
  try {
    const data = req.body;

    // Validate required fields
    if (!data.propAddr) {
      return res.status(400).json({ error: 'Property address is required' });
    }

    if (!data.type) {
      return res.status(400).json({ error: 'Contract type is required' });
    }

    // Generate the document
    const docBuffer = await generateDocx(data);

    // Send the document
    const filename = (data.propAddr || 'Agreement').replace(/[^a-zA-Z0-9]/g, '_').slice(0, 30).replace(/_+$/, '');
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="${filename}.docx"`
    });
    res.send(docBuffer);
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
