# 🚀 Astonish Listing Agreement Generator

Professional commercial real estate listing agreement generator for Astonish Commercial Real Estate Services.

## Features

✅ **Generate Agreements** - Create customized listing agreements in seconds  
✅ **Multiple Formats** - Download as Word (.docx) or PDF  
✅ **Edit Sections** - Customize agreement text without code  
✅ **Team Sync** - Share sections with colleagues  
✅ **Real-time Preview** - See changes as you type  

## Agreement Types

- Exclusive Right to Lease
- Exclusive Right to Sell
- Exclusive Right to Sell and Lease
- Exclusive Right to Sell (Auction)

## Getting Started (Local)

### Prerequisites
- Node.js 18+ ([nodejs.org](https://nodejs.org))
- npm (comes with Node.js)

### Installation

```bash
# Clone the repo or download the files
cd "Listing Agreement Sync"

# Install dependencies
npm install

# Start the server
npm start
```

Open http://localhost:3737 in your browser.

## Deployment to Vercel

1. Push this repo to GitHub
2. Go to [vercel.com](https://vercel.com)
3. Import the GitHub repo
4. Click Deploy
5. Share the URL with your team!

## File Structure

```
├── server.js              # Express server
├── package.json           # Dependencies
├── sections.json          # Agreement sections (editable)
├── vercel.json            # Vercel configuration
├── files/
│   ├── index.html         # Web UI
│   ├── generate.js        # Document generation
│   └── style.css          # Styles
```

## Usage

### Generate an Agreement

1. Select agreement type
2. Fill in property and owner info
3. Customize as needed
4. Click Generate
5. Download as Word or PDF

### Edit Sections

1. Click ✏️ Edit Sections
2. Choose category (Common, Lease, Sell, etc.)
3. Select section to edit
4. Make changes and save
5. Changes apply to future documents

### Team Sync (Local Setup Only)

1. Click ☁️ Team Sync
2. Upload/Download sections with team

## Support

For issues or questions, contact:  
📧 info@astonishcommercial.com  
📞 513.334.3624

---

**Made with ❤️ for Astonish Commercial Real Estate Services**
