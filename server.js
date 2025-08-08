const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

// Middleware to parse JSON bodies
app.use(express.json());

// Serve static files from the current directory
app.use(express.static(path.join(__dirname)));

// Endpoint to receive email ID from the add-in
app.post('/log-email-id', (req, res) => {
    const { emailId, status, timestamp } = req.body;
    const time = timestamp ? new Date(timestamp).toLocaleTimeString() : new Date().toLocaleTimeString();
    
    if (emailId) {
        console.log(`📧 [${time}] Email ID read:`, emailId);
    } else {
        console.log(`📧 [${time}] No email ID available:`, status || 'Unknown reason');
    }
    res.json({ success: true });
});

// Specific route for taskpane.html (serves from addin folder)
app.get('/taskpane.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'addin', 'taskpane.html'));
});

// Serve CSS and JS files from addin folder
app.get('/taskpane.css', (req, res) => {
    res.sendFile(path.join(__dirname, 'addin', 'taskpane.css'));
});

app.get('/taskpane.js', (req, res) => {
    res.sendFile(path.join(__dirname, 'addin', 'taskpane.js'));
});

// Serve assets (icons, etc.)
app.get('/assets/:filename', (req, res) => {
    res.sendFile(path.join(__dirname, 'assets', req.params.filename));
});

// Serve manifest.xml
app.get('/manifest.xml', (req, res) => {
    res.type('application/xml');
    res.sendFile(path.join(__dirname, 'manifest.xml'));
});

// Create self-signed certificate for HTTPS (for development only)
const createSelfSignedCert = () => {
    const certPath = path.join(__dirname, 'localhost.pem');
    const keyPath = path.join(__dirname, 'localhost-key.pem');
    
    if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
        console.log('Self-signed certificates not found. Please create them manually or use a tool like mkcert.');
        console.log('For development, you can use HTTP server instead by setting USE_HTTP=true');
        return null;
    }
    
    return {
        key: fs.readFileSync(keyPath),
        cert: fs.readFileSync(certPath)
    };
};

// Check if we should use HTTP for development
const useHttp = process.env.USE_HTTP === 'true';

if (useHttp) {
    // HTTP server for development (less secure but easier setup)
    app.listen(port, () => {
        console.log(`🚀 Outlook Add-in server running at http://localhost:${port}`);
        console.log(`📁 Serving files from: ${__dirname}`);
        console.log(`📄 Manifest: http://localhost:${port}/manifest.xml`);
        console.log(`📱 Task pane: http://localhost:${port}/taskpane.html`);
        console.log('⚠️  Using HTTP - some Office features may not work. Consider setting up HTTPS for full functionality.');
    });
} else {
    // HTTPS server (recommended for Office add-ins)
    const credentials = createSelfSignedCert();
    
    if (credentials) {
        const httpsServer = https.createServer(credentials, app);
        
        httpsServer.listen(port, () => {
            console.log(`🚀 Outlook Add-in server running at https://localhost:${port}`);
            console.log(`📁 Serving files from: ${__dirname}`);
            console.log(`📄 Manifest: https://localhost:${port}/manifest.xml`);
            console.log(`📱 Task pane: https://localhost:${port}/taskpane.html`);
            console.log('✅ Using HTTPS - full Office add-in functionality available');
        });
    } else {
        console.log('❌ Could not start HTTPS server. Starting HTTP server instead...');
        app.listen(port, () => {
            console.log(`🚀 Outlook Add-in server running at http://localhost:${port}`);
            console.log('⚠️  Using HTTP - consider setting up HTTPS certificates for full functionality.');
        });
    }
}
