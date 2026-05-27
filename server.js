const express = require('express');
const path = require('path');
const dotenv = require('dotenv');
const bcrypt = require('bcryptjs');
const crypto = require('crypto');

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());

// In-memory session store (sessionId -> { username, createdAt })
const activeSessions = new Map();

// Helper to parse cookies from request headers
function getSessionIdFromCookies(req) {
  const cookieHeader = req.headers.cookie;
  if (!cookieHeader) return null;
  const cookies = cookieHeader.split(';').reduce((acc, cookie) => {
    const parts = cookie.split('=');
    const key = parts[0] ? parts[0].trim() : '';
    const val = parts.slice(1).join('=');
    if (key && val) acc[key] = decodeURIComponent(val);
    return acc;
  }, {});
  return cookies['session_id'] || null;
}

// Block access to sensitive files
app.use((req, res, next) => {
  const file = path.basename(req.path).toLowerCase();
  const sensitiveFiles = ['.env', 'server.js', 'package.json', 'package-lock.json', 'readme.md', '.gitignore'];
  if (sensitiveFiles.includes(file) || req.path.includes('/.git') || req.path.includes('/node_modules')) {
    return res.status(403).send('Access Denied');
  }
  next();
});

// Serve public static assets (login CSS, login JS)
app.use('/public', express.static(path.join(__dirname, 'public')));

// Public login route
app.get('/login', (req, res) => {
  const sessionId = getSessionIdFromCookies(req);
  if (sessionId && activeSessions.has(sessionId)) {
    return res.redirect('/');
  }
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// Login API Endpoint
app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  
  const expectedUser = process.env.USER_ID || 'rane';
  const expectedHash = process.env.PASSWORD_HASH;
  
  if (!username || !password) {
    return res.status(400).json({ success: false, message: 'Username and password are required.' });
  }
  
  if (username !== expectedUser) {
    return res.status(401).json({ success: false, message: 'Invalid username or password.' });
  }
  
  if (!expectedHash) {
    console.error('PASSWORD_HASH is not defined in environment variables.');
    return res.status(500).json({ success: false, message: 'Server configuration error.' });
  }
  
  bcrypt.compare(password, expectedHash, (err, isMatch) => {
    if (err) {
      console.error('Error comparing passwords:', err);
      return res.status(500).json({ success: false, message: 'Internal server error.' });
    }
    
    if (isMatch) {
      // Create session
      const sessionId = crypto.randomUUID();
      activeSessions.set(sessionId, { username, createdAt: Date.now() });

      // Set session cookie
      res.setHeader('Set-Cookie', `session_id=${sessionId}; HttpOnly; Path=/; SameSite=Strict`);
      return res.json({ success: true });
    } else {
      return res.status(401).json({ success: false, message: 'Invalid username or password.' });
    }
  });
});

// Session Validation Middleware
const checkAuth = (req, res, next) => {
  const sessionId = getSessionIdFromCookies(req);
  if (sessionId && activeSessions.has(sessionId)) {
    return next();
  }
  
  if (req.path.startsWith('/api/')) {
    return res.status(401).json({ success: false, message: 'Unauthorized' });
  }
  
  res.redirect('/login');
};

// Protect all remaining routes
app.use(checkAuth);

// Logout API Endpoint
app.post('/api/logout', (req, res) => {
  const sessionId = getSessionIdFromCookies(req);
  if (sessionId) {
    activeSessions.delete(sessionId);
  }
  res.setHeader('Set-Cookie', 'session_id=; HttpOnly; Path=/; SameSite=Strict; Max-Age=0');
  res.json({ success: true });
});

// API Proxy to fetch Google Sheets data securely (prevents public endpoint exposure and client URL leak)
app.get('/api/data', async (req, res) => {
  const defaultSheetUrl = 'https://script.google.com/macros/s/AKfycbytaI36PDf09D7O2RicMWEkGn-JXiew3zPL6bc3OLGKTc0klmd0gUj9ZCfdg2JvY9Sb/exec';
  const sheetUrl = req.query.url || defaultSheetUrl;

  // SSRF Protection: ensure target URL is a Google Sheets/scripts endpoint
  try {
    const parsedUrl = new URL(sheetUrl);
    if (!parsedUrl.hostname.endsWith('google.com')) {
      return res.status(400).json({ success: false, message: 'Invalid sheet URL. Host must be google.com.' });
    }
  } catch (err) {
    return res.status(400).json({ success: false, message: 'Invalid URL format.' });
  }

  // Convert to exportable format if standard sheet URL
  let targetUrl = sheetUrl;
  const isAppsScript = sheetUrl.includes('script.google.com/macros/s/');
  
  if (!isAppsScript) {
    if (sheetUrl.includes('/pub')) {
      if (!sheetUrl.includes('output=csv')) {
        targetUrl = sheetUrl.replace(/\/pub([?#].*)?$/, '/pub?output=csv');
      }
    } else {
      const idMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (idMatch) {
        const sheetId = idMatch[1];
        let gid = '0';
        const gidMatch = sheetUrl.match(/[?&#]gid=([0-9]+)/);
        if (gidMatch) {
          gid = gidMatch[1];
        }
        targetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${gid}`;
      }
    }
  }

  try {
    const response = await fetch(targetUrl);
    if (!response.ok) {
      throw new Error(`Google Sheets responded with status ${response.status}`);
    }

    if (isAppsScript) {
      const data = await response.json();
      res.json(data);
    } else {
      const csvText = await response.text();
      res.setHeader('Content-Type', 'text/csv');
      res.send(csvText);
    }
  } catch (err) {
    console.error('Error fetching sheet data:', err);
    res.status(500).json({ success: false, message: 'Failed to fetch spreadsheet data.' });
  }
});

// Serve ledger page securely
app.get('/ledger', (req, res) => {
  res.sendFile(path.join(__dirname, 'secure', 'ledger.html'));
});

// Serve all other secure static dashboard files (app.js, styles.css, index.html)
app.use(express.static(path.join(__dirname, 'secure')));

// SPA routing fallback
app.get('/{*splat}', (req, res, next) => {
  if (req.path.includes('.')) {
    return next();
  }
  res.sendFile(path.join(__dirname, 'secure', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`Server is running at http://localhost:${PORT}`);
});
