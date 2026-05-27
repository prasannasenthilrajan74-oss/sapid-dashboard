const express = require('express');
const path = require('path');
const dotenv = require('dotenv');
const bcrypt = require('bcryptjs');

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());

// Block access to sensitive files
app.use((req, res, next) => {
  const file = path.basename(req.path).toLowerCase();
  const sensitiveFiles = ['.env', 'server.js', 'package.json', 'package-lock.json', 'readme.md', '.gitignore'];
  if (sensitiveFiles.includes(file) || req.path.includes('/.git') || req.path.includes('/node_modules')) {
    return res.status(403).send('Access Denied');
  }
  next();
});

// Serve static assets from the current directory
app.use(express.static(__dirname));

// Fallback to index.html for undefined routes
app.get('/{*splat}', (req, res, next) => {
  if (req.path.includes('.')) {
    return next();
  }
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Login Endpoint
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
      return res.json({ success: true });
    } else {
      return res.status(401).json({ success: false, message: 'Invalid username or password.' });
    }
  });
});

app.listen(PORT, () => {
  console.log(`Server is running at http://localhost:${PORT}`);
});
