const express = require('express');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_FILE = path.join(__dirname, 'data.json');
const STATIC_DIR = path.join(__dirname, '..', 'outputs');

// Serve static files from outputs/
app.use(express.static(STATIC_DIR, {
  setHeaders: (res, filePath) => {
    if (filePath.endsWith('.html')) {
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
      res.setHeader('Pragma', 'no-cache');
      res.setHeader('Expires', '0');
    } else if (filePath.endsWith('.js')) {
      res.setHeader('Content-Type', 'application/javascript; charset=utf-8');
    } else if (filePath.endsWith('.css')) {
      res.setHeader('Content-Type', 'text/css; charset=utf-8');
    } else if (/\.(png|jpg|jpeg|gif|webp|svg)$/i.test(filePath)) {
      res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
      res.setHeader('Pragma', 'no-cache');
      res.setHeader('Expires', '0');
    }
  }
}));

app.use(express.raw({ type: 'application/json', limit: '1mb' }));

// CORS - allow all origins
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// Manual JSON parse with proper UTF-8
app.use((req, res, next) => {
  if (req.body && Buffer.isBuffer(req.body) && req.body.length > 0) {
    try {
      req.body = JSON.parse(req.body.toString('utf8'));
    } catch (e) {
      return res.status(400).json({ ok: false, msg: 'JSON 解析失败' });
    }
  }
  next();
});

// Load / save helpers
function loadData() {
  try {
    if (fs.existsSync(DATA_FILE)) {
      const raw = fs.readFileSync(DATA_FILE, 'utf-8');
      const data = JSON.parse(raw);
      return Array.isArray(data) ? data : [];
    }
  } catch (e) { console.error('loadData error:', e.message); }
  return [];
}

function saveData(data) {
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf-8');
    return true;
  } catch (e) {
    console.error('saveData error:', e.message);
    return false;
  }
}

// POST /api/submit - append a submission
app.post('/api/submit', (req, res) => {
  const entry = req.body;
  if (!entry || !entry.refCode || !entry.name || !entry.phone) {
    return res.status(400).json({ ok: false, msg: '缺少必填字段（推荐编码/姓名/手机号）' });
  }
  const data = loadData();
  data.push(entry);
  if (saveData(data)) {
    console.log(`[submit] refCode=${entry.refCode} name=${entry.name} total=${data.length}`);
    return res.json({ ok: true, total: data.length });
  }
  res.status(500).json({ ok: false, msg: '服务器存储失败' });
});

// GET /api/submissions - return all submissions
app.get('/api/submissions', (req, res) => {
  const data = loadData();
  res.json(data);
});

// POST /api/clear - clear all submissions
app.post('/api/clear', (req, res) => {
  if (saveData([])) {
    console.log('[clear] all data cleared');
    return res.json({ ok: true });
  }
  res.status(500).json({ ok: false, msg: '清空失败' });
});

// Health check
app.get('/api/health', (req, res) => {
  res.json({ ok: true, time: new Date().toISOString() });
});

app.listen(PORT, () => {
  console.log(`EUGOOO form server running on port ${PORT}`);
});
