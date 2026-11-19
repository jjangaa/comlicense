const fs = require('fs');
const path = require('path');

function readAsset(filename) {
  const p1 = path.join(__dirname, 'public', filename);
  const p2 = path.join(__dirname, filename);
  if (fs.existsSync(p1)) return fs.readFileSync(p1, 'utf8');
  if (fs.existsSync(p2)) return fs.readFileSync(p2, 'utf8');
  return '';
}

const HTML = readAsset('index.html');
const CSS = readAsset('styles.css');
const JS = readAsset('client.js');

module.exports = (req, res) => {
  const urlPath = decodeURIComponent((req.url || '/').split('?')[0]);

  if (urlPath.endsWith('styles.css')) {
    res.setHeader('Content-Type', 'text/css; charset=utf-8');
    res.statusCode = 200;
    return res.end(CSS);
  }

  if (urlPath.endsWith('client.js') || urlPath.endsWith('app.js')) {
    res.setHeader('Content-Type', 'application/javascript; charset=utf-8');
    res.statusCode = 200;
    return res.end(JS);
  }

  if (urlPath.endsWith('.xlsm') || urlPath.endsWith('.xlsx') || urlPath.endsWith('.png')) {
    const base = path.basename(urlPath);
    const etcFile = path.join(__dirname, 'etc', base);
    if (fs.existsSync(etcFile)) {
      const buf = fs.readFileSync(etcFile);
      if (base.endsWith('.png')) {
        res.setHeader('Content-Type', 'image/png');
      } else {
        res.setHeader('Content-Type', 'application/octet-stream');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(base)}`);
      }
      res.statusCode = 200;
      return res.end(buf);
    }
  }

  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.statusCode = 200;
  return res.end(HTML);
};
