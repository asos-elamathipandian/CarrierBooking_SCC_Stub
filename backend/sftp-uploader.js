'use strict';

const SftpClient = require('ssh2-sftp-client');
const fs   = require('fs');
const path = require('path');

const OUTPUT_DIR = path.join(__dirname, '..', 'output');

function loadConfig() {
  const configPath = path.join(__dirname, '..', 'config', 'sftp.config.json');
  if (!fs.existsSync(configPath)) return null; // local mode
  const cfg = JSON.parse(fs.readFileSync(configPath, 'utf8'));
  const required = ['host', 'port', 'username', 'privateKeyPath', 'uploadPath'];
  const missing = required.filter(k => !cfg[k]);
  if (missing.length > 0) throw new Error(`SFTP config missing: ${missing.join(', ')}`);
  return cfg;
}

/**
 * LOCAL FALLBACK: save XML to output/ folder when SFTP config is absent.
 */
function saveLocally(filename, xmlContent) {
  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  const outPath = path.join(OUTPUT_DIR, filename);
  fs.writeFileSync(outPath, xmlContent, 'utf8');
  return {
    remotePath: outPath,
    filename,
    bytesSent: Buffer.byteLength(xmlContent, 'utf8'),
    uploadedAt: new Date().toISOString(),
    localMode: true,
    note: 'Saved locally to output/ (SFTP not configured)'
  };
}

/**
 * Upload xmlContent as filename to E2open SFTP.
 */
async function upload(filename, xmlContent) {
  const cfg = loadConfig();
  if (!cfg) return saveLocally(filename, xmlContent);

  const keyPath = path.isAbsolute(cfg.privateKeyPath)
    ? cfg.privateKeyPath
    : path.join(__dirname, '..', cfg.privateKeyPath);

  if (!fs.existsSync(keyPath)) {
    throw new Error(`Private key file not found: ${keyPath}`);
  }

  const privateKey = fs.readFileSync(keyPath);
  const sftp = new SftpClient();

  const remotePath = cfg.uploadPath.endsWith('/')
    ? `${cfg.uploadPath}${filename}`
    : `${cfg.uploadPath}/${filename}`;

  try {
    await sftp.connect({
      host: cfg.host,
      port: parseInt(cfg.port, 10),
      username: cfg.username,
      privateKey,
      readyTimeout: 20000,
      retries: 2,
      retry_minTimeout: 2000
    });

    const buffer = Buffer.from(xmlContent, 'utf8');
    await sftp.put(buffer, remotePath);

    return {
      remotePath,
      filename,
      bytesSent: buffer.length,
      uploadedAt: new Date().toISOString()
    };
  } finally {
    await sftp.end().catch(() => {});
  }
}

module.exports = { upload };
