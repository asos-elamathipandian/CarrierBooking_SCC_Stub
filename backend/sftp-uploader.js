'use strict';

const SftpClient = require('ssh2-sftp-client');
const fs   = require('fs');
const path = require('path');

const OUTPUT_DIR = path.join(__dirname, '..', 'output');

function loadConfig() {
  const host       = process.env.SFTP_HOST;
  const port       = process.env.SFTP_PORT       || '22';
  const username   = process.env.SFTP_USERNAME;
  const password   = process.env.SFTP_PASSWORD;
  const keyPath    = process.env.SFTP_PRIVATE_KEY_PATH;
  const passphrase = process.env.SFTP_PRIVATE_KEY_PASSPHRASE || null;
  const uploadPath = process.env.SFTP_UPLOAD_PATH || '/inbound/vbkreq/';

  if (!host || !username) return null; // local mode — SFTP not configured
  if (!password && !keyPath) return null; // need at least one auth method

  return { host, port, username, password: password || null, privateKeyPath: keyPath || null, passphrase, uploadPath };
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

  const keyPath = cfg.privateKeyPath
    ? (path.isAbsolute(cfg.privateKeyPath)
        ? cfg.privateKeyPath
        : path.join(__dirname, '..', cfg.privateKeyPath))
    : null;

  if (keyPath && !fs.existsSync(keyPath)) {
    throw new Error(`Private key file not found: ${keyPath}`);
  }

  const authOpts = {};
  if (cfg.password)    authOpts.password   = cfg.password;
  if (keyPath)         authOpts.privateKey = fs.readFileSync(keyPath);
  if (cfg.passphrase)  authOpts.passphrase = cfg.passphrase;

  const sftp = new SftpClient();

  const remotePath = cfg.uploadPath.endsWith('/')
    ? `${cfg.uploadPath}${filename}`
    : `${cfg.uploadPath}/${filename}`;

  try {
    await sftp.connect({
      host: cfg.host,
      port: parseInt(cfg.port, 10),
      username: cfg.username,
      ...authOpts,
      readyTimeout: 20000,
      retries: 2,
      retry_minTimeout: 2000
    });

    const buffer = Buffer.from(xmlContent, 'utf8');
    await sftp.put(buffer, remotePath);

    // Always save a local copy so history download links work
    saveLocally(filename, xmlContent);

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
