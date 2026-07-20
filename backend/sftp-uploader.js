'use strict';

const SftpClient = require('ssh2-sftp-client');
const fs   = require('fs');
const path = require('path');

const OUTPUT_DIR = path.join(__dirname, '..', 'output');

function loadConfig() {
  const isProd = (process.env.SFTP_ENV || '').toLowerCase() === 'prod';

  const host       = isProd ? process.env.PROD_SFTP_HOST       : process.env.SFTP_HOST;
  const port       = isProd ? (process.env.PROD_SFTP_PORT || '22') : (process.env.SFTP_PORT || '22');
  const username   = isProd ? process.env.PROD_SFTP_USERNAME   : process.env.SFTP_USERNAME;
  const password   = isProd ? null                              : (process.env.SFTP_PASSWORD || null);
  const keyPath    = isProd ? process.env.PROD_SFTP_PRIVATE_KEY_PATH : process.env.SFTP_PRIVATE_KEY_PATH;
  const passphrase = isProd ? (process.env.PROD_SFTP_PASSPHRASE || null) : (process.env.SFTP_PRIVATE_KEY_PASSPHRASE || null);
  const uploadPath = isProd ? (process.env.PROD_SFTP_REMOTE_DIR || '/inbound/vbkreq/') : (process.env.SFTP_UPLOAD_PATH || '/inbound/vbkreq/');

  if (!host || !username) return null; // local mode — SFTP not configured
  if (!password && !keyPath) return null; // need at least one auth method

  return { host, port, username, password, privateKeyPath: keyPath || null, passphrase, uploadPath, isProd };
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
    sftpEnv: 'LOCAL',
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
      uploadedAt: new Date().toISOString(),
      sftpEnv: cfg.isProd ? 'PROD' : 'TEST'
    };
  } finally {
    await sftp.end().catch(() => {});
  }
}

/**
 * Upload multiple files over a SINGLE shared SFTP connection.
 * files: Array of { filename, xmlContent }
 * Returns Array of { filename, ok, remotePath, bytesSent, uploadedAt, sftpEnv, localMode, error }
 */
async function uploadBatch(files) {
  if (!files || files.length === 0) return [];

  const cfg = loadConfig();

  // No SFTP configured — fall back to saving all files locally
  if (!cfg) {
    return files.map(({ filename, xmlContent }) => {
      const result = saveLocally(filename, xmlContent);
      return { filename, ok: true, ...result };
    });
  }

  const keyPath = cfg.privateKeyPath
    ? (path.isAbsolute(cfg.privateKeyPath)
        ? cfg.privateKeyPath
        : path.join(__dirname, '..', cfg.privateKeyPath))
    : null;

  if (keyPath && !fs.existsSync(keyPath)) {
    throw new Error(`Private key file not found: ${keyPath}`);
  }

  const authOpts = {};
  if (cfg.password)   authOpts.password   = cfg.password;
  if (keyPath)        authOpts.privateKey = fs.readFileSync(keyPath);
  if (cfg.passphrase) authOpts.passphrase = cfg.passphrase;

  const sftp    = new SftpClient();
  const results = [];
  const CONCURRENCY = 5; // simultaneous puts on the same connection

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

    // Upload in chunks of CONCURRENCY simultaneous puts
    for (let i = 0; i < files.length; i += CONCURRENCY) {
      const chunk = files.slice(i, i + CONCURRENCY);
      const chunkResults = await Promise.all(chunk.map(async ({ filename, xmlContent }) => {
        const remotePath = cfg.uploadPath.endsWith('/')
          ? `${cfg.uploadPath}${filename}`
          : `${cfg.uploadPath}/${filename}`;
        try {
          const buffer = Buffer.from(xmlContent, 'utf8');
          await sftp.put(buffer, remotePath);
          saveLocally(filename, xmlContent);
          return {
            filename, ok: true, remotePath,
            bytesSent: buffer.length,
            uploadedAt: new Date().toISOString(),
            sftpEnv: cfg.isProd ? 'PROD' : 'TEST'
          };
        } catch (err) {
          return { filename, ok: false, error: err.message };
        }
      }));
      results.push(...chunkResults);
    }
  } finally {
    await sftp.end().catch(() => {});
  }

  return results;
}

module.exports = { upload, uploadBatch };
