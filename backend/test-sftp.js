'use strict';

/**
 * test-sftp.js — Validate SFTP connectivity and config.
 *
 * Usage:
 *   npm run test:sftp
 *   node backend/test-sftp.js
 *
 * Reads SFTP_* variables from .env (or the environment).
 * Does NOT upload any file — only connects, lists the upload directory,
 * then disconnects.  Safe to run against production.
 */

require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const SftpClient = require('ssh2-sftp-client');
const fs   = require('fs');
const path = require('path');

// ── helpers ──────────────────────────────────────────────────────────────────

function pass(msg) { console.log(`  ✔  ${msg}`); }
function fail(msg) { console.error(`  ✖  ${msg}`); }
function info(msg) { console.log(`  ·  ${msg}`); }

// ── config validation ─────────────────────────────────────────────────────────

function loadAndValidateConfig() {
  const cfg = {
    host:           process.env.SFTP_HOST,
    port:           parseInt(process.env.SFTP_PORT || '22', 10),
    username:       process.env.SFTP_USERNAME,
    password:       process.env.SFTP_PASSWORD       || null,
    privateKeyPath: process.env.SFTP_PRIVATE_KEY_PATH || null,
    passphrase:     process.env.SFTP_PRIVATE_KEY_PASSPHRASE || null,
    uploadPath:     process.env.SFTP_UPLOAD_PATH    || '/inbound/vbkreq/',
  };

  let ok = true;

  console.log('\n── SFTP Config Check ──────────────────────────────────');

  if (cfg.host) {
    pass(`SFTP_HOST         = ${cfg.host}`);
  } else {
    fail('SFTP_HOST is not set');
    ok = false;
  }

  pass(`SFTP_PORT         = ${cfg.port}`);

  if (cfg.username) {
    pass(`SFTP_USERNAME     = ${cfg.username}`);
  } else {
    fail('SFTP_USERNAME is not set');
    ok = false;
  }

  if (cfg.password) {
    pass('SFTP_PASSWORD     = (set)');
  } else {
    info('SFTP_PASSWORD     = (not set)');
  }

  if (cfg.privateKeyPath) {
    const absKey = path.isAbsolute(cfg.privateKeyPath)
      ? cfg.privateKeyPath
      : path.join(__dirname, '..', cfg.privateKeyPath);

    if (fs.existsSync(absKey)) {
      pass(`SFTP_PRIVATE_KEY  = ${absKey} (file found)`);
      cfg._resolvedKeyPath = absKey;
    } else {
      fail(`SFTP_PRIVATE_KEY  = ${absKey} (FILE NOT FOUND)`);
      ok = false;
    }
  } else {
    info('SFTP_PRIVATE_KEY_PATH = (not set)');
  }

  if (cfg.passphrase) {
    pass('SFTP_PRIVATE_KEY_PASSPHRASE = (set)');
  } else {
    info('SFTP_PRIVATE_KEY_PASSPHRASE = (not set — key must be unencrypted)');
  }

  if (!cfg.password && !cfg.privateKeyPath) {
    fail('No auth method set — provide SFTP_PASSWORD or SFTP_PRIVATE_KEY_PATH');
    ok = false;
  }

  pass(`SFTP_UPLOAD_PATH  = ${cfg.uploadPath}`);

  return { cfg, ok };
}

// ── connectivity test ─────────────────────────────────────────────────────────

async function testConnectivity(cfg) {
  console.log('\n── Connectivity Test ──────────────────────────────────');

  const authOpts = {};
  if (cfg.password)           authOpts.password   = cfg.password;
  if (cfg._resolvedKeyPath)   authOpts.privateKey = fs.readFileSync(cfg._resolvedKeyPath);
  if (cfg.passphrase)         authOpts.passphrase = cfg.passphrase;

  const sftp = new SftpClient();

  try {
    info(`Connecting to ${cfg.host}:${cfg.port} as ${cfg.username} …`);
    await sftp.connect({
      host: cfg.host,
      port: cfg.port,
      username: cfg.username,
      ...authOpts,
      readyTimeout: 15000,
    });
    pass('Connected successfully');

    // List the upload directory to confirm it exists and is accessible
    info(`Listing remote path: ${cfg.uploadPath}`);
    let listing;
    try {
      listing = await sftp.list(cfg.uploadPath);
      pass(`Directory exists — ${listing.length} item(s) found`);
      if (listing.length > 0) {
        const sample = listing.slice(0, 5).map(f => f.name).join(', ');
        info(`Sample entries: ${sample}`);
      }
    } catch (listErr) {
      fail(`Could not list ${cfg.uploadPath}: ${listErr.message}`);
      info('The directory may not exist yet — uploads will fail until it is created.');
    }

    return true;
  } catch (connErr) {
    fail(`Connection failed: ${connErr.message}`);
    return false;
  } finally {
    await sftp.end().catch(() => {});
    info('Connection closed');
  }
}

// ── main ──────────────────────────────────────────────────────────────────────

(async () => {
  console.log('\n╔══════════════════════════════════════════════╗');
  console.log('║      CarrierBookingStub — SFTP Config Test   ║');
  console.log('╚══════════════════════════════════════════════╝');

  const { cfg, ok: configOk } = loadAndValidateConfig();

  if (!configOk) {
    console.log('\n── Result ─────────────────────────────────────────────');
    fail('Config validation failed — fix the errors above, then re-run.');
    console.log('\nTip: copy .env.example to .env and fill in the values.\n');
    process.exit(1);
  }

  pass('Config validation passed');

  const connOk = await testConnectivity(cfg);

  console.log('\n── Result ─────────────────────────────────────────────');
  if (connOk) {
    pass('SFTP test PASSED — ready to upload VBKREQ files.\n');
    process.exit(0);
  } else {
    fail('SFTP test FAILED — check the errors above.\n');
    process.exit(1);
  }
})();
