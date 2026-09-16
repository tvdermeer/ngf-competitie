#!/usr/bin/env node
/*
 * Encrypt a JSON payload for the TC Ranking pages.
 *
 * Reads the plaintext JSON from stdin and writes a WebCrypto-compatible blob to
 * stdout. The password is taken from the TC_PASSWORD environment variable so it
 * never shows up in the process list.
 *
 * WebCrypto's AES-GCM expects "ciphertext || tag" as a single buffer, so the
 * authentication tag produced by Node is appended to the ciphertext.
 */

const crypto = require('crypto');

const password = process.env.TC_PASSWORD;
const iterations = parseInt(process.env.TC_ITERATIONS || '250000', 10);

if (!password) {
  process.stderr.write('TC_PASSWORD is niet gezet.\n');
  process.exit(1);
}

let input = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', (chunk) => { input += chunk; });
process.stdin.on('end', () => {
  const salt = crypto.randomBytes(16);
  const iv = crypto.randomBytes(12);
  const key = crypto.pbkdf2Sync(password, salt, iterations, 32, 'sha256');

  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
  const ciphertext = Buffer.concat([cipher.update(input, 'utf8'), cipher.final()]);
  const tag = cipher.getAuthTag();

  const out = {
    v: 1,
    alg: 'AES-256-GCM',
    kdf: 'PBKDF2-SHA256',
    iterations: iterations,
    salt: salt.toString('base64'),
    iv: iv.toString('base64'),
    ct: Buffer.concat([ciphertext, tag]).toString('base64')
  };
  process.stdout.write(JSON.stringify(out));
});
