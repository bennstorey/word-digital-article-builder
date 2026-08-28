#!/usr/bin/env node
// Builds word-digital-plugin.js (Content Station SDK plug-in) from index.html
// so the conversion engine has a single source of truth.
//
// Extracts everything between "const TEMPLATES" and the "── UI state" marker
// (templates, parsers, builders — all pure browser JS), then wraps it in the
// plug-in shell from plugin-shell.js.

const fs = require('fs');
const path = require('path');

const root = __dirname;
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

const start = html.indexOf('const TEMPLATES');
const end = html.indexOf('// ─── UI state');
if (start === -1 || end === -1) throw new Error('extraction markers not found in index.html');
const engine = html.slice(start, end);

const shell = fs.readFileSync(path.join(root, 'plugin-shell.js'), 'utf8');
let out = shell.replace('/*__ENGINE__*/', () => engine);

// Stamp a build id so a running Studio can be asked which version it loaded:
// type __wdVersion in the console. Cross-origin plug-in files are cached by
// GitHub Pages for 10 minutes, so "is my build live?" is a real question.
const crypto = require('crypto');
const buildId = crypto.createHash('sha1').update(out).digest('hex').slice(0, 8);
out = out.replace('__BUILD_ID__', buildId);

fs.writeFileSync(path.join(root, 'word-digital-plugin.js'), out);
console.log('wrote word-digital-plugin.js', fs.statSync(path.join(root, 'word-digital-plugin.js')).size, 'bytes');
