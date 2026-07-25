#!/usr/bin/env node
'use strict';
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const pkg = require('../package.json');
const channel = process.env.UPDATE_CHANNEL || (pkg.version.includes('-') ? 'beta' : 'stable');
const repository = process.env.GITHUB_REPOSITORY || 'Monshi10x/Signarama-Illustrator-Helper';
const tag = process.env.GITHUB_REF_NAME || `v${pkg.version}`;
if(tag !== `v${pkg.version}`) throw new Error(`Tag ${tag} does not match package version ${pkg.version}`);
if(channel === 'stable' && pkg.version.includes('-')) throw new Error('A prerelease cannot be published to stable');
const name = `signarama-helper-v${pkg.version}.zip`, file = path.join('dist', name);
const manifest = {schemaVersion: 1, channel, version: pkg.version, publishedAt: new Date().toISOString(), minimumIllustratorVersion: '16.0.0', maximumIllustratorVersion: null, pluginType: 'cep', pluginId: 'com.signarama.helper', packageType: 'zip', downloadUrl: `https://github.com/${repository}/releases/download/${tag}/${name}`, sha256: crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex'), packageSize: fs.statSync(file).size, releaseNotesUrl: `https://github.com/${repository}/releases/tag/${tag}`, mandatory: false};
fs.writeFileSync(path.join('dist', 'update.json'), JSON.stringify(manifest, null, 2) + '\n');
fs.writeFileSync(path.join('dist', 'SHA256SUMS'), `${manifest.sha256}  ${name}\n`);
console.log(JSON.stringify(manifest, null, 2));
