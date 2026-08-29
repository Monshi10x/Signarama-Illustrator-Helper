#!/usr/bin/env node
'use strict';
const fs = require('fs');
const path = require('path');
const {execFileSync} = require('child_process');
const pkg = require('../package.json');
execFileSync(process.execPath, ['scripts/validate-release.js'], {stdio: 'inherit'});
const target = path.join('dist', `signarama-helper-v${pkg.version}`);
fs.rmSync('dist', {recursive: true, force: true}); fs.mkdirSync(target, {recursive: true});
['CSXS', 'data', 'js', 'jsx', 'Proof Templates', 'index.html', 'package.json'].forEach((source) => fs.cpSync(source, path.join(target, source), {recursive: true}));
console.log(`Built ${target}`);
