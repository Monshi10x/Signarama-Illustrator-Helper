#!/usr/bin/env node
'use strict';
const fs = require('fs');
const path = require('path');
const {execFileSync} = require('child_process');
const pkg = require('../package.json');
const folder = path.resolve('dist', `signarama-helper-v${pkg.version}`);
if(!fs.existsSync(folder)) execFileSync(process.execPath, ['scripts/build.js'], {stdio: 'inherit'});
const output = path.resolve('dist', `signarama-helper-v${pkg.version}.zip`);
fs.rmSync(output, {force: true});
execFileSync('zip', ['-q', '-r', output, '.'], {cwd: folder});
console.log(output);
