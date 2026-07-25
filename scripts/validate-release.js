#!/usr/bin/env node
'use strict';
const fs = require('fs');
const pkg = require('../package.json');
const semver = require('../js/update/semver');
const xml = fs.readFileSync('CSXS/manifest.xml', 'utf8');
if(!semver.parse(pkg.version)) throw new Error('package.json version is not semantic');
const bundle = /ExtensionBundleVersion="([^"]+)"/.exec(xml);
const extension = /<Extension Id="com\.signarama\.helper" Version="([^"]+)"/.exec(xml);
if(!bundle || !extension || bundle[1] !== pkg.version || extension[1] !== pkg.version) throw new Error('Manifest versions do not match package.json');
if(!/ExtensionBundleId="com\.signarama\.helper"/.test(xml)) throw new Error('Unexpected plugin ID');
console.log(`Validated com.signarama.helper v${pkg.version}`);
