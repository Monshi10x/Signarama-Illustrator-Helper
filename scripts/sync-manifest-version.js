#!/usr/bin/env node
'use strict';

const fs = require('fs');
const pkg = require('../package.json');
const semver = require('../js/update/semver');

if(!semver.parse(pkg.version)) throw new Error('package.json version is not semantic');

const manifestPath = 'CSXS/manifest.xml';
const original = fs.readFileSync(manifestPath, 'utf8');
if(!/ExtensionBundleVersion="[^"]+"/.test(original) || !/<Extension Id="com\.signarama\.helper" Version="[^"]+"\/?>/.test(original)) {
  throw new Error('Could not locate CEP manifest version attributes');
}
const updated = original
  .replace(/ExtensionBundleVersion="[^"]+"/, `ExtensionBundleVersion="${pkg.version}"`)
  .replace(/(<Extension Id="com\.signarama\.helper" Version=")[^"]+("\/?>)/, `$1${pkg.version}$2`);

fs.writeFileSync(manifestPath, updated);
console.log(`Synchronized CEP manifest to v${pkg.version}`);
