'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const vm = require('vm');
const {execFileSync} = require('child_process');
const semver = require('../js/update/semver');
const manifest = require('../js/update/manifest');
const archive = require('../js/update/archive');
const updater = require('../js/update/updater');
const updateRuntime = require('../js/update/runtime');

const validManifest = (overrides) => Object.assign({schemaVersion: 1, channel: 'stable', version: '1.6.2', publishedAt: '2026-07-25T08:00:00Z', pluginType: 'cep', pluginId: 'com.signarama.helper', packageType: 'zip', downloadUrl: 'https://github.com/Monshi10x/Signarama-Illustrator-Helper/releases/download/v1.6.2/signarama-helper-v1.6.2.zip', sha256: 'a'.repeat(64), packageSize: 100, releaseNotesUrl: 'https://github.com/Monshi10x/Signarama-Illustrator-Helper/releases/tag/v1.6.2', mandatory: false}, overrides);

test('semantic versions compare correctly', () => {
  assert.equal(semver.compare('1.0.0', '1.0.1'), -1); assert.equal(semver.compare('1.0.0', '1.0.0'), 0);
  assert.equal(semver.compare('2.0.0', '1.9.9'), 1); assert.equal(semver.compare('1.10.0', '1.9.0'), 1);
  assert.equal(semver.compare('2.0.0-beta.1', '2.0.0'), -1); assert.equal(semver.compare('2.0.0-beta.2', '2.0.0-beta.1'), 1);
  assert.equal(semver.parse('01.0.0'), null); assert.throws(() => semver.compare('bad', '1.0.0'));
});

test('update manifests are strictly validated', () => {
  assert.equal(manifest.validate(validManifest()).valid, true);
  for(const change of [{version: undefined}, {sha256: undefined}, {downloadUrl: 'http://github.com/x'}, {pluginType: 'uxp'}, {schemaVersion: 2}]) assert.equal(manifest.validate(validManifest(change)).valid, false);
  assert.equal(manifest.validate(validManifest(), {owner: 'Monshi10x', repository: 'Signarama-Illustrator-Helper'}).valid, true);
  assert.equal(manifest.validate(validManifest({downloadUrl: 'https://github.com/other/repo/releases/download/v1.6.2/signarama-helper-v1.6.2.zip'}), {owner: 'Monshi10x', repository: 'Signarama-Illustrator-Helper'}).valid, false);
});

test('GitHub release redirect hosts are permitted without allowing lookalike domains', () => {
  assert.equal(manifest.isAllowedHttpsUrl('https://release-assets.githubusercontent.com/github-production-release-asset/file?token=temporary'), true);
  assert.equal(manifest.isAllowedHttpsUrl('https://release-assets.githubusercontent.com.evil.example/file'), false);
  assert.equal(manifest.isAllowedHttpsUrl('http://release-assets.githubusercontent.com/file'), false);
});

test('system-wide Windows CEP installations require elevation', () => {
  assert.equal(updateRuntime.requiresElevation('C:\\Program Files (x86)\\Common Files\\Adobe\\CEP\\extensions\\Signarama-Illustrator-Helper', 'win32'), true);
  assert.equal(updateRuntime.requiresElevation('C:\\Program Files\\Adobe\\CEP\\extensions\\Signarama-Illustrator-Helper', 'win32'), true);
  assert.equal(updateRuntime.requiresElevation('C:\\Users\\designer\\AppData\\Roaming\\Adobe\\CEP\\extensions\\Signarama-Illustrator-Helper', 'win32'), false);
  assert.equal(updateRuntime.requiresElevation('/Library/Application Support/Adobe/CEP/extensions/Signarama-Illustrator-Helper', 'darwin'), false);
});

test('release selection respects stable and beta channels', () => {
  const releases = [{version: '3.0.0', draft: true, hasRequiredAsset: true}, {version: '2.0.0-beta.1', prerelease: true, hasRequiredAsset: true}, {version: '1.5.0', hasRequiredAsset: false}, {version: '1.4.0', hasRequiredAsset: true}, {version: '0.9.0', hasRequiredAsset: true}];
  assert.equal(manifest.selectRelease(releases, '1.0.0', 'stable').version, '1.4.0');
  assert.equal(manifest.selectRelease(releases, '1.0.0', 'beta').version, '2.0.0-beta.1');
  assert.equal(manifest.selectRelease(releases, '2.0.0', 'beta'), null);
});

test('SHA-256 accepts matching data and distinguishes mismatches and missing files', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'srh-hash-')), file = path.join(dir, 'package.zip'); fs.writeFileSync(file, 'safe');
  const actual = crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
  assert.equal(actual, crypto.createHash('sha256').update('safe').digest('hex')); assert.notEqual(actual, '0'.repeat(64)); assert.throws(() => fs.readFileSync(path.join(dir, 'missing')));
});

test('archive path validation blocks traversal and absolute paths', () => {
  assert.equal(archive.validateArchiveEntries(['CSXS/manifest.xml', 'index.html']), true);
  assert.throws(() => archive.validateArchiveEntries(['../outside'])); assert.throws(() => archive.validateArchiveEntries(['C:\\outside'])); assert.throws(() => archive.validateArchiveEntries([]));
});

test('staged package validation checks manifest, ID, and version', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'srh-stage-')); fs.mkdirSync(path.join(dir, 'CSXS')); fs.writeFileSync(path.join(dir, 'index.html'), 'ok');
  fs.writeFileSync(path.join(dir, 'CSXS', 'manifest.xml'), '<ExtensionManifest ExtensionBundleId="com.signarama.helper" ExtensionBundleVersion="2.0.0"/>');
  assert.equal(updater.validateStaging(dir, 'com.signarama.helper', '2.0.0'), true);
  assert.throws(() => updater.validateStaging(dir, 'wrong.id', '2.0.0')); fs.unlinkSync(path.join(dir, 'CSXS', 'manifest.xml')); assert.throws(() => updater.validateStaging(dir, 'com.signarama.helper', '2.0.0'));
});

test('failed replacement restores the existing installation and external settings', async () => {
  if(process.platform === 'win32') return;
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'srh-rollback-')), install = path.join(dir, 'plugin'), source = path.join(dir, 'source'), zip = path.join(dir, 'update.zip'), settings = path.join(dir, 'settings.json');
  for(const folder of [install, source]) {fs.mkdirSync(path.join(folder, 'CSXS'), {recursive: true}); fs.writeFileSync(path.join(folder, 'index.html'), folder === install ? 'old' : 'new');}
  fs.writeFileSync(path.join(install, 'CSXS', 'manifest.xml'), '<ExtensionManifest ExtensionBundleId="com.signarama.helper" ExtensionBundleVersion="1.0.0"/>');
  fs.writeFileSync(path.join(source, 'CSXS', 'manifest.xml'), '<ExtensionManifest ExtensionBundleId="com.signarama.helper" ExtensionBundleVersion="2.0.0"/>'); fs.writeFileSync(settings, '{"preserved":true}');
  execFileSync('zip', ['-q', '-r', zip, '.'], {cwd: source});
  await assert.rejects(updater.install({installPath: install, packagePath: zip, installedVersion: '1.0.0', targetVersion: '2.0.0', pluginId: 'com.signarama.helper'}, {beforeReplacement: () => {throw new Error('simulated replacement failure');}}));
  assert.equal(fs.readFileSync(path.join(install, 'index.html'), 'utf8'), 'old'); assert.equal(fs.readFileSync(settings, 'utf8'), '{"preserved":true}');
});

test('update UI loads modules from the decoded CEP extension path and shows update notification', async () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'js', 'update', 'ui.js'), 'utf8');
  const listeners = {}, required = []; let scheduled;
  const elements = new Proxy({}, {get(target, id) {
    if(!target[id]) target[id] = {addEventListener() {}, classList: {toggle(name, hidden) {if(name === 'hidden') this.hidden = hidden;}, hidden: true}, appendChild() {}, textContent: '', value: '', checked: false, disabled: false};
    return target[id];
  }});
  const runtime = {
    readPreferences: () => ({automaticUpdatesEnabled: true, updateChannel: 'stable'}),
    writePreferences() {},
    recentLogEvents: () => [],
    checkForUpdate: async () => ({status: 'available', manifest: {version: '1.0.1', publishedAt: '2026-07-25T08:00:00Z', packageSize: 100}})
  };
  const context = {
    document: {getElementById: (id) => elements[id], addEventListener: (name, callback) => {listeners[name] = callback;}},
    console: {log() {}, error() {}},
    process: {platform: 'win32'},
    __dirname: '.',
    __adobe_cep__: {getSystemPath: () => 'file:///C:/Program%20Files/Adobe/CEP/extensions/Signarama-Illustrator-Helper'},
    setTimeout(callback) {scheduled = callback;},
    require(request) {
      required.push(request);
      if(request === 'path') return path.win32;
      if(/runtime\.js$/.test(request)) return runtime;
      if(/package\.json$/.test(request)) return {version: '1.0.0'};
      throw new Error('Unexpected require: ' + request);
    }
  };
  vm.runInNewContext(source, context); listeners.DOMContentLoaded();
  assert.ok(required.includes('C:\\Program Files\\Adobe\\CEP\\extensions\\Signarama-Illustrator-Helper\\js\\update\\runtime.js'));
  assert.ok(required.includes('C:\\Program Files\\Adobe\\CEP\\extensions\\Signarama-Illustrator-Helper\\package.json'));
  scheduled();
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(elements.updateNotification.classList.hidden, false);
});
