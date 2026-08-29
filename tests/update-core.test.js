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

test('Windows installer displays each operator-facing update phase', () => {
  const script = fs.readFileSync(path.join(__dirname, '..', 'js', 'update', 'updater.ps1'), 'utf8');
  assert.match(script, /Waiting for user to close Illustrator\. Do not close this PowerShell window\./);
  assert.match(script, /Installing update\.\.\./);
  assert.match(script, /Update installed\. You can now reopen Illustrator\./);
  assert.match(script, /Update failed:/);
});

test('colour scan ignores group containers and exposes chunk progress', () => {
  const script = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'hostscript.jsx'), 'utf8');
  assert.match(script, /item\.typename !== "GroupItem"/);
  assert.match(script, /signarama_helper_stepDocumentColorScan/);
  assert.match(script, /position:state\.position, total:state\.total/);
  assert.doesNotMatch(script, /if\(it\.typename === "GroupItem"\) \{\s*_srh_walkPageItems\(it, cb\)/);
});

test('colour rows support copy, paste, and deferred rich-black repair', () => {
  const script = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
  assert.match(script, /copyButton\.textContent = 'Copy'/);
  assert.match(script, /pasteButton\.textContent = 'Paste'/);
  assert.match(script, /Colour values pasted\. Click Apply to update the document\./);
  assert.match(script, /Rich black values populated \('/);
  assert.doesNotMatch(script, /applyValues\(\(\) => \{showToast\('Rich black updated/);
});

test('colour settings control the rich-black target and paste compatibility', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  const script = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
  for(const id of ['richBlackC', 'richBlackM', 'richBlackY', 'richBlackK']) assert.match(html, new RegExp('id="' + id + '"'));
  assert.match(script, /getRichBlackTarget\(\)/);
  assert.match(script, /btn\.dataset\.colourMode !== copiedColourValues\.mode/);
  assert.match(script, /label\.style\.flex = '0 0 46px'/);
});

test('lightboxes use the visible view center and Proofs exposes fixed fetch hosts', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  const main = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
  const host = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'hostscript.jsx'), 'utf8');
  assert.match(host, /var viewCenter = doc\.views\[0\]\.centerPoint/);
  assert.match(html, /<select id="corebridgeProxyBaseUrl">/);
  assert.match(html, /value="http:\/\/localhost:8080"/);
  assert.match(html, /value="https:\/\/signschedulerapp\.ts\.r\.appspot\.com" selected/);
  assert.match(html, /<span>Proof Template<\/span>/);
  assert.match(html, /data-template-filename="PROOF TEMPLATE - Landscape  v2\.ai"/);
  assert.match(main, /function corebridgeProxyBaseUrl\(\)/);
  assert.match(main, /extensionPath \+ '\/Proof Templates\/' \+ proofTemplateFilename/);
  assert.match(main, /return select \? select\.value : 'https:\/\/signschedulerapp\.ts\.r\.appspot\.com'/);
  assert.doesNotMatch(main, /const corebridgeProxyBaseUrl =/);
});

test('Proofs only pulls data on tab visits and explicit button clicks', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  const main = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
  assert.match(main, /tabId === 'tab-corebridge'[\s\S]*scheduleCorebridgeInitialFetch\(\)/);
  assert.match(main, /executeCorebridgePullData\(\{toastOnSuccess: false, toastOnError: false\}\)/);
  assert.match(main, /corebridgePullData\.onclick = async/);
  assert.doesNotMatch(main, /10 \* 60 \* 1000/);
  assert.doesNotMatch(main, /_eAutoPull/);
  assert.match(main, /corebridgeInitialFetchStarted = false;/);
  assert.doesNotMatch(html, /id="corebridgePullControlsWrap"/);
  assert.match(html, /id="btnCopyOutlineScaleA4"[\s\S]*id="btnCorebridgePullData"[\s\S]*<\/div>\s*<div id="corebridgeFetchStatus"/);
});

test('Scripts tab supports bundled files, selected files, and pasted code', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  const main = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
  const host = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'hostscript.jsx'), 'utf8');
  assert.match(html, /data-tab="tab-scripts"/);
  for(const id of ['predefinedScriptsList', 'btnRunScriptFile', 'scriptCode', 'btnRunScriptCode']) assert.match(html, new RegExp('id="' + id + '"'));
  assert.match(main, /signarama_helper_listPredefinedScripts\(\"/);
  assert.match(main, /Resolved scripts folder:/);
  assert.match(main, /Folder entries:/);
  assert.match(host, /var _srh_hostScriptFolderPath = \(function\(\)/);
  assert.match(host, /requestedPath \|\| \(_srh_hostScriptFolderPath \+ '\/scripts'\)/);
  assert.match(host, /function _srh_scriptListResponse\(/);
  assert.match(host, /if\(\/\\\.\(jsx\|js\)\$\/i\.test\(entryName\)\) files\.push\(entries\[e\]\)/);
  assert.doesNotMatch(host, /String\(entries\[e\]\.typename\) === 'File'/);
  assert.doesNotMatch(host, /function _srh_predefinedScriptsFolder\(\) \{[\s\S]{0,200}new File\(\$\.fileName\)/);
  assert.match(host, /function signarama_helper_chooseAndRunScriptFile\(\)/);
  assert.match(host, /function signarama_helper_runScriptCode\(source\)/);
  assert.equal(fs.existsSync(path.join(__dirname, '..', 'jsx', 'scripts', 'Select All Artwork.jsx')), true);
});

test('LED layouts use viewport center, letter LEDs stay on their offset paths, and lightbox measures exclude strokes', () => {
  const host = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'hostscript.jsx'), 'utf8');
  const letterLayout = host.slice(host.indexOf('function signarama_helper_drawLetterLayout'), host.indexOf('function _srh_addLightboxMeasures'));
  assert.match(host, /function _srh_getViewportCenter/);
  assert.match(host, /var viewCenter2 = _getViewCenter\(doc\)/);
  assert.doesNotMatch(letterLayout, /\.translate\(/);
  assert.match(letterLayout, /Samples already use the offset paths' document coordinates/);
  assert.match(host, /Lightbox dimensions describe the path geometry, never the stroke extents/);
  assert.match(host, /try \{b = item\.geometricBounds;\}/);
  assert.match(host, /working\.applyEffect\(fx1\)/);
  assert.match(host, /_srh_letterExpandItems\(doc, working\)/);
  assert.doesNotMatch(host, /wrapper\.applyEffect\(fx1\)/);
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

test('Proof creation reuses already-fetched rows after selecting a different item', () => {
  const main = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
  assert.match(main, /function useCachedCorebridgeDataForCriteria\(criteria\)/);
  assert.match(main, /corebridgeLastAllData\.filter\(\(row\) =>/);
  assert.match(main, /useCachedCorebridgeDataForCriteria\(getCorebridgeCriteriaFromFields\(\)\)/);
  assert.match(main, /\(corebridgeCriteriaChanged\(criteriaNow\) \|\| !corebridgeHasFetchedData\) && !useCachedCorebridgeDataForCriteria\(criteriaNow\)/);
  assert.match(main, /if\(!corebridgeLastSecondaryFetchResults\)[\s\S]{0,300}executeCorebridgeSecondaryFetches/);
});

test('Corebridge flash arrows layer remains unlocked during proof workflows', () => {
  const host = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'hostscript.jsx'), 'utf8');
  const start = host.indexOf('function _srh_corebridge_getArrowLayer');
  const end = host.indexOf('function _srh_corebridge_findTextFrameByName', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const arrowLayerFunctions = host.slice(start, end);
  assert.match(arrowLayerFunctions, /layer\.locked = false/);
  assert.doesNotMatch(arrowLayerFunctions, /(?:layer|arrowLayer)\.locked = true/);
});
