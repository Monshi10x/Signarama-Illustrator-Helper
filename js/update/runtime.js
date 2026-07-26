'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const https = require('https');
const crypto = require('crypto');
const {spawn} = require('child_process');
const semver = require('./semver');
const manifestUtil = require('./manifest');

const PLUGIN_ID = 'com.signarama.helper';
const OWNER = 'Monshi10x';
const REPOSITORY = 'Signarama-Illustrator-Helper';
const API = `https://api.github.com/repos/${OWNER}/${REPOSITORY}/releases`;

function dataRoot() {
  const base = process.platform === 'win32' ? (process.env.APPDATA || os.homedir()) : path.join(os.homedir(), 'Library', 'Application Support');
  return path.join(base, 'Signarama', 'Illustrator Helper');
}
function requiresElevation(installPath, platform) {
  const targetPlatform = platform || process.platform;
  return targetPlatform === 'win32' && /^[a-z]:\\program files(?: \(x86\))?\\/i.test(path.win32.resolve(installPath));
}
function ensureDirectories() {
  const root = dataRoot();
  ['downloads', 'staging', 'backups', 'logs'].forEach((name) => fs.mkdirSync(path.join(root, 'updates', name), {recursive: true}));
  return root;
}
function logEvent(event, details) {
  const dir = path.join(ensureDirectories(), 'updates', 'logs');
  const files = fs.readdirSync(dir).filter((name) => /\.jsonl$/.test(name)).sort().reverse();
  files.slice(10).forEach((name) => {try {fs.unlinkSync(path.join(dir, name));} catch(_ignore) {}});
  fs.appendFileSync(path.join(dir, new Date().toISOString().slice(0, 10) + '.jsonl'), JSON.stringify(Object.assign({timestamp: new Date().toISOString(), event}, details || {})) + '\n');
}
function recentLogEvents(limit) {
  const dir = path.join(ensureDirectories(), 'updates', 'logs');
  const files = fs.readdirSync(dir).filter((name) => /\.jsonl$/.test(name)).sort().reverse();
  const events = [];
  for(const file of files) {
    const lines = fs.readFileSync(path.join(dir, file), 'utf8').split(/\r?\n/).filter(Boolean).reverse();
    for(const line of lines) {try {events.push(JSON.parse(line));} catch(_ignore) {} if(events.length >= (limit || 20)) return events.reverse();}
  }
  return events.reverse();
}
function preferencesPath() {return path.join(ensureDirectories(), 'update-preferences.json');}
function readPreferences() {
  const defaults = {schemaVersion: 1, automaticUpdatesEnabled: true, lastCheckedAt: null, ignoredVersion: null, updateChannel: 'stable'};
  try {return Object.assign(defaults, JSON.parse(fs.readFileSync(preferencesPath(), 'utf8')));} catch(_error) {return defaults;}
}
function writePreferences(value) {
  const target = preferencesPath(), temporary = target + '.tmp';
  fs.writeFileSync(temporary, JSON.stringify(value, null, 2), {mode: 0o600});
  fs.renameSync(temporary, target);
}
function request(url, options, redirects) {
  options = options || {}; redirects = redirects || 0;
  if(!manifestUtil.isAllowedHttpsUrl(url)) return Promise.reject(new Error('Update URL is not permitted'));
  return new Promise((resolve, reject) => {
    const req = https.get(url, {headers: {'User-Agent': 'Signarama-Illustrator-Helper-Updater', Accept: options.accept || 'application/vnd.github+json'}}, (res) => {
      if([301, 302, 303, 307, 308].includes(res.statusCode)) {
        res.resume();
        if(redirects >= 5) return reject(new Error('Too many update redirects'));
        return request(new URL(res.headers.location, url).toString(), options, redirects + 1).then(resolve, reject);
      }
      if(res.statusCode === 403 && res.headers['x-ratelimit-remaining'] === '0') {res.resume(); return reject(new Error('GitHub update service rate limit reached'));}
      if(res.statusCode < 200 || res.statusCode >= 300) {res.resume(); return reject(new Error('Update server returned HTTP ' + res.statusCode));}
      resolve(res);
    });
    req.setTimeout(options.timeout || 15000, () => req.destroy(new Error('Update request timed out')));
    req.on('error', reject);
  });
}
async function json(url) {
  const response = await request(url), chunks = [];
  for await (const chunk of response) chunks.push(chunk);
  return JSON.parse(Buffer.concat(chunks).toString('utf8'));
}
async function checkForUpdate(installedVersion, manual, onActivity) {
  const activity = typeof onActivity === 'function' ? onActivity : () => {};
  const prefs = readPreferences(), now = Date.now();
  activity('Reading update preferences (' + prefs.updateChannel + ' channel).');
  if(!manual && (!prefs.automaticUpdatesEnabled || (prefs.lastCheckedAt && now - Date.parse(prefs.lastCheckedAt) < 86400000))) {activity('Automatic check is not due.'); return {status: 'not-due'};}
  let releases;
  activity('Requesting releases from ' + API + '.');
  try {releases = await json(API); activity('Received ' + releases.length + ' published release' + (releases.length === 1 ? '' : 's') + '.'); if(!releases.length) activity('No GitHub Releases are published. Branch version changes are not downloadable updates.'); logEvent('check-complete', {installedVersion, downloadDomain: 'api.github.com', checkStatus: 'success'});}
  catch(error) {logEvent('check-failed', {installedVersion, downloadDomain: 'api.github.com', checkStatus: 'failed', errorCode: 'CHECK_FAILED', message: error.message}); throw error;}
  prefs.lastCheckedAt = new Date().toISOString(); writePreferences(prefs);
  if(!releases.length) return {status: 'no-releases'};
  for(const release of releases) {
    if(release.draft || (prefs.updateChannel === 'stable' && release.prerelease)) continue;
    const asset = (release.assets || []).find((item) => item.name === 'update.json');
    if(!asset) continue;
    activity('Inspecting update manifest for ' + (release.tag_name || release.name || 'release') + '.');
    const candidate = await json(asset.browser_download_url);
    const checked = manifestUtil.validate(candidate, {pluginId: PLUGIN_ID, pluginType: 'cep', owner: OWNER, repository: REPOSITORY});
    if(!checked.valid) {activity('Ignored an invalid update manifest: ' + checked.errors.join(', ') + '.'); continue;}
    if(!semver.isNewer(candidate.version, installedVersion)) {activity('Version ' + candidate.version + ' is not newer than installed version ' + installedVersion + '.'); continue;}
    if(!manual && prefs.ignoredVersion === candidate.version && !candidate.mandatory) return {status: 'ignored'};
    activity('Update ' + candidate.version + ' is available.');
    return {status: 'available', manifest: candidate};
  }
  activity('No newer compatible release was found.');
  return {status: 'current'};
}
async function downloadUpdate(manifest, onProgress) {
  const checked = manifestUtil.validate(manifest, {pluginId: PLUGIN_ID, pluginType: 'cep', owner: OWNER, repository: REPOSITORY});
  if(!checked.valid) throw new Error('Invalid update manifest: ' + checked.errors.join(', '));
  const target = path.join(ensureDirectories(), 'updates', 'downloads', `plugin-v${manifest.version}.zip`);
  const temporary = target + '.partial';
  const response = await request(manifest.downloadUrl, {accept: 'application/octet-stream', timeout: 30000});
  const hash = crypto.createHash('sha256'), output = fs.createWriteStream(temporary, {mode: 0o600});
  let received = 0;
  try {
    for await (const chunk of response) {received += chunk.length; hash.update(chunk); if(!output.write(chunk)) await new Promise((r) => output.once('drain', r)); if(onProgress) onProgress(received, manifest.packageSize || +(response.headers['content-length'] || 0));}
    await new Promise((resolve, reject) => output.end((error) => error ? reject(error) : resolve()));
    if(manifest.packageSize && received !== manifest.packageSize) throw new Error('Downloaded package size does not match manifest');
    if(hash.digest('hex').toLowerCase() !== manifest.sha256.toLowerCase()) throw new Error('Package verification failed');
    fs.renameSync(temporary, target); logEvent('download-complete', {targetVersion: manifest.version, downloadDomain: new URL(manifest.downloadUrl).hostname, downloadStatus: 'success', checksumResult: 'verified'}); return target;
  } catch(error) {output.destroy(); try {fs.unlinkSync(temporary);} catch(_ignore) {} logEvent('download-failed', {targetVersion: manifest.version, downloadDomain: new URL(manifest.downloadUrl).hostname, downloadStatus: 'failed', checksumResult: /verification/i.test(error.message) ? 'mismatch' : 'not-verified', errorCode: 'DOWNLOAD_FAILED', message: error.message}); throw error;}
}
async function launchUpdater(packagePath, manifest, installPath) {
  const configPath = path.join(ensureDirectories(), 'updates', 'pending-update.json');
  fs.writeFileSync(configPath, JSON.stringify({packagePath, installPath, installedVersion: require('../../package.json').version, targetVersion: manifest.version, pluginId: PLUGIN_ID}, null, 2), {mode: 0o600});
  const updaterPath = path.join(__dirname, process.platform === 'win32' ? 'updater.ps1' : 'updater.js');
  let child, elevated = false;
  if(requiresElevation(installPath)) {
    const quote = (value) => String(value).replace(/'/g, "''");
    const command = `Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File','"${quote(updaterPath)}"','"${quote(configPath)}"') -Verb RunAs -PassThru | Select-Object -ExpandProperty Id`;
    logEvent('updater-elevation-requested', {targetVersion: manifest.version, executable: 'powershell.exe', updaterPath});
    child = spawn('powershell.exe', ['-NoProfile', '-Command', command], {windowsHide: false});
    elevated = true;
    let output = '', errors = '';
    child.stdout.on('data', (chunk) => {output += chunk.toString();});
    child.stderr.on('data', (chunk) => {errors += chunk.toString();});
    await new Promise((resolve, reject) => {
      child.on('error', reject);
      child.on('close', (code) => {
        if(code === 0) resolve();
        else reject(new Error(errors.trim() || 'Windows elevation was cancelled or failed (exit ' + code + ')'));
      });
    }).catch((error) => {logEvent('updater-launch-failed', {errorCode: 'LAUNCH_FAILED', message: error.message}); throw error;});
    logEvent('updater-launched', {targetVersion: manifest.version, executable: 'powershell.exe', updaterPath, elevationRequested: true, processId: +(output.trim()) || null});
  } else {
    child = spawn(process.execPath, [updaterPath, configPath], {detached: true, stdio: 'ignore'});
    child.on('error', (error) => logEvent('updater-launch-failed', {errorCode: 'LAUNCH_FAILED', message: error.message}));
    logEvent('updater-launched', {targetVersion: manifest.version, executable: process.execPath, elevationRequested: false, processId: child.pid || null});
    child.unref();
  }
  return configPath;
}

module.exports = {PLUGIN_ID, OWNER, REPOSITORY, dataRoot, requiresElevation, readPreferences, writePreferences, recentLogEvents, checkForUpdate, downloadUpdate, launchUpdater};
