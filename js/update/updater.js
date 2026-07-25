#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const {execFile} = require('child_process');
const {promisify} = require('util');
const execFileAsync = promisify(execFile);
const semver = require('./semver');
const {validateArchiveEntries} = require('./archive');

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
function rootPath() {
  const base = process.platform === 'win32' ? (process.env.APPDATA || os.homedir()) : path.join(os.homedir(), 'Library', 'Application Support');
  return path.join(base, 'Signarama', 'Illustrator Helper', 'updates');
}
function log(event, details) {
  const dir = path.join(rootPath(), 'logs'); fs.mkdirSync(dir, {recursive: true});
  const files = fs.readdirSync(dir).filter((name) => name.endsWith('.jsonl')).sort().reverse();
  files.slice(10).forEach((name) => {try {fs.unlinkSync(path.join(dir, name));} catch(_ignore) {}});
  const file = path.join(dir, new Date().toISOString().slice(0, 10) + '.jsonl');
  fs.appendFileSync(file, JSON.stringify(Object.assign({timestamp: new Date().toISOString(), event}, details || {})) + '\n');
}
async function illustratorRunning() {
  try {
    if(process.platform === 'win32') {
      const result = await execFileAsync('tasklist', ['/FI', 'IMAGENAME eq Illustrator.exe', '/NH']);
      return /Illustrator\.exe/i.test(result.stdout);
    }
    const result = await execFileAsync('pgrep', ['-x', 'Adobe Illustrator']); return !!result.stdout.trim();
  } catch(_error) {return false;}
}
async function waitForIllustrator(timeoutMs) {
  const expires = Date.now() + timeoutMs;
  while(await illustratorRunning()) {if(Date.now() > expires) throw new Error('Illustrator remains open'); await sleep(2000);}
}
async function archiveEntries(packagePath) {
  if(process.platform === 'win32') {
    const script = "$z=[IO.Compression.ZipFile]::OpenRead($args[0]);try{$z.Entries|%{$_.FullName}}finally{$z.Dispose()}";
    const result = await execFileAsync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script, packagePath], {maxBuffer: 10 * 1024 * 1024});
    return result.stdout.split(/\r?\n/).filter(Boolean);
  }
  const result = await execFileAsync('unzip', ['-Z1', packagePath], {maxBuffer: 10 * 1024 * 1024});
  return result.stdout.split(/\r?\n/).filter(Boolean);
}
async function extract(packagePath, destination) {
  fs.mkdirSync(destination, {recursive: true});
  validateArchiveEntries(await archiveEntries(packagePath));
  if(process.platform === 'win32') {
    const script = 'Expand-Archive -LiteralPath $args[0] -DestinationPath $args[1] -Force';
    await execFileAsync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script, packagePath, destination]);
  } else await execFileAsync('unzip', ['-q', packagePath, '-d', destination]);
}
function validateStaging(directory, expectedId, expectedVersion) {
  const manifestPath = path.join(directory, 'CSXS', 'manifest.xml');
  if(!fs.existsSync(manifestPath) || !fs.existsSync(path.join(directory, 'index.html'))) throw new Error('Required plugin files are missing');
  const xml = fs.readFileSync(manifestPath, 'utf8');
  const id = /ExtensionBundleId="([^"]+)"/.exec(xml), version = /ExtensionBundleVersion="([^"]+)"/.exec(xml);
  if(!id || id[1] !== expectedId) throw new Error('Plugin ID does not match');
  if(!version || !semver.parse(version[1]) || version[1] !== expectedVersion) throw new Error('Plugin version does not match');
  return true;
}
async function install(config, hooks) {
  hooks = hooks || {};
  const installPath = path.resolve(config.installPath), packagePath = path.resolve(config.packagePath);
  if(!fs.existsSync(packagePath) || path.extname(packagePath).toLowerCase() !== '.zip') throw new Error('Validated update package is missing');
  if(!fs.existsSync(path.join(installPath, 'CSXS', 'manifest.xml'))) throw new Error('Target is not a CEP plugin installation');
  const root = rootPath(), staging = path.join(root, 'staging', config.targetVersion + '-' + Date.now());
  const backup = path.join(root, 'backups', config.installedVersion);
  fs.mkdirSync(path.dirname(backup), {recursive: true});
  await extract(packagePath, staging); validateStaging(staging, config.pluginId, config.targetVersion);
  if(hooks.afterStaging) await hooks.afterStaging(staging);
  if(fs.existsSync(backup)) fs.rmSync(backup, {recursive: true, force: true});
  fs.renameSync(installPath, backup);
  try {
    if(hooks.beforeReplacement) await hooks.beforeReplacement();
    fs.renameSync(staging, installPath); validateStaging(installPath, config.pluginId, config.targetVersion);
    log('installation-complete', {installedVersion: config.installedVersion, targetVersion: config.targetVersion, backupPath: backup, installationResult: 'success'});
    return {backupPath: backup};
  } catch(error) {
    try {if(fs.existsSync(installPath)) fs.rmSync(installPath, {recursive: true, force: true}); if(fs.existsSync(backup)) fs.renameSync(backup, installPath); log('rollback-complete', {targetVersion: config.targetVersion, rollbackResult: 'success', errorCode: 'INSTALL_FAILED', message: error.message});}
    catch(rollbackError) {log('rollback-failed', {targetVersion: config.targetVersion, rollbackResult: 'failed', errorCode: 'ROLLBACK_FAILED', message: rollbackError.message});}
    throw error;
  }
}
async function main() {
  const configFile = process.argv[2];
  if(!configFile) throw new Error('Missing updater configuration');
  const config = JSON.parse(fs.readFileSync(configFile, 'utf8'));
  if(config.pluginId !== 'com.signarama.helper' || !semver.parse(config.targetVersion) || !semver.parse(config.installedVersion)) throw new Error('Invalid updater configuration');
  log('update-started', {installedVersion: config.installedVersion, targetVersion: config.targetVersion});
  await waitForIllustrator(30 * 60 * 1000); await install(config); fs.unlinkSync(configFile);
}

module.exports = {archiveEntries, extract, validateStaging, install, waitForIllustrator};
if(require.main === module) main().catch((error) => {log('installation-failed', {installationResult: 'failed', errorCode: 'INSTALL_FAILED', message: error.message}); process.exitCode = 1;});
