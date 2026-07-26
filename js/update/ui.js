/* global SignaramaUpdaterUI */
(function(root) {
  'use strict';
  function el(id) {return document.getElementById(id);}
  function log(message, level) {
    var output = el('updateDevLog');
    var text = '[' + new Date().toLocaleTimeString() + '] ' + (level || 'INFO') + '  ' + message;
    if(output) {
      var lines = output.textContent ? output.textContent.split('\n') : [];
      lines.push(text); output.textContent = lines.slice(-200).join('\n'); output.scrollTop = output.scrollHeight;
    }
    if(root.console) {(level === 'ERROR' && root.console.error ? root.console.error : root.console.log).call(root.console, '[Updater] ' + message);}
  }
  function extensionDirectory(path) {
    var directory = root.__adobe_cep__ ? root.__adobe_cep__.getSystemPath('extension') : __dirname;
    directory = decodeURI(directory);
    if(/^file:\/\//i.test(directory)) directory = process.platform === 'win32' ? directory.replace(/^file:\/\/\/?/i, '') : directory.replace(/^file:\/\//i, '');
    return path.resolve(directory);
  }
  function init() {
    if(el('btnClearUpdateLog')) el('btnClearUpdateLog').addEventListener('click', function() {if(el('updateDevLog')) el('updateDevLog').textContent = ''; log('Developer log cleared.');});
    log('Update panel initialising.');
    if(typeof require !== 'function') {log('Node.js integration is unavailable; the updater cannot start.', 'ERROR'); return;}
    var runtime, packageInfo, path, extensionPath;
    try {
      path = require('path');
      extensionPath = extensionDirectory(path);
      runtime = require(path.join(extensionPath, 'js', 'update', 'runtime.js'));
      packageInfo = require(path.join(extensionPath, 'package.json'));
      log('Update runtime loaded from ' + extensionPath + '.');
    } catch(error) {log('Could not load the update runtime: ' + (error && error.message ? error.message : String(error)), 'ERROR'); return;}
    var current = null, downloaded = null;
    function visible(id, show) {var node = el(id); if(node) node.classList.toggle('hidden', !show);}
    function status(text) {if(el('updateStatus')) el('updateStatus').textContent = text;}
    function actions(available, ready) {
      ['btnUpdateNow', 'btnUpdateLater', 'btnSkipUpdate'].forEach(function(id) {visible(id, available);});
      visible('btnInstallUpdate', ready);
    }
    function showAvailable(manifest) {
      current = manifest; actions(true, false);
      visible('updateNotification', true);
      el('updateVersions').textContent = 'Installed ' + packageInfo.version + ' · Available ' + manifest.version;
      var size = manifest.packageSize ? (manifest.packageSize / 1048576).toFixed(1) + ' MB' : 'Size unavailable';
      el('updateDetails').textContent = 'Published ' + new Date(manifest.publishedAt).toLocaleDateString() + ' · ' + size;
      if(manifest.releaseNotesUrl) {var link = document.createElement('a'); link.href = manifest.releaseNotesUrl; link.textContent = ' · View release notes'; link.target = '_blank'; link.rel = 'noopener noreferrer'; el('updateDetails').appendChild(link);}
      visible('updateDetails', true); status('An update is available. Review the release notes before installing major updates.');
    }
    async function check(manual) {
      log((manual ? 'Manual' : 'Automatic') + ' update check started for version ' + packageInfo.version + '.');
      actions(false, false); status('Checking for updates…'); el('btnCheckUpdates').disabled = true;
      try {
        var result = await runtime.checkForUpdate(packageInfo.version, manual, function(message) {log(message);});
        if(result.status === 'available') showAvailable(result.manifest);
        else if(result.status === 'current') {visible('updateNotification', false); status('You are using the latest version.');}
        else if(result.status === 'no-releases') status('No published GitHub Release is available. A branch version or tag alone cannot be installed.');
        else if(result.status === 'not-due') status('The next automatic update check is not due yet.');
        else status('This update has been skipped. Use Check for Updates to show it again.');
        log('Update check finished with status: ' + result.status + '.');
      } catch(error) {log('Update check failed: ' + (error && error.stack ? error.stack : String(error)), 'ERROR'); status(/rate limit/i.test(error.message) ? 'The update service rate limit was reached. Please try again later.' : 'Unable to reach the update server. The plugin remains available offline.');}
      finally {el('btnCheckUpdates').disabled = false;}
    }
    el('updateVersions').textContent = 'Installed ' + packageInfo.version;
    var prefs = runtime.readPreferences(); el('automaticUpdatesEnabled').checked = prefs.automaticUpdatesEnabled; el('updateChannel').value = prefs.updateChannel;
    function savePrefs() {var value = runtime.readPreferences(); value.automaticUpdatesEnabled = el('automaticUpdatesEnabled').checked; value.updateChannel = el('updateChannel').value; value.ignoredVersion = null; runtime.writePreferences(value);}
    el('automaticUpdatesEnabled').addEventListener('change', savePrefs); el('updateChannel').addEventListener('change', savePrefs);
    el('btnCheckUpdates').addEventListener('click', function() {check(true);});
    el('btnUpdateLater').addEventListener('click', function() {actions(false, false); status('Reminder postponed until the next scheduled check.');});
    el('btnSkipUpdate').addEventListener('click', function() {var value = runtime.readPreferences(); value.ignoredVersion = current.version; runtime.writePreferences(value); actions(false, false); visible('updateNotification', false); status('Version ' + current.version + ' will be skipped.');});
    el('btnUpdateNow').addEventListener('click', async function() {
      log('Download requested for version ' + current.version + '.');
      actions(false, false); visible('updateProgress', true); status('Downloading update…');
      try {
        downloaded = await runtime.downloadUpdate(current, function(received, total) {el('updateProgress').value = total ? Math.min(100, received / total * 100) : 0; status('Downloading update… ' + received + (total ? ' / ' + total + ' bytes' : ' bytes'));});
        visible('updateProgress', false); actions(false, true); status('The update is ready to install. Illustrator must restart.'); log('Download completed and passed SHA-256 verification.');
      } catch(error) {log('Update download failed: ' + (error && error.stack ? error.stack : String(error)), 'ERROR'); visible('updateProgress', false); status(/verification/i.test(error.message) ? 'Package verification failed. Nothing was installed.' : 'Download failed. Nothing was installed.');}
    });
    el('btnInstallUpdate').addEventListener('click', function() {
      try {log('Handing the verified package to the installer.'); runtime.launchUpdater(downloaded, current, extensionPath); actions(false, false); status('Updater started. Save your work and close Illustrator; it will not be force-quit.'); log('Installer started; waiting for Illustrator to close.');}
      catch(error) {log('Installation handoff failed: ' + (error && error.stack ? error.stack : String(error)), 'ERROR'); status('Installation handoff failed. Nothing was replaced.');}
    });
    setTimeout(function() {check(false);}, 1500);
  }
  root.SignaramaUpdaterUI = {init: init};
  document.addEventListener('DOMContentLoaded', init);
}(this));
