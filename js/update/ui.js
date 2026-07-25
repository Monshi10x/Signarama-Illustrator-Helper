/* global SignaramaUpdaterUI */
(function(root) {
  'use strict';
  function init() {
    if(typeof require !== 'function') return;
    var runtime;
    try {runtime = require('./js/update/runtime');} catch(_error) {return;}
    var packageInfo = require('./package.json'), path = require('path');
    var current = null, downloaded = null;
    function el(id) {return document.getElementById(id);}
    function visible(id, show) {var node = el(id); if(node) node.classList.toggle('hidden', !show);}
    function status(text) {if(el('updateStatus')) el('updateStatus').textContent = text;}
    function actions(available, ready) {
      ['btnUpdateNow', 'btnUpdateLater', 'btnSkipUpdate'].forEach(function(id) {visible(id, available);});
      visible('btnInstallUpdate', ready);
    }
    function showAvailable(manifest) {
      current = manifest; actions(true, false);
      el('updateVersions').textContent = 'Installed ' + packageInfo.version + ' · Available ' + manifest.version;
      var size = manifest.packageSize ? (manifest.packageSize / 1048576).toFixed(1) + ' MB' : 'Size unavailable';
      el('updateDetails').textContent = 'Published ' + new Date(manifest.publishedAt).toLocaleDateString() + ' · ' + size;
      if(manifest.releaseNotesUrl) {var link = document.createElement('a'); link.href = manifest.releaseNotesUrl; link.textContent = ' · View release notes'; link.target = '_blank'; link.rel = 'noopener noreferrer'; el('updateDetails').appendChild(link);}
      visible('updateDetails', true); status('An update is available. Review the release notes before installing major updates.');
    }
    async function check(manual) {
      actions(false, false); status('Checking for updates…'); el('btnCheckUpdates').disabled = true;
      try {
        var result = await runtime.checkForUpdate(packageInfo.version, manual);
        if(result.status === 'available') showAvailable(result.manifest);
        else if(result.status === 'current') status('You are using the latest version.');
        else if(result.status === 'not-due') status('The next automatic update check is not due yet.');
        else status('This update has been skipped. Use Check for Updates to show it again.');
      } catch(error) {status(/rate limit/i.test(error.message) ? 'The update service rate limit was reached. Please try again later.' : 'Unable to reach the update server. The plugin remains available offline.');}
      finally {el('btnCheckUpdates').disabled = false;}
    }
    el('updateVersions').textContent = 'Installed ' + packageInfo.version;
    var prefs = runtime.readPreferences(); el('automaticUpdatesEnabled').checked = prefs.automaticUpdatesEnabled; el('updateChannel').value = prefs.updateChannel;
    function savePrefs() {var value = runtime.readPreferences(); value.automaticUpdatesEnabled = el('automaticUpdatesEnabled').checked; value.updateChannel = el('updateChannel').value; value.ignoredVersion = null; runtime.writePreferences(value);}
    el('automaticUpdatesEnabled').addEventListener('change', savePrefs); el('updateChannel').addEventListener('change', savePrefs);
    el('btnCheckUpdates').addEventListener('click', function() {check(true);});
    el('btnUpdateLater').addEventListener('click', function() {actions(false, false); status('Reminder postponed until the next scheduled check.');});
    el('btnSkipUpdate').addEventListener('click', function() {var value = runtime.readPreferences(); value.ignoredVersion = current.version; runtime.writePreferences(value); actions(false, false); status('Version ' + current.version + ' will be skipped.');});
    el('btnUpdateNow').addEventListener('click', async function() {
      actions(false, false); visible('updateProgress', true); status('Downloading update…');
      try {
        downloaded = await runtime.downloadUpdate(current, function(received, total) {el('updateProgress').value = total ? Math.min(100, received / total * 100) : 0; status('Downloading update… ' + received + (total ? ' / ' + total + ' bytes' : ' bytes'));});
        visible('updateProgress', false); actions(false, true); status('The update is ready to install. Illustrator must restart.');
      } catch(error) {visible('updateProgress', false); status(/verification/i.test(error.message) ? 'Package verification failed. Nothing was installed.' : 'Download failed. Nothing was installed.');}
    });
    el('btnInstallUpdate').addEventListener('click', function() {
      try {var extensionPath = root.__adobe_cep__ ? decodeURI(root.__adobe_cep__.getSystemPath('extension')).replace(/^file:\/\//, '') : path.resolve(__dirname); runtime.launchUpdater(downloaded, current, extensionPath); actions(false, false); status('Updater started. Save your work and close Illustrator; it will not be force-quit.');}
      catch(_error) {status('Installation handoff failed. Nothing was replaced.');}
    });
    setTimeout(function() {check(false);}, 1500);
  }
  root.SignaramaUpdaterUI = {init: init};
  document.addEventListener('DOMContentLoaded', init);
}(this));
