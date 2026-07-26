(function(root, factory) {
  var api = factory(typeof module === 'object' && module.exports ? require('./semver') : root.SignaramaSemver);
  if(typeof module === 'object' && module.exports) module.exports = api;
  else root.SignaramaUpdateManifest = api;
}(this, function(semver) {
  var ALLOWED_HOSTS = ['github.com', 'api.github.com', 'objects.githubusercontent.com', 'github-releases.githubusercontent.com', 'release-assets.githubusercontent.com'];
  function httpsUrl(value, allowedHosts) {
    try {
      var parsed = new URL(value);
      return parsed.protocol === 'https:' && (allowedHosts || ALLOWED_HOSTS).indexOf(parsed.hostname.toLowerCase()) !== -1;
    } catch(_error) {return false;}
  }
  function validate(input, options) {
    var errors = [], value = input && typeof input === 'object' ? input : {};
    options = options || {};
    if(value.schemaVersion !== 1) errors.push('Unsupported schema version');
    if(!semver.parse(value.version)) errors.push('Invalid or missing version');
    if(['stable', 'beta'].indexOf(value.channel) < 0) errors.push('Invalid channel');
    if(value.pluginType !== (options.pluginType || 'cep')) errors.push('Incorrect plugin type');
    if(value.packageType !== 'zip') errors.push('Unsupported package type');
    if(!/^[a-fA-F0-9]{64}$/.test(value.sha256 || '')) errors.push('Invalid or missing checksum');
    if(!httpsUrl(value.downloadUrl, options.allowedHosts)) errors.push('Invalid download URL');
    else if(options.owner && options.repository) {
      var parsedDownload = new URL(value.downloadUrl);
      var expectedPrefix = '/' + options.owner + '/' + options.repository + '/releases/download/v' + value.version + '/';
      var expectedName = 'signarama-helper-v' + value.version + '.zip';
      if(parsedDownload.hostname !== 'github.com' || parsedDownload.pathname.indexOf(expectedPrefix) !== 0 || parsedDownload.pathname.slice(-expectedName.length) !== expectedName) errors.push('Unexpected release asset');
    }
    if(value.releaseNotesUrl && !httpsUrl(value.releaseNotesUrl, options.allowedHosts)) errors.push('Invalid release notes URL');
    if(value.packageSize != null && (!Number.isSafeInteger(value.packageSize) || value.packageSize < 1)) errors.push('Invalid package size');
    if(value.pluginId !== (options.pluginId || 'com.signarama.helper')) errors.push('Incorrect plugin ID');
    return {valid: errors.length === 0, errors: errors, manifest: value};
  }
  function selectRelease(releases, installedVersion, channel) {
    channel = channel || 'stable';
    return (releases || []).filter(function(release) {
      if(!release || release.draft || !semver.parse(release.version)) return false;
      if(channel === 'stable' && (release.prerelease || semver.parse(release.version).prerelease.length)) return false;
      if(channel === 'beta' && release.channel && ['stable', 'beta'].indexOf(release.channel) < 0) return false;
      return !!release.hasRequiredAsset && semver.isNewer(release.version, installedVersion);
    }).sort(function(a, b) {return semver.compare(b.version, a.version);})[0] || null;
  }
  return {ALLOWED_HOSTS: ALLOWED_HOSTS, validate: validate, selectRelease: selectRelease, isAllowedHttpsUrl: httpsUrl};
}));
