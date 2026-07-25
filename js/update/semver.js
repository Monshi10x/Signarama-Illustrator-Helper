(function(root, factory) {
  var api = factory();
  if(typeof module === 'object' && module.exports) module.exports = api;
  else root.SignaramaSemver = api;
}(this, function() {
  var RE = /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$/;
  function parse(value) {
    var match = RE.exec(String(value || '').replace(/^v/, ''));
    if(!match) return null;
    return {raw: String(value), major: +match[1], minor: +match[2], patch: +match[3], prerelease: match[4] ? match[4].split('.') : []};
  }
  function compareIdentifiers(a, b) {
    var an = /^\d+$/.test(a), bn = /^\d+$/.test(b);
    if(an && bn) return (+a === +b) ? 0 : (+a > +b ? 1 : -1);
    if(an !== bn) return an ? -1 : 1;
    return a === b ? 0 : (a > b ? 1 : -1);
  }
  function compare(a, b) {
    a = parse(a); b = parse(b);
    if(!a || !b) throw new Error('Invalid semantic version');
    for(var key of ['major', 'minor', 'patch']) if(a[key] !== b[key]) return a[key] > b[key] ? 1 : -1;
    if(!a.prerelease.length || !b.prerelease.length) return a.prerelease.length === b.prerelease.length ? 0 : (a.prerelease.length ? -1 : 1);
    for(var i = 0; i < Math.max(a.prerelease.length, b.prerelease.length); i++) {
      if(a.prerelease[i] == null) return -1;
      if(b.prerelease[i] == null) return 1;
      var result = compareIdentifiers(a.prerelease[i], b.prerelease[i]);
      if(result) return result;
    }
    return 0;
  }
  return {parse: parse, compare: compare, isNewer: function(candidate, installed) {return compare(candidate, installed) > 0;}};
}));
