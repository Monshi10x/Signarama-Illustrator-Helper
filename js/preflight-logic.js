(function(root, factory) {
  const api = factory();
  if(typeof module !== 'undefined' && module.exports) module.exports = api;
  if(root) root.PreflightLogic = api;
})(typeof self !== 'undefined' ? self : this, function() {
  function distance(a, b) {const x = b[0] - a[0], y = b[1] - a[1]; return Math.sqrt(x * x + y * y);}
  function pointLineDistance(p, a, b) {
    const len = distance(a, b);
    if(!len) return distance(p, a);
    return Math.abs((b[0] - a[0]) * (a[1] - p[1]) - (a[0] - p[0]) * (b[1] - a[1])) / len;
  }
  function splitCurve(c) {
    const mid = (a, b) => [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2];
    const a = mid(c[0], c[1]), b = mid(c[1], c[2]), d = mid(c[2], c[3]);
    const e = mid(a, b), f = mid(b, d), g = mid(e, f);
    return [[c[0], a, e, g], [g, f, d, c[3]]];
  }
  function flattenCurve(curve, tolerance, out, depth) {
    out = out || []; depth = depth || 0;
    if(depth >= 12 || (pointLineDistance(curve[1], curve[0], curve[3]) <= tolerance && pointLineDistance(curve[2], curve[0], curve[3]) <= tolerance)) {
      out.push([curve[0], curve[3]]); return out;
    }
    const halves = splitCurve(curve);
    flattenCurve(halves[0], tolerance, out, depth + 1);
    flattenCurve(halves[1], tolerance, out, depth + 1);
    return out;
  }
  function lineOverlap(a, b, tolerance, minimum) {
    const alen = distance(a.a, a.b), blen = distance(b.a, b.b);
    if(!alen || !blen) return null;
    if(pointLineDistance(b.a, a.a, a.b) > tolerance || pointLineDistance(b.b, a.a, a.b) > tolerance) return null;
    const ux = (a.b[0] - a.a[0]) / alen, uy = (a.b[1] - a.a[1]) / alen;
    const project = p => (p[0] - a.a[0]) * ux + (p[1] - a.a[1]) * uy;
    const b0 = project(b.a), b1 = project(b.b);
    const start = Math.max(0, Math.min(b0, b1)), end = Math.min(alen, Math.max(b0, b1));
    if(end - start < minimum) return null;
    const overlapLength = end - start;
    return {length: overlapLength, kind: (overlapLength >= alen - tolerance && overlapLength >= blen - tolerance) ? 'full duplicate' : 'partial overlap', a: [a.a[0] + ux * start, a.a[1] + uy * start], b: [a.a[0] + ux * end, a.a[1] + uy * end]};
  }
  function prepareSegments(paths, tolerance) {
    const out = [];
    (paths || []).forEach(path => (path.segments || []).forEach((s, segmentIndex) => {
      const curve = [s.a, s.c1, s.c2, s.b];
      flattenCurve(curve, tolerance, []).forEach((line, flatIndex) => out.push({
        a: line[0], b: line[1], pathIndex: path.pathIndex, objectName: path.objectName,
        layerName: path.layerName, segmentIndex, flatIndex
      }));
    }));
    return out;
  }
  function findOverlapsAsync(paths, options) {
    const tolerance = options.tolerancePt, minimum = options.minimumPt, cell = options.gridCellPt;
    const segments = prepareSegments(paths, tolerance);
    const grid = {}, pairs = [], seen = {};
    segments.forEach((s, index) => {
      const minX = Math.min(s.a[0], s.b[0]) - tolerance, maxX = Math.max(s.a[0], s.b[0]) + tolerance;
      const minY = Math.min(s.a[1], s.b[1]) - tolerance, maxY = Math.max(s.a[1], s.b[1]) + tolerance;
      for(let x = Math.floor(minX / cell); x <= Math.floor(maxX / cell); x++) for(let y = Math.floor(minY / cell); y <= Math.floor(maxY / cell); y++) {
        const key = x + ':' + y, bucket = grid[key] || (grid[key] = []);
        bucket.forEach(other => {
          if(segments[other].pathIndex === s.pathIndex && segments[other].segmentIndex === s.segmentIndex) return;
          const pairKey = other < index ? other + ':' + index : index + ':' + other;
          if(!seen[pairKey]) {seen[pairKey] = true; pairs.push([other, index]);}
        });
        bucket.push(index);
      }
    });
    return new Promise(resolve => {
      const issues = []; let cursor = 0;
      function batch() {
        const stop = Math.min(cursor + 600, pairs.length);
        for(; cursor < stop; cursor++) {
          const a = segments[pairs[cursor][0]], b = segments[pairs[cursor][1]];
          const overlap = lineOverlap(a, b, tolerance, minimum);
          if(overlap) issues.push({a: overlap.a, b: overlap.b, lengthPt: overlap.length, kind: overlap.kind, pathIndices: [a.pathIndex, b.pathIndex], objects: [a.objectName, b.objectName], layers: [a.layerName, b.layerName]});
        }
        if(cursor < pairs.length) setTimeout(batch, 0); else resolve({issues, comparedPairs: pairs.length, segmentCount: segments.length});
      }
      batch();
    });
  }
  return {flattenCurve, lineOverlap, findOverlapsAsync};
});
