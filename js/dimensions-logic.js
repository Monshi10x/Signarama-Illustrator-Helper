(function(root, factory) {
  var api = factory();
  if(typeof module === 'object' && module.exports) module.exports = api;
  if(root) root.DimensionLogic = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  function finite(value, fallback, options) {
    var n = Number(value);
    if(!Number.isFinite(n)) return fallback;
    if(options && options.positive && !(n > 0)) return fallback;
    if(options && options.nonNegative && n < 0) return fallback;
    return n;
  }

  function normalizeRotationDegrees(value) {
    var angle = finite(value, 0);
    while(angle > 180) angle -= 360;
    while(angle <= -180) angle += 360;
    return angle;
  }

  function normalizePayload(input) {
    var value = input || {};
    return {
      offsetMm: finite(value.offsetMm, 10),
      ticLenMm: finite(value.ticLenMm, 2, {nonNegative: true}),
      textPt: finite(value.textPt, 10, {positive: true}),
      strokePt: finite(value.strokePt, 1, {positive: true}),
      decimals: Math.max(0, Math.floor(finite(value.decimals, 0, {nonNegative: true}))),
      labelGapMm: finite(value.labelGapMm, 0),
      arrowheadSizePt: finite(value.arrowheadSizePt, 0, {nonNegative: true}),
      areaApproximationStep: finite(value.areaApproximationStep, 10, {positive: true}),
      scaleAppearance: finite(value.scaleAppearance, 1, {positive: true}),
      measureIncludeStroke: !!value.measureIncludeStroke,
      measureClippedContent: !!value.measureClippedContent,
      includeArrowhead: !!value.includeArrowhead,
      showAreaApproximation: !!value.showAreaApproximation,
      textColor: value.textColor || '#000000',
      lineColor: value.lineColor || '#000000'
    };
  }

  function inspectMultiResult(side, result, completedMeasures) {
    var text = String(result || '');
    return /^Error:/i.test(text) ? {
      ok: false,
      message: 'Error: combined dimensions stopped at side=' + side + ' after ' + completedMeasures + ' completed measure' + (completedMeasures === 1 ? '' : 's') + '. ' + text.replace(/^Error:\s*/i, '')
    } : {ok: true};
  }

  return {finite: finite, normalizeRotationDegrees: normalizeRotationDegrees, normalizePayload: normalizePayload, inspectMultiResult: inspectMultiResult};
});
