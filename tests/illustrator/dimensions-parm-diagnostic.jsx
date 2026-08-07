//#target illustrator
(function() {
  var repetitions = 50;
  var originalDoc = app.documents.length ? app.activeDocument : null;
  var originalSelection = [];
  try {if(originalDoc && originalDoc.selection) for(var s = 0; s < originalDoc.selection.length; s++) originalSelection.push(originalDoc.selection[s]);} catch(_eSel) { }
  var testDoc = null;
  var failures = [];
  try {
    testDoc = app.documents.add(DocumentColorSpace.RGB, 600, 600);
    var rect = testDoc.pathItems.rectangle(400, 100, 200, 100);
    var sides = ['TOP', 'BOTTOM', 'LEFT', 'RIGHT'];
    for(var i = 0; i < sides.length; i++) {
      for(var n = 1; n <= repetitions; n++) {
        testDoc.selection = null;
        rect.selected = true;
        var result = _dim_run({side: sides[i], offsetMm: 10, ticLenMm: 2, textPt: 10, strokePt: 1, decimals: 1, labelGapMm: 2, scaleAppearance: 1});
        if(/^Error:/i.test(String(result))) {
          failures.push('side=' + sides[i] + ' | iteration=' + n + ' | ' + result);
          break;
        }
        try {var layer = testDoc.layers.getByName('Dimensions'); layer.remove();} catch(_eLayer) { }
      }
    }
  } catch(e) {
    failures.push('harness | ' + _dim_errorDetails(e, 'diagnosticHarness', {}));
  } finally {
    try {if(testDoc) testDoc.close(SaveOptions.DONOTSAVECHANGES);} catch(_eClose) { }
    try {
      if(originalDoc) {
        originalDoc.activate();
        originalDoc.selection = null;
        for(var r = 0; r < originalSelection.length; r++) try {originalSelection[r].selected = true;} catch(_eRestoreOne) { }
      }
    } catch(_eRestore) { }
  }
  alert(failures.length ? ('Dimensions diagnostic failures:\n' + failures.join('\n')) : ('Dimensions diagnostic passed: ' + (repetitions * 4) + ' single-side runs.'));
})();
