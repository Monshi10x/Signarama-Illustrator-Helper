const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const hostscript = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'hostscript.jsx'), 'utf8');

test('cutfile documents are created with millimetre document units', () => {
  assert.match(hostscript, /docPreset\.units\s*=\s*RulerUnits\.Millimeters/);
  assert.match(hostscript, /addDocument\(startupPreset, docPreset, false\)/);
  assert.doesNotMatch(hostscript, /targetDoc\.rulerUnits\s*=/);
});

test('cutfile artboards are centered on the initial maximum-size canvas', () => {
  assert.match(hostscript, /canvasCenterX\s*=\s*\(Number\(initialArtboard\[0\]\)\s*\+\s*Number\(initialArtboard\[2\]\)\)\s*\/\s*2/);
  assert.match(hostscript, /artboardRect\s*=\s*\[artLeft, artTop, artRight, artBottom\]/);
});

test('the active cutfile tick only writes a file path when it changed', () => {
  const tickStart = hostscript.indexOf('function signarama_helper_cutfile_tickFilePathLabels()');
  const resizeStart = hostscript.indexOf('function _srh_cutfileResizeFilePathTextToArtboard');
  const resizeFunction = hostscript.slice(resizeStart, tickStart);

  assert.match(resizeFunction, /if\(currentContents === filePath\) return false/);
  assert.ok(resizeFunction.indexOf('currentContents === filePath') < resizeFunction.indexOf('tf.contents = filePath'));
});

test('a label-only path update is saved only when the document was already clean', () => {
  const tickStart = hostscript.indexOf('function signarama_helper_cutfile_tickFilePathLabels()');
  const tickEnd = hostscript.indexOf('function signarama_helper_makeRouterCutfile()', tickStart);
  const tickFunction = hostscript.slice(tickStart, tickEnd);

  assert.match(tickFunction, /wasCleanBeforeTick\s*=\s*!!doc\.saved/);
  assert.match(tickFunction, /if\(updated && wasCleanBeforeTick\)/);
  assert.match(tickFunction, /doc\.save\(\)/);
  assert.ok(tickFunction.indexOf('wasCleanBeforeTick = !!doc.saved') < tickFunction.indexOf('_srh_cutfileResizeFilePathTextToArtboard'));
  assert.ok(tickFunction.indexOf('if(updated && wasCleanBeforeTick)') < tickFunction.indexOf('doc.save()'));
});

test('cutfile creation fits the view to the fitted artboard', () => {
  const createStart = hostscript.indexOf('function _srh_cutfileCreateFromSelection');
  const resizeStart = hostscript.indexOf('function _srh_cutfileResizeFilePathTextToArtboard');
  const createFunction = hostscript.slice(createStart, resizeStart);

  assert.match(createFunction, /targetDoc\.artboards\.setActiveArtboardIndex\(0\)/);
  assert.match(createFunction, /app\.executeMenuCommand\('fitin'\)/);
});
