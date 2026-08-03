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

test('cutfile creation fits the view to the fitted artboard', () => {
  const createStart = hostscript.indexOf('function _srh_cutfileCreateFromSelection');
  const resizeStart = hostscript.indexOf('function _srh_cutfileResizeFilePathTextToArtboard');
  const createFunction = hostscript.slice(createStart, resizeStart);

  assert.match(createFunction, /targetDoc\.artboards\.setActiveArtboardIndex\(0\)/);
  assert.match(createFunction, /app\.executeMenuCommand\('fitin'\)/);
});
