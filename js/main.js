/* global CSInterface, SystemPath */
(function () {
  function $(id){ return document.getElementById(id); }
  function $all(sel){ return Array.prototype.slice.call(document.querySelectorAll(sel)); }

  function log(line){
    const el = $('log');
    if (el) { el.textContent += String(line) + '\n'; el.scrollTop = el.scrollHeight; }
  }
  function logDim(line){
    const el = $('dimensionsLog');
    if (el) { el.textContent += String(line) + '\n'; el.scrollTop = el.scrollHeight; }
  }
  function setStatus(txt){
    const el = $('status');
    if (el) el.textContent = txt || '';
  }
  function num(v){
    var n = parseFloat(v);
    return isFinite(n) ? n : 0;
  }

  if (typeof CSInterface === 'undefined') {
    alert('CSInterface.js NOT loaded. Fix script paths.');
    return;
  }

  const cs = new CSInterface();

  document.addEventListener('DOMContentLoaded', () => {
    wireTabs();
    wireInfoIcons();
    wireActions();

    log('Panel booting…');

    cs.evalScript('app.name', function(name){
      setStatus(name ? ('Connected: ' + name) : 'Not connected');
      log(name ? ('Connected to: ' + name) : 'Could not query app.name');
    });

    loadJSX(() => {
      cs.evalScript('typeof signarama_helper_fitArtboardToArtwork', function (type) {
        log('JSX check: signarama_helper_fitArtboardToArtwork is ' + type);
        if (type !== 'function') log('ERROR: JSX not loaded (check path/case).');
      });
    });
  });

  function wireTabs(){
    const tabs = $all('.tab[data-tab]');
    const panels = $all('[role="tabpanel"]');

    function activate(tabId){
      tabs.forEach(t => t.classList.toggle('active', t.getAttribute('data-tab') === tabId));
      panels.forEach(p => p.classList.toggle('hidden', p.id !== tabId));
    }

    tabs.forEach(t => {
      t.addEventListener('click', () => activate(t.getAttribute('data-tab')));
    });

    // default
    if (tabs.length) activate(tabs[0].getAttribute('data-tab'));
  }

  function wireInfoIcons(){
    // Use native tooltip via title attribute.
    $all('.info[data-tip]').forEach((el) => {
      el.setAttribute('title', el.getAttribute('data-tip'));
      el.addEventListener('click', (e) => {
        // Don't trigger parent button.
        e.preventDefault();
        e.stopPropagation();
      });
    });
  }

  function wireActions(){
    const fit = $('btnFitArtboard');
    if (fit) fit.onclick = () => callJSX('signarama_helper_fitArtboardToArtwork()', res => res && log(res));

    const ab = $('btnArtboardPerItem');
    if (ab) ab.onclick = () => callJSX('signarama_helper_createArtboardsFromSelection()', res => res && log(res));

    const a4 = $('btnCopyOutlineScaleA4');
    if (a4) a4.onclick = () => callJSX('signarama_helper_duplicateOutlineScaleA4()', res => res && log(res));

    const preset = $('bleedPreset');
    if (preset) {
      preset.onchange = () => {
        const v = preset.value;
        if (v === 'acm') {
          $('bleedTop').value = 20; $('bleedRight').value = 20; $('bleedBottom').value = 20; $('bleedLeft').value = 20;
        } else if (v === 'print') {
          $('bleedTop').value = 3; $('bleedRight').value = 3; $('bleedBottom').value = 3; $('bleedLeft').value = 3;
        } else if (v === 'none') {
          $('bleedTop').value = 0; $('bleedRight').value = 0; $('bleedBottom').value = 0; $('bleedLeft').value = 0;
        }
      };
    }

    const applyBleed = $('btnApplyBleed');
    if (applyBleed) {
      applyBleed.onclick = () => {
        const t = num($('bleedTop').value);
        const r = num($('bleedRight').value);
        const b = num($('bleedBottom').value);
        const l = num($('bleedLeft').value);
        callJSX('signarama_helper_applyBleed(' + [t,l,b,r].join(',') + ')', res => res && log(res));
      };
    }


    const applyPathBleed = $('btnApplyPathBleed');
    if (applyPathBleed) {
      applyPathBleed.onclick = () => {
        const amt = num(($('pathBleedAmount') && $('pathBleedAmount').value) || 0);
        const cut = !!(($('pathBleedCutline') && $('pathBleedCutline').checked));
        const payload = JSON.stringify({ offsetMm: amt, createCutline: cut });
        const safe = payload.replace(/\\/g,'\\\\').replace(/'/g, "\\'");
        callJSX("signarama_helper_applyPathBleed('" + safe + "')", res => res && log(res));
      };
    }

    const path = $('btnAddPathText');
    if (path) path.onclick = () => callJSX('signarama_helper_addFilePathTextToArtboards()', res => res && log(res));

    const outlineAll = $('btnOutlineAllText');
    if (outlineAll) outlineAll.onclick = () => callJSX('signarama_helper_outlineAllText()', res => res && log(res));

    const setFillsStrokes = $('btnSetFillsStrokes');
    if (setFillsStrokes) setFillsStrokes.onclick = () => callJSX('signarama_helper_setAllFillsStrokes()', res => res && log(res));

    wireDimensions();
    wireLightbox();

    const clear = $('btnClearLog');
    if (clear) clear.onclick = () => { const el = $('log'); if (el) el.textContent=''; };
  }

  function wireLightbox(){
    const createBtn = $('btnCreateLightbox');
    if (!createBtn) return;
    createBtn.onclick = () => {
      const payload = {
        widthMm: num(($('lightboxWidthMm') && $('lightboxWidthMm').value) || 0),
        heightMm: num(($('lightboxHeightMm') && $('lightboxHeightMm').value) || 0),
        depthMm: num(($('lightboxDepthMm') && $('lightboxDepthMm').value) || 0),
        type: ($('lightboxType') && $('lightboxType').value) || 'Acrylic',
        supportCount: parseInt((($('lightboxSupportCount') && $('lightboxSupportCount').value) || 0), 10) || 0,
        ledOffsetMm: num(($('lightboxLedOffsetMm') && $('lightboxLedOffsetMm').value) || 0)
      };
      const json = JSON.stringify(payload).replace(/\\/g,'\\\\').replace(/"/g, '\\"');
      callJSX('signarama_helper_createLightbox("' + json + '")', res => res && log(res));
    };
  }

  function wireDimensions(){
    const dimClear = $('btnDimClear');
    if (dimClear) dimClear.onclick = () => callJSX('atlas_dimensions_clear()', res => res && logDim(res));

    function buildPayload(){
      return {
        offsetMm: num(($('offsetMm') && $('offsetMm').value) || 0),
        ticLenMm: num(($('ticLenMm') && $('ticLenMm').value) || 0),
        textPt: num(($('textPt') && $('textPt').value) || 0),
        strokePt: num(($('strokePt') && $('strokePt').value) || 0),
        decimals: parseInt((($('decimals') && $('decimals').value) || 0), 10),
        labelGapMm: num(($('labelGapMm') && $('labelGapMm').value) || 0),
        measureClippedContent: !!($('measureClippedContent') && $('measureClippedContent').checked),
        arrowheadSizePt: num(($('arrowheadSizePt') && $('arrowheadSizePt').value) || 0),
        includeArrowhead: !!($('includeArrowhead') && $('includeArrowhead').checked),
        textColor: ($('textColor') && $('textColor').value) || '#ff00ff',
        lineColor: ($('lineColor') && $('lineColor').value) || '#000000',
        scaleAppearance: num(($('scaleAppearance') && $('scaleAppearance').value) || 100) / 100
      };
    }

    function runSides(sides){
      const payload = buildPayload();
      payload.sides = sides;
      const json = JSON.stringify(payload).replace(/\\/g,'\\\\').replace(/"/g, '\\"');
      callJSX('atlas_dimensions_runMulti("' + json + '")', function(res){
        if (res) logDim(res);
      });
    }

    function runLineMeasure(){
      const payload = buildPayload();
      const json = JSON.stringify(payload).replace(/\\/g,'\\\\').replace(/"/g, '\\"');
      callJSX('atlas_dimensions_runLine("' + json + '")', function(res){
        if (res) logDim(res);
      });
    }

    const map = {
      btnTL: () => runSides(['TOP', 'LEFT']),
      btnT: () => runSides(['TOP']),
      btnRT: () => runSides(['RIGHT', 'TOP']),
      btnL: () => runSides(['LEFT']),
      btnCenterText: () => runSides(['CENTER_TEXT']),
      btnR: () => runSides(['RIGHT']),
      btnBL: () => runSides(['BOTTOM', 'LEFT']),
      btnB: () => runSides(['BOTTOM']),
      btnRB: () => runSides(['RIGHT', 'BOTTOM']),
      btnLineMeasure: () => runLineMeasure()
    };

    Object.keys(map).forEach(id => {
      const el = $(id);
      if (el) el.onclick = map[id];
    });
  }

  function loadJSX(done){
    try {
      var extDir = cs.getSystemPath(SystemPath.EXTENSION).replace(/\\/g, '/');
      var cmd = '$.evalFile("' + extDir + '/jsx/hostscript.jsx")';
      cs.evalScript(cmd, function(){
        log('JSX loaded.');
        if (done) done();
      });
    } catch (e) {
      log('Failed to load JSX: ' + (e && e.message ? e.message : e));
    }
  }

  function callJSX(fnCall, cb) {
    var wrapped = '(function(){try{return ' + fnCall + ' }catch(e){return "Error: " + e}})()';
    cs.evalScript(wrapped, function (res) { if (cb) cb(res); });
  }
})();
