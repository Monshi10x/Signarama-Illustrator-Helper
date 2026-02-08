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

  function cmykKey(c, m, y, k){
    function r(v){ return Math.round(v * 100) / 100; }
    return [r(c), r(m), r(y), r(k)].join(',');
  }

  const colourEditState = {
    lastEdit: null
  };

  function buildDimensionPayload(){
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
      if (tabId === 'tab-colours' && typeof window.refreshColours === 'function') {
        window.refreshColours();
      }
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
    wireLedLayout();
    wireLedDepiction();
    wireColours();

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
        ledOffsetMm: num(($('lightboxLedOffsetMm') && $('lightboxLedOffsetMm').value) || 0),
        addMeasures: !!($('lightboxAddMeasures') && $('lightboxAddMeasures').checked),
        measureOptions: buildDimensionPayload()
      };
      const json = JSON.stringify(payload).replace(/\\/g,'\\\\').replace(/"/g, '\\"');
      callJSX('signarama_helper_createLightbox("' + json + '")', res => res && log(res));
    };

    const createBtnLed = $('btnCreateLightboxWithLedPanel');
    if (createBtnLed) {
      createBtnLed.onclick = () => {
        const payload = {
          widthMm: num(($('lightboxWidthMm') && $('lightboxWidthMm').value) || 0),
          heightMm: num(($('lightboxHeightMm') && $('lightboxHeightMm').value) || 0),
          depthMm: num(($('lightboxDepthMm') && $('lightboxDepthMm').value) || 0),
          type: ($('lightboxType') && $('lightboxType').value) || 'Acrylic',
          supportCount: parseInt((($('lightboxSupportCount') && $('lightboxSupportCount').value) || 0), 10) || 0,
          ledOffsetMm: num(($('lightboxLedOffsetMm') && $('lightboxLedOffsetMm').value) || 0),
          ledWatt: num(($('ledWatt') && $('ledWatt').value) || 0),
          ledWidthMm: num(($('ledWidthMm') && $('ledWidthMm').value) || 0),
          ledHeightMm: num(($('ledHeightMm') && $('ledHeightMm').value) || 0),
          allowanceWmm: num(($('ledAllowanceWmm') && $('ledAllowanceWmm').value) || 0),
          allowanceHmm: num(($('ledAllowanceHmm') && $('ledAllowanceHmm').value) || 0),
          maxLedsInSeries: parseInt((($('ledMaxInSeries') && $('ledMaxInSeries').value) || 0), 10) || 50,
          flipLed: !!($('ledFlip') && $('ledFlip').checked),
          addMeasures: !!($('lightboxAddMeasures') && $('lightboxAddMeasures').checked),
          measureOptions: buildDimensionPayload()
        };
        const json = JSON.stringify(payload).replace(/\\/g,'\\\\').replace(/"/g, '\\"');
        callJSX('signarama_helper_createLightboxWithLedPanel("' + json + '")', res => res && log(res));
      };
    }
  }

  function wireLedLayout(){
    const ledBtn = $('btnDrawLEDs');
    if (!ledBtn) return;
    ledBtn.onclick = () => {
      const payload = {
        ledWatt: num(($('ledWatt') && $('ledWatt').value) || 0),
        ledWidthMm: num(($('ledWidthMm') && $('ledWidthMm').value) || 0),
        ledHeightMm: num(($('ledHeightMm') && $('ledHeightMm').value) || 0),
        allowanceWmm: num(($('ledAllowanceWmm') && $('ledAllowanceWmm').value) || 0),
        allowanceHmm: num(($('ledAllowanceHmm') && $('ledAllowanceHmm').value) || 0),
        maxLedsInSeries: parseInt((($('ledMaxInSeries') && $('ledMaxInSeries').value) || 0), 10) || 50,
        flipLed: !!($('ledFlip') && $('ledFlip').checked),
        layoutWidthMm: num(($('ledLayoutWidthMm') && $('ledLayoutWidthMm').value) || 0),
        layoutHeightMm: num(($('ledLayoutHeightMm') && $('ledLayoutHeightMm').value) || 0),
        depthMm: num(($('lightboxDepthMm') && $('lightboxDepthMm').value) || 0),
        boxWidthMm: num(($('lightboxWidthMm') && $('lightboxWidthMm').value) || 0),
        boxHeightMm: num(($('lightboxHeightMm') && $('lightboxHeightMm').value) || 0)
      };
      const json = JSON.stringify(payload).replace(/\\/g,'\\\\').replace(/"/g, '\\"');
      callJSX('signarama_helper_drawLedLayout("' + json + '")', res => res && log(res));
    };
  }

  function wireLedDepiction(){
    const w = $('ledWidthMm');
    const h = $('ledHeightMm');
    const flip = $('ledFlip');
    const rect = $('ledDepictionRect');
    const line = $('ledDepictionLine');
    const text = $('ledDepictionText');

    function update(){
      const width = num((w && w.value) || 0) || 0;
      const height = num((h && h.value) || 0) || 0;
      const isFlip = !!(flip && flip.checked);
      const dispW = isFlip ? height : width;
      const dispH = isFlip ? width : height;

      if(rect) {
        const maxW = 100;
        const maxH = 40;
        let scale = 1;
        if(dispW > 0 && dispH > 0) {
          scale = Math.min(maxW / dispW, maxH / dispH, 1);
        }
        rect.style.width = Math.max(10, dispW * scale) + 'px';
        rect.style.height = Math.max(6, dispH * scale) + 'px';
      }
      if(line) {
        if(isFlip) {
          line.style.left = '50%';
          line.style.right = 'auto';
          line.style.top = '0';
          line.style.bottom = '0';
          line.style.width = '1px';
          line.style.height = '100%';
        } else {
          line.style.left = '0';
          line.style.right = '0';
          line.style.top = '50%';
          line.style.bottom = 'auto';
          line.style.width = '100%';
          line.style.height = '1px';
        }
      }
      if(text) {
        text.textContent = 'LED ' + dispW + 'mm × ' + dispH + 'mm';
      }
    }

    if(w) w.addEventListener('input', update);
    if(h) h.addEventListener('input', update);
    if(flip) flip.addEventListener('change', update);
    update();
  }

  function wireColours(){
    const refreshBtn = $('btnRefreshColours');
    if (refreshBtn) refreshBtn.onclick = () => refreshColours();
    window.refreshColours = refreshColours;
  }

  function refreshColours(){
    const list = $('coloursList');
    const countEl = $('coloursCount');
    if (!list) return;
    const focused = document.activeElement;
    const focusedMeta = focused && focused.dataset && focused.dataset.focusKey ? {
      focusKey: focused.dataset.focusKey,
      focusHex: focused.dataset.focusHex,
      focusType: focused.dataset.focusType,
      index: focused.dataset.focusIndex
    } : null;
    list.innerHTML = '';
    if (countEl) countEl.textContent = 'Document colours: 0';

    callJSX('signarama_helper_getDocumentColors()', function(res){
      if (!res) { log('Colours: no response from JSX.'); return; }
      let data = null;
      let debug = null;
      try { data = JSON.parse(res); } catch(_e1) {
        try { data = JSON.parse(JSON.parse(res)); } catch(_e2) {
          try { data = Function('return ' + res)(); } catch(_e3) { data = null; }
        }
      }
      if (!data) {
        log('Colours raw response: ' + res);
        return;
      }
      if (data && data.colors && Array.isArray(data.colors)) {
        debug = data.debug || null;
        data = data.colors;
      }
      if (!data || !Array.isArray(data)) return;
      if (countEl) countEl.textContent = 'Document colours: ' + data.length + (debug ? ' (items: ' + (debug.totalItems || 0) + ', scanned: ' + (debug.scanned || 0) + ', path: ' + (debug.pathItems || 0) + ', text: ' + (debug.textFrames || 0) + ', fallback: ' + (debug.fallbackUsed ? 'yes' : 'no') + ')' : '');

      data.forEach((entry) => {
        if (colourEditState.lastEdit && entry.type === colourEditState.lastEdit.type) {
          const expectedHex = cmykToHex(
            colourEditState.lastEdit.c,
            colourEditState.lastEdit.m,
            colourEditState.lastEdit.y,
            colourEditState.lastEdit.k
          );
          const matchesEdit = entry.key === colourEditState.lastEdit.toKey ||
            (entry.hex && entry.hex.toLowerCase() === expectedHex.toLowerCase());
          if (matchesEdit) {
            entry.c = colourEditState.lastEdit.c;
            entry.m = colourEditState.lastEdit.m;
            entry.y = colourEditState.lastEdit.y;
            entry.k = colourEditState.lastEdit.k;
            entry.label = (entry.type || 'fill').toUpperCase() + '  C ' + entry.c + ' M ' + entry.m + ' Y ' + entry.y + ' K ' + entry.k;
            log('Colours: matched lastEdit toKey=' + colourEditState.lastEdit.toKey + ' type=' + (entry.type || '') +
              ' values=' + [entry.c, entry.m, entry.y, entry.k].join(',') + ' hexMatch=' +
              (entry.hex ? entry.hex.toLowerCase() === expectedHex.toLowerCase() : false));
          }
        }
        log('Colours: row values key=' + (entry.key || '') + ' type=' + (entry.type || '') +
          ' cmyk=' + [entry.c, entry.m, entry.y, entry.k].join(',') + ' hex=' + (entry.hex || ''));

        const row = document.createElement('div');
        row.className = 'row';
        row.style.alignItems = 'center';
        row.style.gap = '10px';

        const swatch = document.createElement('div');
        swatch.style.width = '34px';
        swatch.style.height = '34px';
        swatch.style.border = '1px solid #444';
        swatch.style.borderRadius = '6px';
        swatch.style.background = entry.hex || '#000000';

        const label = document.createElement('div');
        label.textContent = entry.label || '';
        label.style.fontSize = '12px';
        label.style.color = '#bcbcbc';

        const cInput = document.createElement('input');
        cInput.type = 'number'; cInput.step = '0.1'; cInput.min = '0'; cInput.max = '100';
        cInput.value = entry.c || 0;
        const mInput = document.createElement('input');
        mInput.type = 'number'; mInput.step = '0.1'; mInput.min = '0'; mInput.max = '100';
        mInput.value = entry.m || 0;
        const yInput = document.createElement('input');
        yInput.type = 'number'; yInput.step = '0.1'; yInput.min = '0'; yInput.max = '100';
        yInput.value = entry.y || 0;
        const kInput = document.createElement('input');
        kInput.type = 'number'; kInput.step = '0.1'; kInput.min = '0'; kInput.max = '100';
        kInput.value = entry.k || 0;

        const inputs = [cInput, mInput, yInput, kInput];
        inputs.forEach(inp => {
          inp.style.width = '35px';
        });

        function applyCmyk() {
          const cRaw = parseFloat(cInput.value);
          const mRaw = parseFloat(mInput.value);
          const yRaw = parseFloat(yInput.value);
          const kRaw = parseFloat(kInput.value);
          const c = isFinite(cRaw) ? cRaw : num(entry.c);
          const m = isFinite(mRaw) ? mRaw : num(entry.m);
          const y = isFinite(yRaw) ? yRaw : num(entry.y);
          const k = isFinite(kRaw) ? kRaw : num(entry.k);
          const payload = JSON.stringify({
            fromKey: entry.key || '',
            fromType: entry.type || '',
            toCmyk: {c, m, y, k}
          }).replace(/\\/g,'\\\\').replace(/"/g, '\\"');
          log('Colours: replace fromKey=' + (entry.key || '') + ' type=' + (entry.type || '') +
            ' to=' + [c, m, y, k].join(',') + ' (inputs=' +
            [cInput.value, mInput.value, yInput.value, kInput.value].join(',') + ')');
          callJSX('signarama_helper_replaceColor("' + payload + '")', () => {
            swatch.style.background = cmykToHex(c, m, y, k);
            colourEditState.lastEdit = {
              type: entry.type || '',
              fromKey: entry.key || '',
              toKey: cmykKey(c, m, y, k),
              c, m, y, k
            };
            log('Colours: refresh after replace toKey=' + colourEditState.lastEdit.toKey);
            refreshColours();
          });
        }

        inputs.forEach((inp, idx) => {
          inp.addEventListener('change', applyCmyk);
          inp.addEventListener('focus', () => {
            try { inp.select(); } catch(_eSel) { }
          });
          inp.addEventListener('keydown', (e) => {
            if(e.key === 'Tab') {
              e.preventDefault();
              const next = inputs[(idx + 1) % inputs.length];
              if(next) { next.focus(); try { next.select(); } catch(_eSel2) { } }
            }
          });
        });

        // Focus restore metadata
        cInput.dataset.focusKey = entry.key || '';
        mInput.dataset.focusKey = entry.key || '';
        yInput.dataset.focusKey = entry.key || '';
        kInput.dataset.focusKey = entry.key || '';
        cInput.dataset.focusType = entry.type || '';
        mInput.dataset.focusType = entry.type || '';
        yInput.dataset.focusType = entry.type || '';
        kInput.dataset.focusType = entry.type || '';
        cInput.dataset.focusHex = entry.hex || '';
        mInput.dataset.focusHex = entry.hex || '';
        yInput.dataset.focusHex = entry.hex || '';
        kInput.dataset.focusHex = entry.hex || '';
        cInput.dataset.focusIndex = "0";
        mInput.dataset.focusIndex = "1";
        yInput.dataset.focusIndex = "2";
        kInput.dataset.focusIndex = "3";

        row.appendChild(swatch);
        row.appendChild(label);
        row.appendChild(cInput);
        row.appendChild(mInput);
        row.appendChild(yInput);
        row.appendChild(kInput);
        list.appendChild(row);
      });

      if (debug) {
        const dbg = document.createElement('div');
        dbg.className = 'small';
        dbg.style.marginTop = '8px';
        dbg.textContent = 'Debug: total=' + (debug.totalItems || 0) +
          ', scanned=' + (debug.scanned || 0) +
          ', path=' + (debug.pathItems || 0) +
          ', text=' + (debug.textFrames || 0) +
          ', docPaths=' + (debug.totalPathItems || 0) +
          ', docText=' + (debug.totalTextFrames || 0) +
          ', fallback=' + (debug.fallbackUsed ? 'yes' : 'no') +
          ', samples=' + (debug.sampleTypes || []).join(', ');
        list.appendChild(dbg);
        log('Colours debug: ' + dbg.textContent);
      }

      if (focusedMeta) {
        let nextFocus = null;
        let focusKey = focusedMeta.focusKey;
        if (
          colourEditState.lastEdit &&
          focusKey &&
          focusedMeta.focusType === colourEditState.lastEdit.type &&
          focusKey === colourEditState.lastEdit.fromKey
        ) {
          focusKey = colourEditState.lastEdit.toKey;
        }
        if (focusedMeta.focusKey) {
          const selectorKey = '[data-focus-key=\"' + focusKey + '\"][data-focus-index=\"' + focusedMeta.index + '\"]';
          nextFocus = list.querySelector(selectorKey);
        }
        if (!nextFocus && focusedMeta.focusHex) {
          const selectorHex = '[data-focus-hex=\"' + focusedMeta.focusHex + '\"][data-focus-index=\"' + focusedMeta.index + '\"]';
          nextFocus = list.querySelector(selectorHex);
        }
        if (nextFocus) {
          nextFocus.focus();
          try { nextFocus.select(); } catch(_eSel3) { }
        }
      }
    });
  }

  function cmykToHex(c, m, y, k) {
    const C = Math.max(0, Math.min(100, c)) / 100;
    const M = Math.max(0, Math.min(100, m)) / 100;
    const Y = Math.max(0, Math.min(100, y)) / 100;
    const K = Math.max(0, Math.min(100, k)) / 100;
    const r = Math.round(255 * (1 - C) * (1 - K));
    const g = Math.round(255 * (1 - M) * (1 - K));
    const b = Math.round(255 * (1 - Y) * (1 - K));
    const h = v => v.toString(16).padStart(2, '0');
    return '#' + h(r) + h(g) + h(b);
  }

  function wireDimensions(){
    const dimClear = $('btnDimClear');
    if (dimClear) dimClear.onclick = () => callJSX('atlas_dimensions_clear()', res => res && logDim(res));

    function runSides(sides){
      const payload = buildDimensionPayload();
      payload.sides = sides;
      const json = JSON.stringify(payload).replace(/\\/g,'\\\\').replace(/"/g, '\\"');
      callJSX('atlas_dimensions_runMulti("' + json + '")', function(res){
        if (res) logDim(res);
      });
    }

    function runLineMeasure(){
      const payload = buildDimensionPayload();
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
