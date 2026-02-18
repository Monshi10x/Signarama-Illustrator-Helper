/* global CSInterface, SystemPath */
(function() {
  function $(id) {return document.getElementById(id);}
  function $all(sel) {return Array.prototype.slice.call(document.querySelectorAll(sel));}

  function log(line) {
    const el = $('log');
    if(el) {el.textContent += String(line) + '\n'; el.scrollTop = el.scrollHeight;}
  }
  function logDim(line) {
    const el = $('dimensionsLog');
    if(el) {el.textContent += String(line) + '\n'; el.scrollTop = el.scrollHeight;}
  }
  function setStatus(txt) {
    const el = $('status');
    if(el) el.textContent = txt || '';
  }
  function num(v) {
    var n = parseFloat(v);
    return isFinite(n) ? n : 0;
  }
  function ensureToastUi() {
    if($('srhToastRoot')) return;
    const style = document.createElement('style');
    style.id = 'srhToastStyles';
    style.textContent = [
      '#srhToastRoot{position:fixed;top:14px;right:14px;z-index:2147483000;display:flex;flex-direction:column;gap:10px;pointer-events:none;}',
      '.srh-toast{min-width:250px;max-width:420px;background:#141414;color:#f2f2f2;border:1px solid #3a3a3a;border-radius:12px;box-shadow:inset 0px 1px 2px rgba(255,255,255,0.12),5px 8px 10px rgba(0,0,0,0.45),0 0px 1px rgba(0,0,0,0.6);overflow:hidden;pointer-events:auto;opacity:0;transform:translateY(-6px);transition:opacity .14s ease,transform .14s ease;}',
      '.srh-toast.in{opacity:1;transform:translateY(0);}',
      '.srh-toast-head{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:8px 10px 4px 10px;font-size:12px;font-weight:800;}',
      '.srh-toast-title{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}',
      '.srh-toast-close{appearance:none;border:1px solid #4b4b4b;background:#1f1f1f;color:#d9d9d9;border-radius:8px;width:20px;height:20px;line-height:16px;cursor:pointer;padding:0;}',
      '.srh-toast-close:hover{background:#2a2a2a;}',
      '.srh-toast-body{padding:0 10px 9px 10px;font-size:12px;line-height:1.35;word-break:break-word;}',
      '.srh-toast-progressWrap{height:4px;background:#222;}',
      '.srh-toast-progress{height:100%;width:100%;transform-origin:left center;background:#5a5a5a;}',
      '.srh-toast.info{border-left:3px solid #3b82f6;}',
      '.srh-toast.success{border-left:3px solid #22c55e;}',
      '.srh-toast.warn{border-left:3px solid #f59e0b;}',
      '.srh-toast.error{border-left:3px solid #dc2626;}'
    ].join('');
    document.head.appendChild(style);

    const root = document.createElement('div');
    root.id = 'srhToastRoot';
    document.body.appendChild(root);
  }
  function showToast(message, options) {
    ensureToastUi();
    const opts = options || {};
    const type = opts.type || 'info';
    const duration = typeof opts.duration === 'number' ? Math.max(400, opts.duration) : 5000;
    const title = opts.title || (type === 'error' ? 'Operation failed' : 'Operation finished');
    const root = $('srhToastRoot');
    if(!root) return;

    const toast = document.createElement('div');
    toast.className = 'srh-toast ' + type;
    const head = document.createElement('div');
    head.className = 'srh-toast-head';
    const titleEl = document.createElement('div');
    titleEl.className = 'srh-toast-title';
    titleEl.textContent = title;
    const close = document.createElement('button');
    close.type = 'button';
    close.className = 'srh-toast-close';
    close.textContent = 'x';
    const body = document.createElement('div');
    body.className = 'srh-toast-body';
    body.textContent = String(message || title);
    const pWrap = document.createElement('div');
    pWrap.className = 'srh-toast-progressWrap';
    const pBar = document.createElement('div');
    pBar.className = 'srh-toast-progress';

    head.appendChild(titleEl);
    head.appendChild(close);
    pWrap.appendChild(pBar);
    toast.appendChild(head);
    toast.appendChild(body);
    toast.appendChild(pWrap);
    root.prepend(toast);

    let raf = null;
    const start = performance.now();
    let closed = false;
    function finish() {
      if(closed) return;
      closed = true;
      if(raf) cancelAnimationFrame(raf);
      toast.classList.remove('in');
      setTimeout(() => {if(toast.isConnected) toast.remove();}, 180);
    }
    function tick(now) {
      if(closed) return;
      const elapsed = now - start;
      const frac = Math.min(1, elapsed / duration);
      pBar.style.transform = 'scaleX(' + (1 - frac) + ')';
      if(frac >= 1) {
        finish();
        return;
      }
      raf = requestAnimationFrame(tick);
    }
    close.addEventListener('click', (e) => {e.stopPropagation(); finish();});
    requestAnimationFrame(() => {
      toast.classList.add('in');
      raf = requestAnimationFrame(tick);
    });
  }
  function notifyOperationResult(res, options) {
    const opts = options || {};
    if(opts.showToast === false) return;
    const text = String(res || '').trim();
    const isError = /^error:/i.test(text);
    const isWarn = !isError && (/^no\b/i.test(text) || /\bno\s+(selection|document|items?|artboards?)\b/i.test(text));
    const type = isError ? 'error' : (isWarn ? 'warn' : 'success');
    const title = opts.toastTitle || (isError ? 'Operation failed' : 'Operation finished');
    const message = text || opts.emptyMessage || title;
    showToast(message, {type: type, title: title, duration: opts.toastDuration || 5000});
  }
  function runButtonJsxOperation(fnCall, options) {
    const opts = options || {};
    callJSX(fnCall, function(res) {
      if(opts.logFn && res) opts.logFn(res);
      if(opts.onResult) opts.onResult(res);
      notifyOperationResult(res, opts);
    });
  }

  function cmykKey(c, m, y, k) {
    function r(v) {return Math.round(v * 100) / 100;}
    return [r(c), r(m), r(y), r(k)].join(',');
  }

  function rgbKey(r, g, b) {
    function r2(v) {return Math.round(v * 100) / 100;}
    return [r2(r), r2(g), r2(b)].join(',');
  }

  const colourEditState = {
    lastEdit: null,
    mode: 'CMYK'
  };
  let lightboxMeasureLiveTimer = null;
  let lightboxMeasureLiveInFlight = false;
  let isLargeArtboard = false;
  let refreshLightboxArtboardScaleNotice = null;
  let activateTabFn = null;
  let persistSettingsTimer = null;
  let isApplyingPanelSettings = false;
  let schedulePanelSettingsSave = null;

  function jsxEscapeDoubleQuoted(text) {
    return String(text).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  }

  function collectPanelSettings() {
    const settings = {};
    $all('input[id], select[id], textarea[id]').forEach((el) => {
      if(!el || !el.id) return;
      const type = (el.type || '').toLowerCase();
      if(type === 'checkbox' || type === 'radio') settings[el.id] = !!el.checked;
      else settings[el.id] = el.value;
    });
    const activeTab = document.querySelector('.tab[data-tab].active');
    if(activeTab) settings.__activeTab = activeTab.getAttribute('data-tab') || '';
    return settings;
  }

  function applyPanelSettings(settings) {
    if(!settings || typeof settings !== 'object') return;
    isApplyingPanelSettings = true;
    try {
      Object.keys(settings).forEach((key) => {
        if(key === '__activeTab') return;
        const el = $(key);
        if(!el) return;
        const type = (el.type || '').toLowerCase();
        const value = settings[key];
        if(type === 'checkbox' || type === 'radio') el.checked = !!value;
        else if(value != null) el.value = String(value);
        try {el.dispatchEvent(new Event('input', {bubbles: true}));} catch(_eInp) { }
        try {el.dispatchEvent(new Event('change', {bubbles: true}));} catch(_eChg) { }
      });
      if(settings.__activeTab && activateTabFn) activateTabFn(String(settings.__activeTab));
    } finally {
      isApplyingPanelSettings = false;
    }
  }

  function persistPanelSettings() {
    if(isApplyingPanelSettings) return;
    const payload = collectPanelSettings();
    const json = JSON.stringify(payload);
    const safe = jsxEscapeDoubleQuoted(json);
    callJSX('signarama_helper_panelSettingsSave("' + safe + '")', function(res) {
      if(res && String(res).indexOf('Error:') === 0) log('Settings save failed: ' + res);
    });
  }

  function setupPanelSettingsPersistence() {
    function scheduleSave() {
      if(isApplyingPanelSettings) return;
      if(persistSettingsTimer) clearTimeout(persistSettingsTimer);
      persistSettingsTimer = setTimeout(persistPanelSettings, 120);
    }
    schedulePanelSettingsSave = scheduleSave;

    $all('input[id], select[id], textarea[id]').forEach((el) => {
      el.addEventListener('input', scheduleSave);
      el.addEventListener('change', scheduleSave);
    });
    $all('.tab[data-tab]').forEach((el) => {
      el.addEventListener('click', scheduleSave);
    });
  }

  function loadPanelSettings(done) {
    callJSX('signarama_helper_panelSettingsLoad()', function(res) {
      const txt = String(res || '');
      if(txt && txt !== 'NO_SETTINGS') {
        try {
          applyPanelSettings(JSON.parse(txt));
          log('Panel settings loaded.');
        } catch(e) {
          log('Panel settings parse failed: ' + (e && e.message ? e.message : e));
        }
      }
      if(done) done();
    });
  }

  function buildDimensionPayload() {
    return {
      offsetMm: num(($('offsetMm') && $('offsetMm').value) || 0),
      ticLenMm: num(($('ticLenMm') && $('ticLenMm').value) || 0),
      textPt: num(($('textPt') && $('textPt').value) || 0),
      strokePt: num(($('strokePt') && $('strokePt').value) || 0),
      decimals: parseInt((($('decimals') && $('decimals').value) || 0), 10),
      labelGapMm: num(($('labelGapMm') && $('labelGapMm').value) || 0),
      measureIncludeStroke: !!($('measureIncludeStroke') && $('measureIncludeStroke').checked),
      measureClippedContent: !!($('measureClippedContent') && $('measureClippedContent').checked),
      arrowheadSizePt: num(($('arrowheadSizePt') && $('arrowheadSizePt').value) || 0),
      includeArrowhead: !!($('includeArrowhead') && $('includeArrowhead').checked),
      textColor: ($('textColor') && $('textColor').value) || '#ff00ff',
      lineColor: ($('lineColor') && $('lineColor').value) || '#000000',
      scaleAppearance: num(($('scaleAppearance') && $('scaleAppearance').value) || 100) / 100
    };
  }

  if(typeof CSInterface === 'undefined') {
    alert('CSInterface.js NOT loaded. Fix script paths.');
    return;
  }

  const cs = new CSInterface();

  document.addEventListener('DOMContentLoaded', () => {
    wireTabs();
    wireInfoIcons();
    wireActions();

    log('Panel booting…');

    cs.evalScript('app.name', function(name) {
      setStatus(name ? ('Connected: ' + name) : 'Not connected');
      log(name ? ('Connected to: ' + name) : 'Could not query app.name');
    });

    loadJSX(() => {
      cs.evalScript('typeof signarama_helper_fitArtboardToArtwork', function(type) {
        log('JSX check: signarama_helper_fitArtboardToArtwork is ' + type);
        if(type !== 'function') log('ERROR: JSX not loaded (check path/case).');
        loadPanelSettings(function() {
          setupPanelSettingsPersistence();
          if(typeof refreshLightboxArtboardScaleNotice === 'function') refreshLightboxArtboardScaleNotice();
        });
      });
    });
  });

  function wireTabs() {
    const tabs = $all('.tab[data-tab]');
    const panels = $all('[role="tabpanel"]');

    function activate(tabId) {
      tabs.forEach(t => t.classList.toggle('active', t.getAttribute('data-tab') === tabId));
      panels.forEach(p => p.classList.toggle('hidden', p.id !== tabId));
      if(tabId === 'tab-colours' && typeof window.refreshColours === 'function') {
        window.refreshColours();
      }
      if(tabId === 'tab-lightbox' && typeof refreshLightboxArtboardScaleNotice === 'function') {
        refreshLightboxArtboardScaleNotice();
      }
    }
    activateTabFn = activate;

    tabs.forEach(t => {
      t.addEventListener('click', () => activate(t.getAttribute('data-tab')));
    });

    // default
    if(tabs.length) activate(tabs[0].getAttribute('data-tab'));
  }

  function wireInfoIcons() {
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

  function wireActions() {
    const fit = $('btnFitArtboard');
    if(fit) fit.onclick = () => runButtonJsxOperation('signarama_helper_fitArtboardToArtwork()', {logFn: log, toastTitle: 'Fit artboard'});

    const ab = $('btnArtboardPerItem');
    if(ab) ab.onclick = () => runButtonJsxOperation('signarama_helper_createArtboardsFromSelection()', {logFn: log, toastTitle: 'Artboard per selection'});

    const a4 = $('btnCopyOutlineScaleA4');
    if(a4) a4.onclick = () => runButtonJsxOperation('signarama_helper_duplicateOutlineScaleA4()', {logFn: log, toastTitle: 'Scale artwork for proof'});

    const preset = $('bleedPreset');
    if(preset) {
      preset.onchange = () => {
        const v = preset.value;
        if(v === 'acm') {
          $('bleedTop').value = 20; $('bleedRight').value = 20; $('bleedBottom').value = 20; $('bleedLeft').value = 20;
        } else if(v === 'print') {
          $('bleedTop').value = 3; $('bleedRight').value = 3; $('bleedBottom').value = 3; $('bleedLeft').value = 3;
        } else if(v === 'none') {
          $('bleedTop').value = 0; $('bleedRight').value = 0; $('bleedBottom').value = 0; $('bleedLeft').value = 0;
        }
      };
    }

    const applyBleed = $('btnApplyBleed');
    if(applyBleed) {
      applyBleed.onclick = () => {
        const t = num($('bleedTop').value);
        const r = num($('bleedRight').value);
        const b = num($('bleedBottom').value);
        const l = num($('bleedLeft').value);
        const keepOriginal = !!(($('bleedKeepOriginal') && $('bleedKeepOriginal').checked));
        const excludeClipped = !!(($('bleedExcludeClippedContent') && $('bleedExcludeClippedContent').checked));
        const expandArtboards = !!(($('bleedExpandArtboards') && $('bleedExpandArtboards').checked));
        runButtonJsxOperation('signarama_helper_applyBleed(' + [t, l, b, r, excludeClipped, keepOriginal, expandArtboards].join(',') + ')', {logFn: log, toastTitle: 'Apply bleed'});
      };
    }


    const applyPathBleed = $('btnApplyPathBleed');
    if(applyPathBleed) {
      applyPathBleed.onclick = () => {
        const amt = num(($('pathBleedAmount') && $('pathBleedAmount').value) || 0);
        const cut = !!(($('pathBleedCutline') && $('pathBleedCutline').checked));
        const outlineText = !!(($('pathBleedOutlineText') && $('pathBleedOutlineText').checked));
        const outlineStroke = !!(($('pathBleedOutlineStroke') && $('pathBleedOutlineStroke').checked));
        const autoWeld = !!(($('pathBleedAutoWeld') && $('pathBleedAutoWeld').checked));
        const payload = JSON.stringify({
          offsetMm: amt,
          createCutline: cut,
          outlineText: outlineText,
          outlineStroke: outlineStroke,
          autoWeld: autoWeld
        });
        const safe = payload.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
        runButtonJsxOperation("signarama_helper_applyPathBleed('" + safe + "')", {logFn: log, toastTitle: 'Apply path bleed'});
      };
    }

    const path = $('btnAddPathText');
    if(path) path.onclick = () => runButtonJsxOperation('signarama_helper_addFilePathTextToArtboards()', {logFn: log, toastTitle: 'Add path labels'});

    const outlineAll = $('btnOutlineAllText');
    if(outlineAll) outlineAll.onclick = () => runButtonJsxOperation('signarama_helper_outlineAllText()', {logFn: log, toastTitle: 'Outline all text'});

    const setFillsStrokes = $('btnSetFillsStrokes');
    if(setFillsStrokes) setFillsStrokes.onclick = () => runButtonJsxOperation('signarama_helper_setAllFillsStrokes()', {logFn: log, toastTitle: 'Set fills/strokes'});

    wireDimensions();
    wireLightbox();
    wireLedLayout();
    wireLedDepiction();
    wireColours();

    const clear = $('btnClearLog');
    if(clear) clear.onclick = () => {const el = $('log'); if(el) el.textContent = ''; showToast('Console cleared.', {type: 'info', title: 'Log', duration: 2500});};
  }

  function wireLightbox() {
    function refreshArtboardScaleNotice() {
      callJSX('app.documents.length ? app.activeDocument.scaleFactor : 1', sfRawDoc => {
        const sf = Number(sfRawDoc);
        const sfText = isFinite(sf) && sf > 0 ? sf : 1;
        const invScale = 1 / sfText;
        const large = sfText >= 10;
        isLargeArtboard = large;
        if(typeof window !== 'undefined') window.isLargeArtboard = large;
        const notice = $('lightboxArtboardScaleNotice');
        const icon = $('lightboxArtboardScaleIcon');
        const text = $('lightboxArtboardScaleText');
        if(!notice || !icon || !text) return;
        log('app.activeDocument.scaleFactor: ' + sfText);
        if(large) {
          notice.classList.add('large');
          icon.textContent = '!';
          text.textContent = 'Large artboard detected (scaleFactor=' + sfText + '). Dimension/lightbox geometry uses x' + invScale.toFixed(4) + ' (1/' + sfText + ').';
        } else {
          notice.classList.remove('large');
          icon.textContent = 'OK';
          text.textContent = 'Standard artboard scale detected (scaleFactor=1). Dimension/lightbox geometry uses x1.0000.';
        }
      });
    }

    function updateSupportSpacingInfo() {
      const out = $('lightboxSupportSpacingInfo');
      if(!out) return;
      const boxWidthMm = num(($('lightboxWidthMm') && $('lightboxWidthMm').value) || 0);
      const supports = parseInt((($('lightboxSupportCount') && $('lightboxSupportCount').value) || 0), 10) || 0;
      const supportWidthMm = 25;
      if(!(boxWidthMm > 0) || supports < 1) {
        out.textContent = 'Support spacing: -';
        return;
      }
      const clearGap = (boxWidthMm - (supports * supportWidthMm)) / (supports + 1);
      const centerGap = clearGap + supportWidthMm;
      if(!(clearGap > 0)) {
        out.textContent = 'Support spacing: insufficient width';
        return;
      }
      out.textContent = 'Support spacing: ' + clearGap.toFixed(1) + 'mm clear (' + centerGap.toFixed(1) + 'mm centers)';
    }

    function runLightboxMeasureLiveTick(force) {
      if(lightboxMeasureLiveInFlight) return;
      lightboxMeasureLiveInFlight = true;
      const payload = {
        force: !!force,
        measureOptions: buildDimensionPayload(),
        isLargeArtboard: isLargeArtboard
      };
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      callJSX('signarama_helper_updateLightboxMeasures("' + json + '")', res => {
        lightboxMeasureLiveInFlight = false;
        if(res && res !== 'No live measure changes.' && res !== 'No live lightbox found.') log(res);
      });
    }

    function stopLightboxMeasureLive() {
      if(lightboxMeasureLiveTimer) {
        clearInterval(lightboxMeasureLiveTimer);
        lightboxMeasureLiveTimer = null;
      }
    }

    function startLightboxMeasureLive() {
      if(lightboxMeasureLiveTimer) return;
      lightboxMeasureLiveTimer = setInterval(() => {
        const liveCkb = $('lightboxUpdateMeasuresLive');
        if(!liveCkb || !liveCkb.checked) {
          stopLightboxMeasureLive();
          return;
        }
        runLightboxMeasureLiveTick(false);
      }, 800);
    }

    const liveCkb = $('lightboxUpdateMeasuresLive');
    if(liveCkb) {
      liveCkb.onchange = () => {
        if(liveCkb.checked) {
          runLightboxMeasureLiveTick(true);
          startLightboxMeasureLive();
        } else {
          stopLightboxMeasureLive();
        }
      };
    }
    const widthField = $('lightboxWidthMm');
    const supportsField = $('lightboxSupportCount');
    if(widthField) widthField.addEventListener('input', updateSupportSpacingInfo);
    if(supportsField) supportsField.addEventListener('input', updateSupportSpacingInfo);
    refreshLightboxArtboardScaleNotice = refreshArtboardScaleNotice;
    updateSupportSpacingInfo();
    refreshArtboardScaleNotice();

    const createBtn = $('btnCreateLightbox');
    if(!createBtn) return;
    createBtn.onclick = () => {
      refreshArtboardScaleNotice();
      const payload = {
        widthMm: num(($('lightboxWidthMm') && $('lightboxWidthMm').value) || 0),
        heightMm: num(($('lightboxHeightMm') && $('lightboxHeightMm').value) || 0),
        depthMm: num(($('lightboxDepthMm') && $('lightboxDepthMm').value) || 0),
        type: ($('lightboxType') && $('lightboxType').value) || 'Acrylic',
        supportCount: parseInt((($('lightboxSupportCount') && $('lightboxSupportCount').value) || 0), 10) || 0,
        ledOffsetMm: num(($('lightboxLedOffsetMm') && $('lightboxLedOffsetMm').value) || 0),
        addMeasures: !!($('lightboxAddMeasures') && $('lightboxAddMeasures').checked),
        updateMeasuresLive: !!($('lightboxUpdateMeasuresLive') && $('lightboxUpdateMeasuresLive').checked),
        measureOptions: buildDimensionPayload(),
        isLargeArtboard: isLargeArtboard
      };
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      runButtonJsxOperation('signarama_helper_createLightbox("' + json + '")', {
        logFn: log,
        toastTitle: 'Create lightbox',
        onResult: () => {
          if(payload.updateMeasuresLive && payload.addMeasures) {
            runLightboxMeasureLiveTick(true);
            startLightboxMeasureLive();
          }
        }
      });
    };

    const createBtnLed = $('btnCreateLightboxWithLedPanel');
    if(createBtnLed) {
      createBtnLed.onclick = () => {
        refreshArtboardScaleNotice();
        const payload = {
          widthMm: num(($('lightboxWidthMm') && $('lightboxWidthMm').value) || 0),
          heightMm: num(($('lightboxHeightMm') && $('lightboxHeightMm').value) || 0),
          depthMm: num(($('lightboxDepthMm') && $('lightboxDepthMm').value) || 0),
          type: ($('lightboxType') && $('lightboxType').value) || 'Acrylic',
          supportCount: parseInt((($('lightboxSupportCount') && $('lightboxSupportCount').value) || 0), 10) || 0,
          ledOffsetMm: num(($('lightboxLedOffsetMm') && $('lightboxLedOffsetMm').value) || 0),
          ledWatt: num(($('ledWatt') && $('ledWatt').value) || 0),
          ledCode: (($('ledCode') && $('ledCode').value) || '').trim(),
          ledVoltage: num(($('ledVoltage') && $('ledVoltage').value) || 0),
          ledWidthMm: num(($('ledWidthMm') && $('ledWidthMm').value) || 0),
          ledHeightMm: num(($('ledHeightMm') && $('ledHeightMm').value) || 0),
          allowanceWmm: num(($('ledAllowanceWmm') && $('ledAllowanceWmm').value) || 0),
          allowanceHmm: num(($('ledAllowanceHmm') && $('ledAllowanceHmm').value) || 0),
          maxLedsInSeries: parseInt((($('ledMaxInSeries') && $('ledMaxInSeries').value) || 0), 10) || 50,
          flipLed: !!($('ledFlip') && $('ledFlip').checked),
          addMeasures: !!($('lightboxAddMeasures') && $('lightboxAddMeasures').checked),
          updateMeasuresLive: !!($('lightboxUpdateMeasuresLive') && $('lightboxUpdateMeasuresLive').checked),
          measureOptions: buildDimensionPayload(),
          isLargeArtboard: isLargeArtboard
        };
        const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
        runButtonJsxOperation('signarama_helper_createLightboxWithLedPanel("' + json + '")', {
          logFn: log,
          toastTitle: 'Create lightbox + LED',
          onResult: () => {
            if(payload.updateMeasuresLive && payload.addMeasures) {
              runLightboxMeasureLiveTick(true);
              startLightboxMeasureLive();
            }
          }
        });
      };
    }
  }

  function wireLedLayout() {
    const ledBtn = $('btnDrawLEDs');
    if(!ledBtn) return;
    ledBtn.onclick = () => {
      const payload = {
        ledWatt: num(($('ledWatt') && $('ledWatt').value) || 0),
        ledCode: (($('ledCode') && $('ledCode').value) || '').trim(),
        ledVoltage: num(($('ledVoltage') && $('ledVoltage').value) || 0),
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
        boxHeightMm: num(($('lightboxHeightMm') && $('lightboxHeightMm').value) || 0),
        isLargeArtboard: isLargeArtboard
      };
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      runButtonJsxOperation('signarama_helper_drawLedLayout("' + json + '")', {logFn: log, toastTitle: 'Draw LED layout'});
    };
  }

  function wireLedDepiction() {
    const w = $('ledWidthMm');
    const h = $('ledHeightMm');
    const flip = $('ledFlip');
    const rect = $('ledDepictionRect');
    const line = $('ledDepictionLine');
    const text = $('ledDepictionText');

    function update() {
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

  function wireColours() {
    const refreshBtn = $('btnRefreshColours');
    if(refreshBtn) refreshBtn.onclick = () => refreshColours({showToastOnComplete: true});
    window.refreshColours = refreshColours;
  }

  function refreshColours(options) {
    const showToastOnComplete = !!(options && options.showToastOnComplete);
    const list = $('coloursList');
    const countEl = $('coloursCount');
    if(!list) return;
    const focused = document.activeElement;
    const focusedMeta = focused && focused.dataset && focused.dataset.focusKey ? {
      focusKey: focused.dataset.focusKey,
      focusHex: focused.dataset.focusHex,
      focusType: focused.dataset.focusType,
      index: focused.dataset.focusIndex
    } : null;
    list.innerHTML = '';
    if(countEl) countEl.textContent = 'Document colours: 0';

    callJSX('signarama_helper_getDocumentColorMode()', function(modeRes) {
      const mode = String(modeRes || '').replace(/"/g, '').trim() || 'CMYK';
      colourEditState.mode = mode;
      const banner = $('colourModeBanner');
      if(banner) {
        banner.style.display = mode === 'RGB' ? 'block' : 'none';
        banner.textContent = 'Current document colour mode is RGB';
      }

      callJSX('signarama_helper_getDocumentColors()', function(res) {
        if(!res) {
          log('Colours: no response from JSX.');
          if(showToastOnComplete) showToast('No response from colour scan.', {type: 'warn', title: 'Refresh colours'});
          return;
        }
        let data = null;
        let debug = null;
        try {data = JSON.parse(res);} catch(_e1) {
          try {data = JSON.parse(JSON.parse(res));} catch(_e2) {
            try {data = Function('return ' + res)();} catch(_e3) {data = null;}
          }
        }
        if(!data) {
          log('Colours raw response: ' + res);
          if(showToastOnComplete) showToast('Failed to parse colour scan response.', {type: 'error', title: 'Refresh colours'});
          return;
        }
        if(data && data.colors && Array.isArray(data.colors)) {
          debug = data.debug || null;
          if(data.mode) {
            colourEditState.mode = data.mode;
            if(banner) {
              banner.style.display = data.mode === 'RGB' ? 'block' : 'none';
              banner.textContent = 'Current document colour mode is RGB';
            }
          }
          data = data.colors;
        }
        if(!data || !Array.isArray(data)) return;
        if(countEl) countEl.textContent = 'Document colours: ' + data.length + (debug ? ' (items: ' + (debug.totalItems || 0) + ', scanned: ' + (debug.scanned || 0) + ', path: ' + (debug.pathItems || 0) + ', text: ' + (debug.textFrames || 0) + ', fallback: ' + (debug.fallbackUsed ? 'yes' : 'no') + ')' : '');
        if(showToastOnComplete) showToast('Found ' + data.length + ' document colours.', {type: 'success', title: 'Refresh colours'});

        function round1(v) {
          return Math.round(v * 10) / 10;
        }
        function clamp(v, min, max) {
          return Math.max(min, Math.min(max, v));
        }
        function hexToRgb(hex) {
          const s0 = String(hex || '').trim().replace('#', '');
          const s = s0.length === 3 ? s0.split('').map(ch => ch + ch).join('') : s0;
          if (s.length !== 6) return null;
          const r = parseInt(s.slice(0, 2), 16);
          const g = parseInt(s.slice(2, 4), 16);
          const b = parseInt(s.slice(4, 6), 16);
          if (![r, g, b].every(v => isFinite(v))) return null;
          return {r, g, b};
        }
        function rgbToCmykValues(r, g, b) {
          const rn = clamp(num(r), 0, 255) / 255;
          const gn = clamp(num(g), 0, 255) / 255;
          const bn = clamp(num(b), 0, 255) / 255;
          const k = 1 - Math.max(rn, gn, bn);
          if (k >= 0.9999) return {c: 0, m: 0, y: 0, k: 100};
          const d = 1 - k;
          return {
            c: ((1 - rn - k) / d) * 100,
            m: ((1 - gn - k) / d) * 100,
            y: ((1 - bn - k) / d) * 100,
            k: k * 100
          };
        }
        function getDisplayCmyk(entry) {
          const c = Number(entry.c);
          const m = Number(entry.m);
          const y = Number(entry.y);
          const k = Number(entry.k);
          if ([c, m, y, k].every(v => isFinite(v))) {
            return {c, m, y, k};
          }
          const rgb = hexToRgb(entry.hex);
          if (rgb) return rgbToCmykValues(rgb.r, rgb.g, rgb.b);
          return rgbToCmykValues(entry.r || 0, entry.g || 0, entry.b || 0);
        }
        function isNearRichBlack(cmyk, tolerance) {
          const tol = isFinite(tolerance) ? tolerance : 8;
          return Math.abs(cmyk.c - 60) <= tol &&
            Math.abs(cmyk.m - 60) <= tol &&
            Math.abs(cmyk.y - 60) <= tol &&
            Math.abs(cmyk.k - 100) <= 3;
        }
        function isBlackLike(cmyk, hex) {
          const byCmyk = cmyk.k >= 90 && (cmyk.c + cmyk.m + cmyk.y) <= 210;
          const rgb = hexToRgb(hex);
          const byRgb = !!(rgb && rgb.r <= 28 && rgb.g <= 28 && rgb.b <= 28);
          return byCmyk || byRgb;
        }

        data.forEach((entry) => {
          if(colourEditState.lastEdit && entry.type === colourEditState.lastEdit.type &&
            colourEditState.lastEdit.mode === colourEditState.mode) {
            const expectedHex = colourEditState.mode === 'RGB'
              ? rgbToHex(colourEditState.lastEdit.r, colourEditState.lastEdit.g, colourEditState.lastEdit.b)
              : cmykToHex(colourEditState.lastEdit.c, colourEditState.lastEdit.m, colourEditState.lastEdit.y, colourEditState.lastEdit.k);
            const matchesEdit = entry.key === colourEditState.lastEdit.toKey ||
              entry.key === colourEditState.lastEdit.fromKey ||
              (entry.hex && entry.hex.toLowerCase() === expectedHex.toLowerCase()) ||
              (entry.hex && colourEditState.lastEdit.fromHex &&
                entry.hex.toLowerCase() === colourEditState.lastEdit.fromHex.toLowerCase());
            if(matchesEdit) {
              if(colourEditState.mode === 'RGB') {
                entry.r = colourEditState.lastEdit.r;
                entry.g = colourEditState.lastEdit.g;
                entry.b = colourEditState.lastEdit.b;
                entry.key = colourEditState.lastEdit.toKey;
                entry.hex = expectedHex;
                entry.label = (entry.type || 'fill').toUpperCase() + '  R ' + entry.r + ' G ' + entry.g + ' B ' + entry.b;
              } else {
                entry.c = colourEditState.lastEdit.c;
                entry.m = colourEditState.lastEdit.m;
                entry.y = colourEditState.lastEdit.y;
                entry.k = colourEditState.lastEdit.k;
                entry.key = colourEditState.lastEdit.toKey;
                entry.hex = expectedHex;
                entry.label = (entry.type || 'fill').toUpperCase() + '  C ' + entry.c + ' M ' + entry.m + ' Y ' + entry.y + ' K ' + entry.k;
              }
              const valueLog = colourEditState.mode === 'RGB'
                ? [entry.r, entry.g, entry.b].join(',')
                : [entry.c, entry.m, entry.y, entry.k].join(',');
              log('Colours: matched lastEdit toKey=' + colourEditState.lastEdit.toKey + ' type=' + (entry.type || '') +
                ' values=' + valueLog + ' hexMatch=' +
                (entry.hex ? entry.hex.toLowerCase() === expectedHex.toLowerCase() : false) +
                ' fromKeyMatch=' + (entry.key === colourEditState.lastEdit.fromKey) +
                ' fromHexMatch=' + (entry.hex && colourEditState.lastEdit.fromHex ?
                  entry.hex.toLowerCase() === colourEditState.lastEdit.fromHex.toLowerCase() : false));
            }
          }
          if(colourEditState.mode === 'RGB') {
            log('Colours: row values key=' + (entry.key || '') + ' type=' + (entry.type || '') +
              ' rgb=' + [entry.r, entry.g, entry.b].join(',') + ' hex=' + (entry.hex || ''));
          } else {
            log('Colours: row values key=' + (entry.key || '') + ' type=' + (entry.type || '') +
              ' cmyk=' + [entry.c, entry.m, entry.y, entry.k].join(',') + ' hex=' + (entry.hex || ''));
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
          swatch.style.minWidth = '34px';
          swatch.style.minHeight = '34px';
          swatch.style.maxWidth = '34px';
          swatch.style.maxHeight = '34px';
          swatch.style.flex = '0 0 34px';
          swatch.style.border = '1px solid #444';
          swatch.style.borderRadius = '6px';
          swatch.style.background = entry.hex || '#000000';

          const cmyk = getDisplayCmyk(entry);
          const typeText = String(entry.type || 'fill').toUpperCase();
          const cmykLine = 'C ' + round1(cmyk.c) + '  M ' + round1(cmyk.m) + '  Y ' + round1(cmyk.y) + '  K ' + round1(cmyk.k);
          const showBlackHazard = isBlackLike(cmyk, entry.hex) && !isNearRichBlack(cmyk, 8);

          const label = document.createElement('div');
          label.style.display = 'flex';
          label.style.flexDirection = 'column';
          label.style.gap = '2px';
          label.style.minWidth = '180px';
          label.style.flex = '1 1 auto';
          const labelTop = document.createElement('div');
          labelTop.textContent = typeText;
          labelTop.style.fontSize = '12px';
          labelTop.style.fontWeight = '700';
          labelTop.style.color = '#d7d7d7';
          const labelBottom = document.createElement('div');
          labelBottom.textContent = cmykLine;
          labelBottom.style.fontSize = '11px';
          labelBottom.style.color = '#bcbcbc';
          label.appendChild(labelTop);
          label.appendChild(labelBottom);

          const inputs = [];
          if(colourEditState.mode === 'RGB') {
            const rInput = document.createElement('input');
            rInput.type = 'number'; rInput.step = '1'; rInput.min = '0'; rInput.max = '255';
            rInput.value = entry.r || 0;
            const gInput = document.createElement('input');
            gInput.type = 'number'; gInput.step = '1'; gInput.min = '0'; gInput.max = '255';
            gInput.value = entry.g || 0;
            const bInput = document.createElement('input');
            bInput.type = 'number'; bInput.step = '1'; bInput.min = '0'; bInput.max = '255';
            bInput.value = entry.b || 0;
            inputs.push(rInput, gInput, bInput);
          } else {
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
            inputs.push(cInput, mInput, yInput, kInput);
          }
          inputs.forEach(inp => {
            inp.style.width = '35px';
          });
          log('Colours: set inputs key=' + (entry.key || '') + ' type=' + (entry.type || '') +
            ' values=' + inputs.map(i => i.value).join(','));

          function applyValues() {
            if(colourEditState.mode === 'RGB') {
              const rRaw = parseFloat(inputs[0].value);
              const gRaw = parseFloat(inputs[1].value);
              const bRaw = parseFloat(inputs[2].value);
              const r = isFinite(rRaw) ? rRaw : num(entry.r);
              const g = isFinite(gRaw) ? gRaw : num(entry.g);
              const b = isFinite(bRaw) ? bRaw : num(entry.b);
              const toHex = rgbToHex(r, g, b);
              const fromKey = entry.key || rgbKey(entry.r, entry.g, entry.b);
              const fromHex = entry.hex || rgbToHex(entry.r, entry.g, entry.b);
              const payload = JSON.stringify({
                fromKey: fromKey,
                fromType: entry.type || '',
                fromMode: 'RGB',
                fromHex: fromHex,
                toHex: toHex
              }).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
              log('Colours: replace fromKey=' + (fromKey || '') + ' type=' + (entry.type || '') +
                ' to=' + [r, g, b].join(',') + ' (inputs=' +
                [inputs[0].value, inputs[1].value, inputs[2].value].join(',') + ')');
              callJSX('signarama_helper_replaceColor("' + payload + '")', (result) => {
                const match = (result || '').match(/Updated\s+(\d+)/i);
                const updated = match ? parseInt(match[1], 10) : 0;
                log('Colours: replace result=' + (result || '') + ' updated=' + updated);
                swatch.style.background = toHex;
                colourEditState.lastEdit = {
                  mode: 'RGB',
                  type: entry.type || '',
                  fromKey: fromKey || '',
                  fromHex: fromHex || '',
                  toKey: rgbKey(r, g, b),
                  r, g, b
                };
                log('Colours: refresh after replace toKey=' + colourEditState.lastEdit.toKey);
                refreshColours();
              });
              return;
            }
            const cRaw = parseFloat(inputs[0].value);
            const mRaw = parseFloat(inputs[1].value);
            const yRaw = parseFloat(inputs[2].value);
            const kRaw = parseFloat(inputs[3].value);
            const c = isFinite(cRaw) ? cRaw : num(entry.c);
            const m = isFinite(mRaw) ? mRaw : num(entry.m);
            const y = isFinite(yRaw) ? yRaw : num(entry.y);
            const k = isFinite(kRaw) ? kRaw : num(entry.k);
            const payload = JSON.stringify({
              fromKey: entry.key || '',
              fromType: entry.type || '',
              fromMode: 'CMYK',
              toCmyk: {c, m, y, k}
            }).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
            log('Colours: replace fromKey=' + (entry.key || '') + ' type=' + (entry.type || '') +
              ' to=' + [c, m, y, k].join(',') + ' (inputs=' +
              [inputs[0].value, inputs[1].value, inputs[2].value, inputs[3].value].join(',') + ')');
            callJSX('signarama_helper_replaceColor("' + payload + '")', (result) => {
              const match = (result || '').match(/Updated\s+(\d+)/i);
              const updated = match ? parseInt(match[1], 10) : 0;
              log('Colours: replace result=' + (result || '') + ' updated=' + updated);
              if(!updated) {
                refreshColours();
                return;
              }
              swatch.style.background = cmykToHex(c, m, y, k);
              colourEditState.lastEdit = {
                mode: 'CMYK',
                type: entry.type || '',
                fromKey: entry.key || '',
                fromHex: entry.hex || '',
                toKey: cmykKey(c, m, y, k),
                c, m, y, k
              };
              log('Colours: refresh after replace toKey=' + colourEditState.lastEdit.toKey);
              refreshColours();
            });
          }

          inputs.forEach((inp, idx) => {
            inp.addEventListener('change', applyValues);
            inp.addEventListener('focus', () => {
              try {inp.select();} catch(_eSel) { }
            });
            inp.addEventListener('keydown', (e) => {
              if(e.key === 'Tab') {
                e.preventDefault();
                const next = inputs[(idx + 1) % inputs.length];
                if(next) {next.focus(); try {next.select();} catch(_eSel2) { } }
              }
            });
          });

          // Focus restore metadata
          inputs.forEach((inp, idx) => {
            inp.dataset.focusKey = entry.key || '';
            inp.dataset.focusType = entry.type || '';
            inp.dataset.focusHex = entry.hex || '';
            inp.dataset.focusIndex = String(idx);
          });

          row.appendChild(swatch);
          row.appendChild(label);
          inputs.forEach(inp => row.appendChild(inp));
          if (showBlackHazard) {
            const hazard = document.createElement('span');
            hazard.textContent = '\u26A0';
            hazard.title = 'Black colour is not close to target rich black (C60 M60 Y60 K100).';
            hazard.style.color = '#ffba00';
            hazard.style.fontWeight = '800';
            hazard.style.fontSize = '16px';
            hazard.style.marginLeft = '6px';
            row.appendChild(hazard);
          }
          list.appendChild(row);
        });

        if(debug) {
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

        if(focusedMeta) {
          let nextFocus = null;
          let focusKey = focusedMeta.focusKey;
          if(
            colourEditState.lastEdit &&
            focusKey &&
            focusedMeta.focusType === colourEditState.lastEdit.type &&
            focusKey === colourEditState.lastEdit.fromKey &&
            colourEditState.lastEdit.mode === colourEditState.mode
          ) {
            focusKey = colourEditState.lastEdit.toKey;
          }
          if(focusedMeta.focusKey) {
            const selectorKey = '[data-focus-key=\"' + focusKey + '\"][data-focus-index=\"' + focusedMeta.index + '\"]';
            nextFocus = list.querySelector(selectorKey);
          }
          if(!nextFocus && focusedMeta.focusHex) {
            const selectorHex = '[data-focus-hex=\"' + focusedMeta.focusHex + '\"][data-focus-index=\"' + focusedMeta.index + '\"]';
            nextFocus = list.querySelector(selectorHex);
          }
          if(nextFocus) {
            nextFocus.focus();
            try {nextFocus.select();} catch(_eSel3) { }
          }
        }
      });
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

  function rgbToHex(r, g, b) {
    const clamp = v => Math.max(0, Math.min(255, Math.round(v)));
    const h = v => clamp(v).toString(16).padStart(2, '0');
    return '#' + h(r) + h(g) + h(b);
  }

  function wireDimensions() {
    const dimClear = $('btnDimClear');
    if(dimClear) dimClear.onclick = () => runButtonJsxOperation('atlas_dimensions_clear()', {logFn: logDim, toastTitle: 'Clear dimensions'});
    const dimSelectionHint = $('dimSelectionHint');
    const textColorInput = $('textColor');
    const lineColorInput = $('lineColor');

    function refreshDimensionSelectionHint() {
      if(!dimSelectionHint) return;
      const dimPanel = $('tab-dimensions');
      if(!dimPanel || dimPanel.classList.contains('hidden')) return;
      callJSX('atlas_dimensions_hasSelection()', function(res) {
        const hasSelection = String(res || '') === '1';
        dimSelectionHint.classList.toggle('hidden', hasSelection);
      });
    }

    function runSides(sides) {
      const payload = buildDimensionPayload();
      payload.sides = sides;
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      runButtonJsxOperation('atlas_dimensions_runMulti("' + json + '")', {
        logFn: logDim,
        toastTitle: 'Add dimensions',
        onResult: function() {
          refreshDimensionSelectionHint();
        }
      });
    }

    function runLineMeasure() {
      const payload = buildDimensionPayload();
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      runButtonJsxOperation('atlas_dimensions_runLine("' + json + '")', {
        logFn: logDim,
        toastTitle: 'Measure line/path',
        onResult: function() {
          refreshDimensionSelectionHint();
        }
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
      if(el) el.onclick = map[id];
    });

    $all('.theme-tile[data-text-color][data-line-color]').forEach(tile => {
      tile.addEventListener('click', () => {
        const textHex = tile.getAttribute('data-text-color') || '';
        const lineHex = tile.getAttribute('data-line-color') || '';
        if(textColorInput && textHex) textColorInput.value = textHex;
        if(lineColorInput && lineHex) lineColorInput.value = lineHex;
        try {if(textColorInput) textColorInput.dispatchEvent(new Event('change', {bubbles: true}));} catch(_eTC) { }
        try {if(lineColorInput) lineColorInput.dispatchEvent(new Event('change', {bubbles: true}));} catch(_eLC) { }
        if(schedulePanelSettingsSave) schedulePanelSettingsSave();
      });
    });

    refreshDimensionSelectionHint();
    setInterval(refreshDimensionSelectionHint, 1500);
  }

  function loadJSX(done) {
    try {
      var extDir = cs.getSystemPath(SystemPath.EXTENSION).replace(/\\/g, '/');
      var cmd = '$.evalFile("' + extDir + '/jsx/hostscript.jsx")';
      cs.evalScript(cmd, function() {
        log('JSX loaded.');
        if(done) done();
      });
    } catch(e) {
      log('Failed to load JSX: ' + (e && e.message ? e.message : e));
    }
  }

  function callJSX(fnCall, cb) {
    var wrapped = '(function(){try{return ' + fnCall + ' }catch(e){return "Error: " + e}})()';
    cs.evalScript(wrapped, function(res) {if(cb) cb(res);});
  }
})();

