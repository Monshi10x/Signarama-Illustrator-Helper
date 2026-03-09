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
      '#srhToastRoot{position:fixed;top:14px;right:14px;left:auto;z-index:2147483000;display:flex;flex-direction:column;align-items:flex-end;gap:10px;pointer-events:none;}',
      '.srh-toast{min-width:250px;max-width:420px;min-height:68px;background:#141414;color:#f2f2f2;border:1px solid #3a3a3a;border-radius:12px;box-shadow:inset 0px 1px 2px rgba(255,255,255,0.12),5px 8px 10px rgba(0,0,0,0.45),0 0px 1px rgba(0,0,0,0.6);overflow:hidden;pointer-events:auto;opacity:0;transform:translateY(-6px);transition:opacity .14s ease,transform .14s ease;}',
      '.srh-toast.in{opacity:1;transform:translateY(0);}',
      '.srh-toast-head{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:10px 12px 6px 12px;font-size:12px;font-weight:800;line-height:1.25;}',
      '.srh-toast-headMain{display:flex;align-items:center;gap:8px;min-width:0;}',
      '.srh-toast-title{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}',
      '.srh-toast-spinner{width:13px;height:13px;border:2px solid rgba(255,255,255,0.28);border-top-color:#fff;border-radius:50%;animation:srhToastSpin .7s linear infinite;flex:0 0 auto;}',
      '.srh-toast-close{appearance:none;border:1px solid #4b4b4b;background:#1f1f1f;color:#d9d9d9;border-radius:8px;width:20px;height:20px;line-height:16px;cursor:pointer;padding:0;}',
      '.srh-toast-close:hover{background:#2a2a2a;}',
      '.srh-toast-body{padding:0 12px 10px 12px;min-height:18px;font-size:12px;line-height:1.4;word-break:break-word;}',
      '.srh-toast-progressWrap{height:4px;background:#222;}',
      '.srh-toast-progress{height:100%;width:100%;transform-origin:left center;background:#5a5a5a;}',
      '@keyframes srhToastSpin{to{transform:rotate(360deg);}}',
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
    const isPersistent = !!opts.persistent;
    const duration = typeof opts.duration === 'number' ? Math.max(400, opts.duration) : 5000;
    const title = opts.title || (type === 'error' ? 'Operation failed' : 'Operation finished');
    const root = $('srhToastRoot');
    if(!root) return null;

    const toast = document.createElement('div');
    toast.className = 'srh-toast ' + type;
    const head = document.createElement('div');
    head.className = 'srh-toast-head';
    const headMain = document.createElement('div');
    headMain.className = 'srh-toast-headMain';
    const spinner = document.createElement('span');
    spinner.className = 'srh-toast-spinner';
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

    if(opts.spinner) headMain.appendChild(spinner);
    headMain.appendChild(titleEl);
    head.appendChild(headMain);
    head.appendChild(close);
    pWrap.appendChild(pBar);
    toast.appendChild(head);
    toast.appendChild(body);
    if(!isPersistent) toast.appendChild(pWrap);
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
      if(isPersistent) return;
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
      if(!isPersistent) raf = requestAnimationFrame(tick);
    });
    return {close: finish};
  }
  function notifyOperationResult(res, options) {
    const opts = options || {};
    if(opts.showToast === false) return;
    const text = String(res || '').trim();
    const isError = /^error:/i.test(text);
    const isWarn = !isError && (/^no\b/i.test(text) || /\bno\s+(selection|document|items?|artboards?)\b/i.test(text));
    const type = isError ? 'error' : (isWarn ? 'warn' : 'success');
    const title = opts.toastTitle || (isError ? 'Operation failed' : 'Operation finished');
    const message = (!isError && !isWarn && opts.toastMessage) ? opts.toastMessage : (text || opts.emptyMessage || title);
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
  let corebridgePageNumberWatchTimer = null;
  let corebridgePageNumberWatcherErrorLogged = false;
  let corebridgeFlashTickPollTimer = null;
  let corebridgeFlashTickPollCount = 0;
  let corebridgeFlashLastTickCount = -1;
  let coloursPendingApplyFns = [];
  let coloursHasPendingChanges = false;
  let isLargeArtboard = false;
  let refreshLightboxArtboardScaleNotice = null;
  let activateTabFn = null;
  let persistSettingsTimer = null;
  let isApplyingPanelSettings = false;
  let schedulePanelSettingsSave = null;
  let scheduleDimensionSelectionHintRefresh = null;

  function jsxEscapeDoubleQuoted(text) {
    return String(text)
      .replace(/\\/g, '\\\\')
      .replace(/"/g, '\\"')
      .replace(/\r/g, '\\r')
      .replace(/\n/g, '\\n');
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
    callJSX('((typeof signarama_helper_panelSettingsSave === "function") ? signarama_helper_panelSettingsSave : ((typeof $ !== "undefined" && $.global && typeof $.global.signarama_helper_panelSettingsSave === "function") ? $.global.signarama_helper_panelSettingsSave : function(){return "Error: settings save function not loaded.";}))("' + safe + '")', function(res) {
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
    callJSX('((typeof signarama_helper_panelSettingsLoad === "function") ? signarama_helper_panelSettingsLoad : ((typeof $ !== "undefined" && $.global && typeof $.global.signarama_helper_panelSettingsLoad === "function") ? $.global.signarama_helper_panelSettingsLoad : function(){return "NO_SETTINGS";}))()', function(res) {
      const txt = String(res || '');
      if(txt && txt !== 'NO_SETTINGS') {
        try {
          const parsed = JSON.parse(txt);
          applyPanelSettings(parsed);
          log('Panel settings loaded.');
        } catch(e) {
          try {
            const migrated = Function('return ' + txt)();
            if(migrated && typeof migrated === 'object') {
              applyPanelSettings(migrated);
              persistPanelSettings();
              log('Panel settings loaded (migrated legacy format).');
            } else {
              log('Panel settings parse failed: ' + (e && e.message ? e.message : e));
            }
          } catch(_eLegacy) {
            log('Panel settings parse failed: ' + (e && e.message ? e.message : e));
          }
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
      areaApproximationStep: num(($('areaApproximationStep') && $('areaApproximationStep').value) || 10),
      includeArrowhead: !!($('includeArrowhead') && $('includeArrowhead').checked),
      showAreaApproximation: !!($('showAreaApproximation') && $('showAreaApproximation').checked),
      textColor: ($('textColor') && $('textColor').value) || '#000000',
      lineColor: ($('lineColor') && $('lineColor').value) || '#000000',
      scaleAppearance: num(($('scaleAppearance') && $('scaleAppearance').value) || 100) / 100
    };
  }

  if(typeof CSInterface === 'undefined') {
    alert('CSInterface.js NOT loaded. Fix script paths.');
    return;
  }

  const cs = new CSInterface();
  const srhGlobalState = (typeof window !== 'undefined')
    ? (window.srhGlobalState = window.srhGlobalState || {})
    : {};
  if(!srhGlobalState.corebridgePartSearchEntriesByName) srhGlobalState.corebridgePartSearchEntriesByName = {};
  if(!srhGlobalState.corebridgePartSearchEntriesRaw) srhGlobalState.corebridgePartSearchEntriesRaw = [];

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
      cs.evalScript('((typeof signarama_helper_fitArtboardToArtwork==="function" || (typeof $!=="undefined" && $.global && typeof $.global.signarama_helper_fitArtboardToArtwork==="function")) ? "function" : "undefined")', function(type) {
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
      if(tabId === 'tab-dimensions' && typeof scheduleDimensionSelectionHintRefresh === 'function') {
        scheduleDimensionSelectionHintRefresh();
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
    const fitArtboardMarginMm = $('fitArtboardMarginMm');
    if(fitArtboardMarginMm) {
      const fitMarginLabel = fitArtboardMarginMm.closest('label');
      if(fitMarginLabel) {
        fitMarginLabel.addEventListener('click', (e) => {
          e.stopPropagation();
        });
      }
      fitArtboardMarginMm.addEventListener('click', (e) => {
        e.stopPropagation();
      });
      fitArtboardMarginMm.addEventListener('change', (e) => {
        e.stopPropagation();
      });
    }
    if(fit) fit.onclick = () => {
      const marginMm = num((fitArtboardMarginMm && fitArtboardMarginMm.value) ? fitArtboardMarginMm.value : 0);
      runButtonJsxOperation('signarama_helper_fitArtboardToArtwork(' + marginMm + ')', {logFn: log, toastTitle: 'Fit artboard'});
    };

    const ab = $('btnArtboardPerItem');
    if(ab) ab.onclick = () => runButtonJsxOperation('signarama_helper_createArtboardsFromSelection()', {logFn: log, toastTitle: 'Artboard per selection'});

    const a4 = $('btnCopyOutlineScaleA4');
    const a4Rasterize = $('a4Rasterize');
    const a4RasterizeQuality = $('a4RasterizeQuality');
    if(a4Rasterize) {
      const rasterizeLabel = a4Rasterize.closest('label');
      if(rasterizeLabel) {
        rasterizeLabel.addEventListener('click', (e) => {
          e.stopPropagation();
        });
      }
      a4Rasterize.addEventListener('click', (e) => {
        e.stopPropagation();
      });
      a4Rasterize.addEventListener('change', (e) => {
        e.stopPropagation();
      });
    }
    if(a4RasterizeQuality) {
      a4RasterizeQuality.addEventListener('click', (e) => {
        e.stopPropagation();
      });
      a4RasterizeQuality.addEventListener('change', (e) => {
        e.stopPropagation();
      });
    }
    if(a4) a4.onclick = () => {
      const payload = {
        rasterize: !!(a4Rasterize && a4Rasterize.checked),
        rasterizeQuality: (a4RasterizeQuality && a4RasterizeQuality.value) ? a4RasterizeQuality.value : 'high'
      };
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      runButtonJsxOperation('signarama_helper_duplicateOutlineScaleA4("' + json + '")', {logFn: log, toastTitle: 'Scale artwork for proof'});
    };

    const preset = $('bleedPreset');
    if(preset) {
      preset.onchange = () => {
        const v = preset.value;
        if(v === 'acm') {
          $('bleedTop').value = 20; $('bleedRight').value = 20; $('bleedBottom').value = 20; $('bleedLeft').value = 20;
        } else if(v === 'windowGraphics') {
          $('bleedTop').value = 0; $('bleedRight').value = 20; $('bleedBottom').value = 20; $('bleedLeft').value = 0;
        } else if(v === 'wallGraphics') {
          $('bleedTop').value = 20; $('bleedRight').value = 50; $('bleedBottom').value = 50; $('bleedLeft').value = 0;
        } else if(v === 'flyers') {
          $('bleedTop').value = 2; $('bleedRight').value = 2; $('bleedBottom').value = 2; $('bleedLeft').value = 2;
        } else if(v === 'aframes') {
          $('bleedTop').value = 20; $('bleedRight').value = 20; $('bleedBottom').value = 20; $('bleedLeft').value = 20;
        } else if(v === 'pullupBanners') {
          $('bleedTop').value = 0; $('bleedRight').value = 0; $('bleedBottom').value = 100; $('bleedLeft').value = 0;
        } else if(v === 'none') {
          $('bleedTop').value = 0; $('bleedRight').value = 0; $('bleedBottom').value = 0; $('bleedLeft').value = 0;
        }
      };
      preset.onchange();
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
        runButtonJsxOperation('signarama_helper_applyBleed(' + [t, l, b, r, excludeClipped, keepOriginal, expandArtboards].join(',') + ')', {logFn: log, toastTitle: 'Apply bleed', toastMessage: 'Finished apply bleed.'});
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

    function ensureCorebridgePageNumberWatcher() {
      if(corebridgePageNumberWatchTimer) return;
      corebridgePageNumberWatchTimer = setInterval(() => {
        callJSX('((typeof signarama_helper_corebridge_updatePageNumbers === "function") ? signarama_helper_corebridge_updatePageNumbers : ((typeof $ !== "undefined" && $.global && typeof $.global.signarama_helper_corebridge_updatePageNumbers === "function") ? $.global.signarama_helper_corebridge_updatePageNumbers : function(){return "NO_FN";}))()', (res) => {
          const txt = String(res || '');
          if(/^Error:/i.test(txt) && !corebridgePageNumberWatcherErrorLogged) {
            corebridgePageNumberWatcherErrorLogged = true;
            log('Page number watcher error: ' + txt);
          }
          if(!/^Error:/i.test(txt)) corebridgePageNumberWatcherErrorLogged = false;
        });
      }, 1800);
    }

    ensureCorebridgePageNumberWatcher();

    const corebridgePullData = $('btnCorebridgePullData');
    const corebridgeOpenProof = $('btnCorebridgeOpenProof');
    const corebridgeCreateProofFromData = $('btnCorebridgeCreateProofFromData');
    const corebridgeCreateProofForSelected = $('btnCorebridgeCreateProofForSelected');
    const corebridgeDevMode = $('corebridgeDevMode');
    const corebridgePullControlsWrap = $('corebridgePullControlsWrap');
    const corebridgeDevDumpWrap = $('corebridgeDevDumpWrap');
    const corebridgeProofMappingsWrap = $('corebridgeProofMappingsWrap');
    const corebridgeFlashFieldsWrap = $('corebridgeFlashFieldsWrap');
    const corebridgeDerivedMappingsWrap = $('corebridgeDerivedMappingsWrap');
    const corebridgeJobNumber = $('corebridgeJobNumber');
    const corebridgeItemNumber = $('corebridgeItemNumber');
    const corebridgeProofPath = $('corebridgeProofPath');
    const corebridgeProofMappings = $('corebridgeProofMappings');
    const corebridgeFlashFields = $('corebridgeFlashFields');
    const corebridgeDerivedMappingsPreview = document.querySelector('textarea[data-corebridge-derived-mappings]');
    const corebridgeDumpHost = $('corebridgeDataDumpHost');
    const corebridgeFetchStatus = $('corebridgeFetchStatus');
    function stopCorebridgeFlashTickPolling(reason) {
      if(corebridgeFlashTickPollTimer) {
        clearInterval(corebridgeFlashTickPollTimer);
        corebridgeFlashTickPollTimer = null;
      }
      if(reason) log('Corebridge flash poll stop: ' + reason);
      corebridgeFlashLastTickCount = -1;
    }
    function startCorebridgeFlashTickPolling() {
      stopCorebridgeFlashTickPolling('restart');
      corebridgeFlashTickPollCount = 0;
      log('Corebridge flash poll start (300ms).');
      corebridgeFlashTickPollTimer = setInterval(() => {
        corebridgeFlashTickPollCount++;
        callJSX('((typeof signarama_helper_corebridge_flashTickTask === "function") ? signarama_helper_corebridge_flashTickTask : ((typeof $ !== "undefined" && $.global && typeof $.global.signarama_helper_corebridge_flashTickTask === "function") ? $.global.signarama_helper_corebridge_flashTickTask : function(){return "ERROR|flashTickTask missing";}))()', (tickRes) => {
          const tickTxt = String(tickRes || '').trim();
          if(/^ERROR\|/i.test(tickTxt)) {
            log('Corebridge flash tick error: ' + tickTxt);
            stopCorebridgeFlashTickPolling('tick error');
            return;
          }
          if(corebridgeFlashTickPollCount <= 10 || corebridgeFlashTickPollCount % 10 === 0 || /^INACTIVE\|/i.test(tickTxt)) {
            log('Corebridge flash poll #' + corebridgeFlashTickPollCount + ': ' + tickTxt);
          }
          if(/^INACTIVE\|/i.test(tickTxt)) {
            stopCorebridgeFlashTickPolling('inactive');
            return;
          }
          const parts = tickTxt.split('|');
          if(parts.length >= 4 && /^ACTIVE$/i.test(parts[0])) {
            const tickCount = parseInt(parts[2], 10) || 0;
            if(corebridgeFlashLastTickCount >= 0 && tickCount <= corebridgeFlashLastTickCount) {
              log('Corebridge flash poll warning: non-incrementing tick count (' + tickCount + ').');
            }
            corebridgeFlashLastTickCount = tickCount;
          }
        });
      }, 300);
    }
    function invalidateCorebridgeFetchCache() {
      corebridgeHasFetchedData = false;
      corebridgeLastFilteredData = [];
      corebridgeLastSecondaryFetchResults = null;
    }
    if(corebridgeJobNumber) {
      corebridgeJobNumber.addEventListener('input', invalidateCorebridgeFetchCache);
      corebridgeJobNumber.addEventListener('change', invalidateCorebridgeFetchCache);
    }
    if(corebridgeItemNumber) {
      corebridgeItemNumber.addEventListener('input', invalidateCorebridgeFetchCache);
      corebridgeItemNumber.addEventListener('change', invalidateCorebridgeFetchCache);
    }
    const corebridgeFetchTimeoutMs = 20000;
    const corebridgeProxyBaseUrl = 'http://localhost:8080';//'https://signschedulerapp.ts.r.appspot.com';
    const corebridgePrimaryDataUrl = corebridgeProxyBaseUrl + '/CB_DesignBoard_Data';
    const corebridgePartSearchEntriesUrl = corebridgeProxyBaseUrl + '/CB_OrderEntryProducts_PartSearchEntries';
    async function preloadCorebridgePartSearchEntries() {
      function tryParseJsonLoose(value) {
        if(value == null) return null;
        if(typeof value === 'object') return value;
        const txt = String(value).trim();
        if(!txt) return null;
        try { return JSON.parse(txt); } catch(_ePsLoose) { return null; }
      }
      function extractPartRows(payload) {
        const queue = [payload];
        const seen = [];
        while(queue.length) {
          const cur = queue.shift();
          if(!cur) continue;
          if(typeof cur === 'string') {
            const parsedCur = tryParseJsonLoose(cur);
            if(parsedCur && seen.indexOf(parsedCur) < 0) queue.push(parsedCur);
            continue;
          }
          if(Array.isArray(cur)) {
            if(cur.length && typeof cur[0] === 'object') {
              const hasPartShape = cur.some((row) => row && (row.PartName != null || row.Name != null || row.name != null || row.DisplayName != null || row.displayName != null));
              if(hasPartShape) return cur;
            }
            for(let i = 0; i < cur.length; i++) queue.push(cur[i]);
            continue;
          }
          if(typeof cur === 'object') {
            if(seen.indexOf(cur) >= 0) continue;
            seen.push(cur);
            const keys = Object.keys(cur);
            for(let k = 0; k < keys.length; k++) queue.push(cur[keys[k]]);
          }
        }
        return [];
      }
      try {
        const res = await fetch(corebridgePartSearchEntriesUrl + '?_ts=' + Date.now(), {
          method: 'GET',
          cache: 'no-store',
          headers: {pragma: 'no-cache', 'cache-control': 'no-cache'}
        });
        const text = await res.text();
        if(!res.ok) throw new Error('HTTP ' + res.status + ' ' + res.statusText);
        const parsed = tryParseJsonLoose(text);
        const rows = extractPartRows(parsed);
        const byName = {};
        rows.forEach((row) => {
          const rawName = String((row && (row.PartName || row.Name || row.name || row.DisplayName || row.displayName)) || '').trim();
          const thicknessRaw = (row && (row.Thickness != null ? row.Thickness : (row.thickness != null ? row.thickness : '')));
          const thickness = String(thicknessRaw == null ? '' : thicknessRaw).trim();
          if(!rawName) return;
          const norm = rawName.toLowerCase().replace(/\s+/g, ' ').trim();
          if(!byName[norm]) byName[norm] = {name: rawName, thickness: thickness};
        });
        srhGlobalState.corebridgePartSearchEntriesByName = byName;
        srhGlobalState.corebridgePartSearchEntriesRaw = rows;
        if(typeof window !== 'undefined') {
          window.corebridgePartSearchEntriesByName = byName;
          window.corebridgePartSearchEntriesRaw = rows;
        }
        const sample = rows.slice(0, 5).map((row) => ({
          name: String((row && (row.PartName || row.Name || row.name || row.DisplayName || row.displayName)) || ''),
          thickness: String((row && (row.Thickness != null ? row.Thickness : (row.thickness != null ? row.thickness : ''))) || '')
        }));
        log('Corebridge part search preload: ' + rows.length + ' rows (' + Object.keys(byName).length + ' indexed).');
        log('Corebridge part search payload sample: ' + JSON.stringify(sample));
      } catch(err) {
        log('Corebridge part search preload failed: ' + (err && err.message ? err.message : err));
      }
    }
    preloadCorebridgePartSearchEntries();
    let corebridgeLastFilteredData = [];
    let corebridgeLastSecondaryFetchResults = null;
    let corebridgeHasFetchedData = false;
    let corebridgeLastFetchCriteria = {jobNumber: "", itemNumber: ""};
    let corebridgeFetchPromise = null;
    const corebridgeDerivedMappingsPreviewText = [
      'Derived.installAddress -> Address Text',
      'Derived.todayDate -> Date Text',
      'Derived.productQty -> Quantity',
      'Derived.lineItemDescription -> Description',
      'Derived.mediaText -> Media Text',
      'Derived.laminateText -> Laminate Text',
      'Derived.substrateText -> Substrate Text',
      'Derived.partsNumbered -> Parts',
      'Derived.notesAll -> Notes',
      'DerivedAssets.addressQrSvg -> Address QR'
    ].join('\n');
    function setCorebridgeDerivedMappingsPreview() {
      if(!corebridgeDerivedMappingsPreview) return;
      corebridgeDerivedMappingsPreview.value = corebridgeDerivedMappingsPreviewText;
    }
    function normalizeCorebridgeProofMappingsField() {
      if(!corebridgeProofMappings) return;
      const raw = String(corebridgeProofMappings.value == null ? '' : corebridgeProofMappings.value);
      if(!raw) return;
      let next = raw;
      next = next.replace(/^\s*Derived\.installAddress\s*->\s*Address\s*$/gm, 'Derived.installAddress -> Address Text');
      next = next.replace(/^\s*Derived\.addressQrUrl\s*->\s*Address QR Code\s*$/gm, 'DerivedAssets.addressQrSvg -> Address QR');
      next = next.replace(/^\s*DerivedAssets\.addressQrSvg\s*->\s*Addres QR\s*$/gm, 'DerivedAssets.addressQrSvg -> Address QR');
      if(next !== raw) {
        corebridgeProofMappings.value = next;
        if(typeof schedulePanelSettingsSave === 'function') schedulePanelSettingsSave();
      }
    }
    function refreshCorebridgeDevModeUi() {
      const show = !!(corebridgeDevMode && corebridgeDevMode.checked);
      if(corebridgePullControlsWrap) corebridgePullControlsWrap.classList.toggle('hidden', !show);
      if(corebridgeDevDumpWrap) corebridgeDevDumpWrap.classList.toggle('hidden', !show);
      if(corebridgeProofMappingsWrap) corebridgeProofMappingsWrap.classList.toggle('hidden', !show);
      if(corebridgeFlashFieldsWrap) corebridgeFlashFieldsWrap.classList.toggle('hidden', !show);
      if(corebridgeDerivedMappingsWrap) corebridgeDerivedMappingsWrap.classList.toggle('hidden', !show);
    }
    setCorebridgeDerivedMappingsPreview();
    normalizeCorebridgeProofMappingsField();
    refreshCorebridgeDevModeUi();
    if(corebridgeDevMode) {
      corebridgeDevMode.addEventListener('change', refreshCorebridgeDevModeUi);
      corebridgeDevMode.addEventListener('input', refreshCorebridgeDevModeUi);
    }
    function normalizeCorebridgeInvoiceNumber(value) {
      return String(value == null ? '' : value).replace(/\D/g, '');
    }
    function getCorebridgeCriteriaFromFields() {
      const jobNumberRaw = (corebridgeJobNumber && corebridgeJobNumber.value ? corebridgeJobNumber.value : '').trim();
      const jobNumber = normalizeCorebridgeInvoiceNumber(jobNumberRaw);
      const itemNumber = (corebridgeItemNumber && corebridgeItemNumber.value ? corebridgeItemNumber.value : '').trim();
      return {jobNumber: jobNumber, itemNumber: itemNumber};
    }
    function corebridgeCriteriaChanged(criteria) {
      const next = criteria || getCorebridgeCriteriaFromFields();
      return next.jobNumber !== corebridgeLastFetchCriteria.jobNumber || next.itemNumber !== corebridgeLastFetchCriteria.itemNumber;
    }
    function buildCorebridgeSecondaryFetchPlan(filteredData) {
      const rows = Array.isArray(filteredData) ? filteredData : [];
      const quoteLevelCalls = [];
      const productNotesCalls = [];
      const quoteLevelSeen = {};
      const productNotesSeen = {};
      const warnings = [];

      rows.forEach((row, index) => {
        const rowIndex = index + 1;
        const orderId = row && row.OrderId != null ? String(row.OrderId).trim() : '';
        const accountId = row && row.AccountId != null ? String(row.AccountId).trim() : '';
        const accountName = row && row.CompanyName != null ? String(row.CompanyName).trim() : '';
        if(orderId && accountId && accountName) {
          const orderKey = [orderId, accountId, accountName].join('|');
          if(!quoteLevelSeen[orderKey]) {
            quoteLevelSeen[orderKey] = true;
            quoteLevelCalls.push({
              cbOrderId: orderId,
              cbAccountId: accountId,
              cbAccountName: accountName
            });
          }
        } else {
          warnings.push('Row ' + rowIndex + ': missing OrderId/AccountId/CompanyName for getOrderData_QuoteLevel.');
        }

        const productIdRaw = row && row.Id != null ? row.Id : (row && row.OrderProductId != null ? row.OrderProductId : (row && row.OrderProductID != null ? row.OrderProductID : ''));
        const orderProductId = String(productIdRaw == null ? '' : productIdRaw).trim();
        if(orderProductId) {
          if(!productNotesSeen[orderProductId]) {
            productNotesSeen[orderProductId] = true;
            productNotesCalls.push({orderProductId: orderProductId});
          }
        } else {
          warnings.push('Row ' + rowIndex + ': missing Id/OrderProductId for getProductNotesAll.');
        }
      });

      return {
        quoteLevelCalls: quoteLevelCalls,
        productNotesCalls: productNotesCalls,
        warnings: warnings
      };
    }
    function buildCorebridgeSecondaryFetchLog(plan) {
      const fetchPlan = plan || {quoteLevelCalls: [], productNotesCalls: [], warnings: []};
      const quoteLevelCalls = Array.isArray(fetchPlan.quoteLevelCalls) ? fetchPlan.quoteLevelCalls : [];
      const productNotesCalls = Array.isArray(fetchPlan.productNotesCalls) ? fetchPlan.productNotesCalls : [];
      const warnings = Array.isArray(fetchPlan.warnings) ? fetchPlan.warnings : [];
      const lines = [];

      lines.push('--- Secondary Fetches Required ---');
      lines.push('Source mapping: OrderId -> CB_OrderID, AccountId -> CB_AccountID, CompanyName -> CB_AccountName');
      lines.push('');
      lines.push('getOrderData_QuoteLevel(CB_OrderID, CB_AccountID, CB_AccountName):');
      if(quoteLevelCalls.length) {
        quoteLevelCalls.forEach((call, idx) => {
          const accountNameSafe = String(call.cbAccountName || '').replace(/"/g, '\\"');
          lines.push(
            (idx + 1) + '. getOrderData_QuoteLevel(' +
            call.cbOrderId + ', ' +
            call.cbAccountId + ', "' +
            accountNameSafe +
            '")'
          );
        });
      } else {
        lines.push('None (missing required fields).');
      }

      lines.push('');
      lines.push('getProductNotesAll(orderProductId):');
      if(productNotesCalls.length) {
        productNotesCalls.forEach((call, idx) => {
          lines.push((idx + 1) + '. getProductNotesAll(' + call.orderProductId + ')');
        });
      } else {
        lines.push('None (missing order product ID).');
      }

      if(warnings.length) {
        lines.push('');
        lines.push('Warnings:');
        warnings.forEach((msg) => {lines.push('- ' + msg);});
      }

      return lines.join('\n');
    }
    function setCorebridgeFetchLoading(isLoading) {
      if(corebridgeFetchStatus) corebridgeFetchStatus.classList.toggle('active', !!isLoading);
      if(corebridgePullData) corebridgePullData.disabled = !!isLoading;
      if(corebridgeCreateProofFromData) corebridgeCreateProofFromData.disabled = !!isLoading;
      if(corebridgeCreateProofForSelected) corebridgeCreateProofForSelected.disabled = !!isLoading;
    }
    function renderCorebridgeDataDump(text) {
      if(!corebridgeDumpHost) return;
      corebridgeDumpHost.innerHTML = '';
      const dumpEl = document.createElement('textarea');
      dumpEl.id = 'corebridgeDataDump';
      dumpEl.className = 'log';
      dumpEl.style.height = '220px';
      dumpEl.style.width = '100%';
      dumpEl.style.boxSizing = 'border-box';
      dumpEl.readOnly = true;
      dumpEl.value = text;
      corebridgeDumpHost.appendChild(dumpEl);
    }
    function appendCorebridgeDataDump(text) {
      if(!corebridgeDumpHost) return;
      const dumpEl = $('corebridgeDataDump');
      if(!dumpEl) {
        renderCorebridgeDataDump(String(text == null ? '' : text));
        return;
      }
      const next = String(text == null ? '' : text);
      dumpEl.value = dumpEl.value ? (dumpEl.value + '\n\n' + next) : next;
      dumpEl.scrollTop = dumpEl.scrollHeight;
    }
    async function fetchCorebridgeQuoteLevelData(options) {
      const opts = options || {};
      const orderId = String(opts.cbOrderId == null ? '' : opts.cbOrderId).trim();
      const accountId = String(opts.cbAccountId == null ? '' : opts.cbAccountId).trim();
      const accountName = String(opts.cbAccountName == null ? '' : opts.cbAccountName).trim();
      if(!orderId || !accountId || !accountName) throw new Error('Missing cbOrderId/cbAccountId/cbAccountName.');

      const fixedUrl =
        corebridgeProxyBaseUrl +
        '/CB_OrderData_QuoteLevel?orderId=' +
        encodeURIComponent(orderId) +
        '&accountId=' +
        encodeURIComponent(accountId) +
        '&accountName=' +
        encodeURIComponent(accountName);

      const res = await fetchWithTimeout(fixedUrl, {
        method: 'GET',
        cache: 'no-store',
        credentials: 'omit'
      });
      const text = await res.text();
      if(!res.ok) throw new Error('HTTP ' + res.status + ' ' + res.statusText);
      let data = null;
      try {data = JSON.parse(text);} catch(_eQJson) {data = text;}
      return {url: fixedUrl, status: res.status, statusText: res.statusText, data: data};
    }
    async function fetchCorebridgeProductNotesAll(options) {
      const opts = options || {};
      const orderProductId = String(opts.orderProductId == null ? '' : opts.orderProductId).trim();
      if(!orderProductId) throw new Error('Missing orderProductId.');
      const url =
        corebridgeProxyBaseUrl +
        '/CB_ProductNotesAll?orderProductId=' +
        encodeURIComponent(orderProductId);
      const res = await fetchWithTimeout(url, {
        method: 'GET',
        cache: 'no-store',
        credentials: 'omit'
      });
      const text = await res.text();
      if(!res.ok) throw new Error('HTTP ' + res.status + ' ' + res.statusText);
      let data = null;
      try {data = JSON.parse(text);} catch(_eNJson) {data = text;}
      return {orderProductId: orderProductId, notesByType: data};
    }
    async function executeCorebridgeSecondaryFetches(options) {
      const opts = options || {};
      const plan = opts.plan || {quoteLevelCalls: [], productNotesCalls: []};
      const quoteLevelCalls = Array.isArray(plan.quoteLevelCalls) ? plan.quoteLevelCalls : [];
      const productNotesCalls = Array.isArray(plan.productNotesCalls) ? plan.productNotesCalls : [];
      const results = {
        quoteLevel: [],
        productNotes: []
      };

      for(let i = 0; i < quoteLevelCalls.length; i++) {
        const call = quoteLevelCalls[i];
        try {
          const quoteData = await fetchCorebridgeQuoteLevelData(call);
          results.quoteLevel.push({
            request: call,
            ok: true,
            response: quoteData
          });
        } catch(err) {
          results.quoteLevel.push({
            request: call,
            ok: false,
            error: (err && err.message) ? err.message : String(err)
          });
        }
      }

      for(let j = 0; j < productNotesCalls.length; j++) {
        const call = productNotesCalls[j];
        try {
          const notesData = await fetchCorebridgeProductNotesAll(call);
          results.productNotes.push({
            request: call,
            ok: true,
            response: notesData
          });
        } catch(err) {
          results.productNotes.push({
            request: call,
            ok: false,
            error: (err && err.message) ? err.message : String(err)
          });
        }
      }

      return results;
    }
    function buildCorebridgeSecondaryFetchResultsLog(results) {
      const r = results || {quoteLevel: [], productNotes: []};
      const quoteRows = Array.isArray(r.quoteLevel) ? r.quoteLevel : [];
      const notesRows = Array.isArray(r.productNotes) ? r.productNotes : [];
      const lines = [];
      lines.push('--- Secondary Fetch Results ---');
      lines.push('');
      lines.push('Quote-level responses:');
      if(!quoteRows.length) lines.push('None');
      quoteRows.forEach((row, idx) => {
        const req = row && row.request ? row.request : {};
        lines.push((idx + 1) + '. getOrderData_QuoteLevel(' + req.cbOrderId + ', ' + req.cbAccountId + ', "' + String(req.cbAccountName || '') + '")');
        if(row && row.ok) {
          lines.push('   Status: OK');
          lines.push('   Data: ' + JSON.stringify(row.response && row.response.data != null ? row.response.data : null, null, 2));
        } else {
          lines.push('   Status: ERROR');
          lines.push('   Error: ' + (row && row.error ? row.error : 'Unknown error'));
        }
      });
      lines.push('');
      lines.push('Product notes responses:');
      if(!notesRows.length) lines.push('None');
      notesRows.forEach((row, idx) => {
        const req = row && row.request ? row.request : {};
        lines.push((idx + 1) + '. getProductNotesAll(' + req.orderProductId + ')');
        if(row && row.ok) {
          lines.push('   Status: OK');
          lines.push('   Data: ' + JSON.stringify(row.response && row.response.notesByType != null ? row.response.notesByType : null, null, 2));
        } else {
          lines.push('   Status: ERROR');
          lines.push('   Error: ' + (row && row.error ? row.error : 'Unknown error'));
        }
      });
      return lines.join('\n');
    }
    function readValueAtPath(obj, path) {
      if(!obj || !path) return undefined;
      const parts = String(path).split('.');
      let cur = obj;
      for(let i = 0; i < parts.length; i++) {
        const key = String(parts[i] || '').trim();
        if(!key) continue;
        if(cur == null || typeof cur !== 'object' || !(key in cur)) return undefined;
        cur = cur[key];
      }
      return cur;
    }
    function parseKoStorageVariable(koString) {
      return String(koString == null ? '' : koString).split('~').join(' ').split('^').join('"');
    }
    function formatDateDdMmYy(inputDate) {
      const d = inputDate instanceof Date ? inputDate : new Date();
      const dd = String(d.getDate()).padStart(2, '0');
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const yy = String(d.getFullYear()).slice(-2);
      return dd + '/' + mm + '/' + yy;
    }
    function tryParseInstallAddress(rawValue) {
      if(rawValue == null) return '';
      if(typeof rawValue === 'object') {
        const direct = rawValue.formattedInstallAddress;
        return direct != null ? String(direct) : '';
      }
      const txt = String(rawValue).trim();
      if(!txt) return '';
      try {
        const parsedDirect = JSON.parse(txt);
        if(parsedDirect && parsedDirect.formattedInstallAddress != null) return String(parsedDirect.formattedInstallAddress);
      } catch(_eDirect) { }
      try {
        const parsedKo = JSON.parse(parseKoStorageVariable(txt));
        if(parsedKo && parsedKo.formattedInstallAddress != null) return String(parsedKo.formattedInstallAddress);
      } catch(_eKo) { }
      return '';
    }
    function extractInstallAddressDetailed(quoteData) {
      const debug = [];
      const m1Candidates = [
        readValueAtPath(quoteData, 'OrderInformation.OrderInformation.M1'),
        readValueAtPath(quoteData, 'OrderInformation.M1'),
        readValueAtPath(quoteData, 'M1')
      ];
      for(let i = 0; i < m1Candidates.length; i++) {
        if(m1Candidates[i] != null) debug.push('M1[' + i + ']=present');
        const parsed = tryParseInstallAddress(m1Candidates[i]);
        if(parsed) return {value: parsed, source: 'M1[' + i + ']', debug: debug};
      }
      const directCandidates = [
        readValueAtPath(quoteData, 'OrderInformation.OrderInformation.formattedInstallAddress'),
        readValueAtPath(quoteData, 'OrderInformation.formattedInstallAddress'),
        readValueAtPath(quoteData, 'formattedInstallAddress')
      ];
      for(let j = 0; j < directCandidates.length; j++) {
        if(directCandidates[j] != null && String(directCandidates[j]).trim()) {
          return {value: String(directCandidates[j]).trim(), source: 'formattedInstallAddress[' + j + ']', debug: debug};
        }
      }
      return {value: '', source: 'none', debug: debug};
    }
    function extractInstallAddress(quoteData) {
      return extractInstallAddressDetailed(quoteData).value;
    }
    function normalizeNoteItems(noteRows) {
      const out = [];
      (Array.isArray(noteRows) ? noteRows : []).forEach((row) => {
        const noteText = row && row.Note != null ? String(row.Note) : '';
        const createdBy = row && row.CreatedByName != null ? String(row.CreatedByName) : '';
        const isHidden = !!(row && row.IsHidden);
        if(isHidden) return;
        if(!noteText) return;
        out.push(createdBy ? (createdBy + ': ' + noteText) : noteText);
      });
      return out;
    }
    function groupPartTypes(items, phrases) {
      const out = {};
      const rows = Array.isArray(items) ? items : [];
      const keys = Array.isArray(phrases) ? phrases : [];
      rows.forEach((item) => {
        const raw = String(item == null ? '' : item).trim();
        if(!raw) return;
        const match = keys.find((p) => raw.toLowerCase().indexOf(String(p).toLowerCase()) === 0);
        if(!match) return;
        const after = raw.slice(String(match).length).trim();
        if(!out[match]) out[match] = [];
        if(after) out[match].push(after);
      });
      return out;
    }
    function htmlToPlainText(value) {
      const html = String(value == null ? '' : value);
      if(!html) return '';
      const brToken = '__SRH_BR__';
      try {
        const el = document.createElement('div');
        el.innerHTML = html.replace(/<br\s*\/?>/gi, brToken);

        const blockTags = ['p', 'div', 'section', 'article', 'header', 'footer', 'blockquote', 'pre', 'ul', 'ol', 'table', 'tr'];
        blockTags.forEach((tag) => {
          const nodes = el.querySelectorAll(tag);
          Array.prototype.forEach.call(nodes, (node) => {
            node.insertAdjacentText('afterend', '\n');
          });
        });

        const lis = el.querySelectorAll('li');
        Array.prototype.forEach.call(lis, (li) => {
          li.insertAdjacentText('afterbegin', '- ');
          li.insertAdjacentText('afterend', '\n');
        });

        const txt = el.textContent || el.innerText || '';
        return String(txt).replace(new RegExp(brToken, 'g'), '\n')
          .replace(/\u00a0/g, ' ')
          .replace(/\r\n/g, '\n')
          .replace(/\r/g, '\n')
          .replace(/\n{3,}/g, '\n\n')
          .replace(/[ \t]+\n/g, '\n')
          .trim();
      } catch(_eHtmlTxt) {
        return html
          .replace(/<br\s*\/?>/gi, brToken)
          .replace(/<\/(p|div|section|article|header|footer|blockquote|pre|li|ul|ol)>/gi, '\n')
          .replace(/<li[^>]*>/gi, '\n- ')
          .replace(/<[^>]+>/g, ' ')
          .replace(new RegExp(brToken, 'g'), '\n')
          .replace(/\u00a0/g, ' ')
          .replace(/\r\n/g, '\n')
          .replace(/\r/g, '\n')
          .replace(/\n{3,}/g, '\n\n')
          .replace(/[ \t]+\n/g, '\n')
          .trim();
      }
    }
    function buildSvgFromDataUrlImage(dataUrl, sizePx) {
      const sz = Number(sizePx) > 0 ? Number(sizePx) : 200;
      const href = String(dataUrl == null ? '' : dataUrl);
      if(!href) return '';
      return (
        '<svg xmlns="http://www.w3.org/2000/svg" width="' + sz + '" height="' + sz + '" viewBox="0 0 ' + sz + ' ' + sz + '">' +
        '<image href="' + href + '" x="0" y="0" width="' + sz + '" height="' + sz + '"/>' +
        '</svg>'
      );
    }
    function tryGenerateQrSvgViaDependency(text) {
      try {
        if(typeof QRCode === 'undefined') return '';
        const host = document.createElement('div');
        host.style.position = 'absolute';
        host.style.left = '-10000px';
        host.style.top = '-10000px';
        host.style.width = '200px';
        host.style.height = '200px';
        document.body.appendChild(host);
        const qrInstance = new QRCode(host, {
          text: String(text == null ? '' : text),
          width: 200,
          height: 200,
          colorDark: '#000',
          colorLight: '#fff',
          correctLevel: QRCode.CorrectLevel ? QRCode.CorrectLevel.L : undefined
        });
        const canvas = host.querySelector('canvas');
        const img = host.querySelector('img');
        let dataUrl = '';
        if(canvas && typeof canvas.toDataURL === 'function') dataUrl = canvas.toDataURL('image/png');
        else if(img && img.src) dataUrl = String(img.src);
        host.remove();
        if(!dataUrl) return '';
        return buildSvgFromDataUrlImage(dataUrl, 200);
      } catch(_eQrDep) {
        return '';
      }
    }
    function tryGenerateQrPngDataUrlViaDependency(text) {
      try {
        if(typeof QRCode === 'undefined') return '';
        const host = document.createElement('div');
        host.style.position = 'absolute';
        host.style.left = '-10000px';
        host.style.top = '-10000px';
        host.style.width = '200px';
        host.style.height = '200px';
        document.body.appendChild(host);
        const qrInstance = new QRCode(host, {
          text: String(text == null ? '' : text),
          width: 200,
          height: 200,
          colorDark: '#000',
          colorLight: '#fff',
          correctLevel: QRCode.CorrectLevel ? QRCode.CorrectLevel.L : undefined
        });
        const canvas = host.querySelector('canvas');
        const out = (canvas && typeof canvas.toDataURL === 'function')
          ? String(canvas.toDataURL('image/png'))
          : '';
        host.remove();
        return out;
      } catch(_eQrPngDep) {
        return '';
      }
    }
    async function fetchQrSvgForText(text) {
      const value = String(text == null ? '' : text).trim();
      if(!value) return '';
      const localSvg = tryGenerateQrSvgViaDependency(value);
      if(localSvg) return localSvg;
      const url = 'https://api.qrserver.com/v1/create-qr-code/?size=512x512&format=svg&data=' + encodeURIComponent(value);
      const res = await fetchWithTimeout(url, {method: 'GET', cache: 'no-store'});
      const svgText = await res.text();
      if(!res.ok) throw new Error('QR HTTP ' + res.status + ' ' + res.statusText);
      return svgText;
    }
    async function buildCorebridgeProofPayload() {
      const primaryRow = (Array.isArray(corebridgeLastFilteredData) && corebridgeLastFilteredData.length) ? corebridgeLastFilteredData[0] : {};
      const orderProductId = String(primaryRow && primaryRow.Id != null ? primaryRow.Id : '').trim();
      const lineItemOrder = parseInt((primaryRow && primaryRow.LineItemOrder != null) ? primaryRow.LineItemOrder : 1, 10) || 1;

      const secondary = corebridgeLastSecondaryFetchResults || {};
      const quoteRows = Array.isArray(secondary.quoteLevel) ? secondary.quoteLevel : [];
      const quoteSuccessRow = quoteRows.find((row) => row && row.ok && row.response && row.response.data != null) || null;
      const quoteData = quoteSuccessRow ? quoteSuccessRow.response.data : null;

      const productNotesRows = Array.isArray(secondary.productNotes) ? secondary.productNotes : [];
      let notesSuccessRow = null;
      if(orderProductId) {
        notesSuccessRow = productNotesRows.find((row) => {
          const reqId = String(row && row.request && row.request.orderProductId != null ? row.request.orderProductId : '').trim();
          return !!(row && row.ok && reqId && reqId === orderProductId);
        }) || null;
      }
      if(!notesSuccessRow) {
        notesSuccessRow = productNotesRows.find((row) => row && row.ok) || null;
      }
      const notesByTypeRaw = (notesSuccessRow && notesSuccessRow.response && Array.isArray(notesSuccessRow.response.notesByType))
        ? notesSuccessRow.response.notesByType
        : [];

      const quoteItems = readValueAtPath(quoteData, 'OrderInformation.OrderInformation.H2');
      const quoteItem = (Array.isArray(quoteItems) && quoteItems.length >= lineItemOrder) ? quoteItems[lineItemOrder - 1] : null;
      const productQty = quoteItem && quoteItem.B0 != null ? String(quoteItem.B0) : '';
      const lineItemDescriptionHtml = quoteItem && quoteItem.I1 != null ? String(quoteItem.I1) : '';
      const lineItemDescription = htmlToPlainText(lineItemDescriptionHtml);
      const partPeekViews = Array.isArray(quoteItem && quoteItem.PartPeekViews) ? quoteItem.PartPeekViews : [];
      const partNames = partPeekViews
        .map((p) => (p && p.O2 != null ? String(p.O2).trim() : ''))
        .filter(Boolean);
      const partsNumbered = partNames.map((name, idx) => (idx + 1) + ': ' + name).join('\n');
      const partsPlain = partNames.join('\n');
      const groupedPartTypes = groupPartTypes(partNames, ['Vinyl -', 'Laminate -', 'ACM -', 'Acrylic -', 'Corflute -', 'Foamed PVC -', 'Aluminium -', 'Mondoclad -', 'Polycarb -', 'Stainless -', 'HDPE -', 'Signwhite -']);
      const mediaText = Array.isArray(groupedPartTypes['Vinyl -']) ? groupedPartTypes['Vinyl -'].join('\n') : '';
      const laminateText = Array.isArray(groupedPartTypes['Laminate -']) ? groupedPartTypes['Laminate -'].join('\n') : '';
      const substratePrefixes = ['ACM -', 'Acrylic -', 'Aluminium -', 'Mondoclad -', 'Foamed PVC -', 'Corflute -', 'Polycarb -', 'Stainless -', 'HDPE -', 'Signwhite -'];
      function deriveSubstrateShortText(prefix, rawValue) {
        const shortName = String(prefix || '').replace(/\s*-\s*$/, '').trim();
        const raw = String(rawValue == null ? '' : rawValue).trim();
        if(!shortName) return '';

        let thickness = '';
        let thicknessFromFetch = false;
        const byName = srhGlobalState && srhGlobalState.corebridgePartSearchEntriesByName
          ? srhGlobalState.corebridgePartSearchEntriesByName
          : {};
        const rawNorm = raw.toLowerCase().replace(/\s+/g, ' ').trim();
        if(rawNorm && byName[rawNorm] && byName[rawNorm].thickness) {
          thickness = String(byName[rawNorm].thickness || '').trim();
          thicknessFromFetch = !!thickness;
        }
        if(!thickness && rawNorm) {
          const matchedKey = Object.keys(byName).find((k) => k.indexOf(rawNorm) >= 0 || rawNorm.indexOf(k) >= 0);
          if(matchedKey && byName[matchedKey] && byName[matchedKey].thickness) {
            thickness = String(byName[matchedKey].thickness || '').trim();
            thicknessFromFetch = !!thickness;
          }
        }
        if(!thickness) {
          const compact = raw.replace(/\s+/g, '');
          const xParts = compact.split('x');
          if(xParts.length > 1) thickness = String(xParts[xParts.length - 1] || '').trim();
          if(!thickness) {
            const m = compact.match(/(\d+(?:\.\d+)?)\s*(?:mm)?$/i);
            if(m && m[1]) thickness = m[1];
          }
        }

        if(!thickness) return shortName;
        if(thicknessFromFetch) {
          const fetchedThickness = /mm$/i.test(thickness) ? thickness : (thickness + 'mm');
          return fetchedThickness + ' ' + shortName;
        }
        let thicknessClean = String(thickness).replace(/\s+/g, '').trim();
        if(!thicknessClean) return shortName;
        thicknessClean = thicknessClean.replace(/mm$/i, '');
        thicknessClean = thicknessClean.replace(/[^0-9.]/g, '');
        if(!thicknessClean) return shortName;
        return thicknessClean + 'mm ' + shortName;
      }
      const substrateRows = [];
      substratePrefixes.forEach((prefix) => {
        const vals = Array.isArray(groupedPartTypes[prefix]) ? groupedPartTypes[prefix] : [];
        vals.forEach((v) => {
          const shortText = deriveSubstrateShortText(prefix, v);
          if(shortText) substrateRows.push(shortText);
        });
      });
      const substrateText = substrateRows.join('\n');

      const installAddressDetails = extractInstallAddressDetailed(quoteData);
      const installAddress = installAddressDetails.value;
      const addressQrUrl = installAddress
        ? ('https://www.google.com/maps/dir/?api=1&destination=' + encodeURIComponent(installAddress))
        : '';
      let addressQrSvg = '';
      let addressQrPngDataUrl = '';
      if(addressQrUrl) {
        try {
          addressQrPngDataUrl = tryGenerateQrPngDataUrlViaDependency(addressQrUrl);
          addressQrSvg = await fetchQrSvgForText(addressQrUrl);
        } catch(_eQr) { }
      }

      const noteTypeNames = {1: 'sales', 2: 'design', 3: 'production', 4: 'customer', 5: 'vendor'};
      const notesByCategory = {
        sales: [],
        design: [],
        production: [],
        customer: [],
        vendor: []
      };
      notesByTypeRaw.forEach((row) => {
        const typeId = Number(row && row.noteTypeId);
        const key = noteTypeNames[typeId];
        if(!key) return;
        const noteRows = row && row.data && Array.isArray(row.data.ProductionNotes) ? row.data.ProductionNotes : [];
        notesByCategory[key] = normalizeNoteItems(noteRows);
      });
      const allNotes = []
        .concat(notesByCategory.sales)
        .concat(notesByCategory.design)
        .concat(notesByCategory.production)
        .concat(notesByCategory.customer)
        .concat(notesByCategory.vendor);
      function buildSection(title, values) {
        const rows = Array.isArray(values) ? values : [];
        if(!rows.length) return '';
        return '---' + title + '---\n' + rows.join('\n');
      }
      const noteSections = [
        buildSection('Sales', notesByCategory.sales),
        buildSection('Design', notesByCategory.design),
        buildSection('Production', notesByCategory.production),
        buildSection('Customer', notesByCategory.customer),
        buildSection('Vendor', notesByCategory.vendor)
      ].filter(Boolean);
      const notesJoined = noteSections.join('\n\n');
      const orderNotesFallback = String((primaryRow && primaryRow.OrderNotes != null) ? primaryRow.OrderNotes : '').trim();
      const notesAllValue = notesJoined || orderNotesFallback;

      return Object.assign({}, primaryRow, {
        Secondary: {
          quoteLevel: quoteData,
          productNotesByType: notesByTypeRaw
        },
        Derived: {
          todayDate: formatDateDdMmYy(new Date()),
          lineItemDescription: lineItemDescription,
          lineItemDescriptionHtml: lineItemDescriptionHtml,
          productQty: productQty,
          installAddress: installAddress,
          addressQrUrl: addressQrUrl,
          partsNumbered: partsNumbered,
          partsPlain: partsPlain,
          partNames: partNames,
          mediaText: mediaText,
          laminateText: laminateText,
          substrateText: substrateText,
          notesSales: notesByCategory.sales.join('\n'),
          notesDesign: notesByCategory.design.join('\n'),
          notesProduction: notesByCategory.production.join('\n'),
          notesCustomer: notesByCategory.customer.join('\n'),
          notesVendor: notesByCategory.vendor.join('\n'),
          notesAll: notesAllValue
        },
        DerivedAssets: {
          addressQrSvg: addressQrSvg,
          addressQrPngDataUrl: addressQrPngDataUrl
        },
        DerivedDebug: {
          installAddressSource: installAddressDetails.source,
          installAddressTrace: installAddressDetails.debug
        }
      });
    }
    async function fetchWithTimeout(url, fetchOptions) {
      if(typeof AbortController !== 'undefined') {
        const controller = new AbortController();
        const timeoutHandle = setTimeout(() => {
          try {controller.abort();} catch(_eAbort) { }
        }, corebridgeFetchTimeoutMs);
        try {
          const requestOptions = Object.assign({}, fetchOptions || {}, {signal: controller.signal});
          return await fetch(url, requestOptions);
        } catch(err) {
          if(err && err.name === 'AbortError') {
            throw new Error('Request timed out after 20 seconds.');
          }
          throw err;
        } finally {
          clearTimeout(timeoutHandle);
        }
      }

      let timeoutHandle = null;
      const request = fetch(url, fetchOptions || {});
      const timeout = new Promise((resolve, reject) => {
        timeoutHandle = setTimeout(() => reject(new Error('Request timed out after 20 seconds.')), corebridgeFetchTimeoutMs);
      });
      try {
        return await Promise.race([request, timeout]);
      } finally {
        if(timeoutHandle) clearTimeout(timeoutHandle);
      }
    }
    async function fetchCorebridgeFilteredData(options) {
      if(corebridgeFetchPromise) return corebridgeFetchPromise;
      const opts = options || {};
      const url = corebridgePrimaryDataUrl + '?_ts=' + Date.now();
      const criteria = getCorebridgeCriteriaFromFields();
      const jobNumber = criteria.jobNumber;
      const itemNumber = criteria.itemNumber;
      if(opts.showLoading !== false) renderCorebridgeDataDump('Loading...\n' + url);
      setCorebridgeFetchLoading(true);
      corebridgeFetchPromise = (async () => {
        let res = null;
        let lastErr = null;
        for(let attempt = 1; attempt <= 2; attempt++) {
          try {
            res = await fetchWithTimeout(url, {
              method: 'GET',
              cache: 'no-store',
              headers: {
                pragma: 'no-cache',
                'cache-control': 'no-cache'
              }
            });
            if(res.ok) break;
            if(attempt === 1 && res.status >= 500) {
              await new Promise((resolve) => setTimeout(resolve, 250));
              continue;
            }
            break;
          } catch(err) {
            lastErr = err;
            if(attempt === 1) {
              await new Promise((resolve) => setTimeout(resolve, 250));
              continue;
            }
            throw err;
          }
        }
        if(!res && lastErr) throw lastErr;
        if(!res) throw new Error('No response from primary data endpoint.');
        const text = await res.text();
        if(!res.ok) throw new Error('HTTP ' + res.status + ' ' + res.statusText);

        const parsed = JSON.parse(text);
        const list = Array.isArray(parsed) ? parsed : [parsed];
        const filteredData = list.filter((row) => {
          const rowInvoice = normalizeCorebridgeInvoiceNumber(row && row.OrderInvoiceNumber);
          if(jobNumber && rowInvoice !== jobNumber) return false;
          if(itemNumber && String(row.LineItemOrder) !== itemNumber) return false;
          return true;
        });

        corebridgeLastFilteredData = filteredData;
        corebridgeHasFetchedData = true;
        corebridgeLastFetchCriteria = {jobNumber: jobNumber, itemNumber: itemNumber};
        const secondaryFetchPlan = buildCorebridgeSecondaryFetchPlan(filteredData);
        if(opts.renderDump !== false) {
          const dumpText = JSON.stringify(filteredData, null, 2) + '\n\n' + buildCorebridgeSecondaryFetchLog(secondaryFetchPlan);
          renderCorebridgeDataDump(dumpText);
        }
        let secondaryFetchResults = null;
        if(opts.executeSecondaryFetches) {
          secondaryFetchResults = await executeCorebridgeSecondaryFetches({plan: secondaryFetchPlan});
          if(opts.renderDump !== false) appendCorebridgeDataDump(buildCorebridgeSecondaryFetchResultsLog(secondaryFetchResults));
        }
        corebridgeLastSecondaryFetchResults = secondaryFetchResults;
        return {
          res: res,
          allData: list,
          filteredData: filteredData,
          jobNumber: jobNumber,
          itemNumber: itemNumber,
          secondaryFetchPlan: secondaryFetchPlan,
          secondaryFetchResults: secondaryFetchResults
        };
      })();
      try {
        return await corebridgeFetchPromise;
      } finally {
        corebridgeFetchPromise = null;
        setCorebridgeFetchLoading(false);
      }
    }
    async function executeCorebridgePullData(options) {
      const opts = options || {};
      try {
        const result = await fetchCorebridgeFilteredData({showLoading: true, renderDump: true, executeSecondaryFetches: true});
        log('Corebridge pull data: ' + result.res.status + ' ' + result.res.statusText);
        if(opts.toastOnSuccess !== false) showToast('Corebridge data pulled.', {type: 'success', title: 'Corebridge'});
        return result;
      } catch(err) {
        const msg = (err && err.message) ? err.message : String(err);
        renderCorebridgeDataDump(
          'URL: ' + corebridgePrimaryDataUrl + '\n' +
          'Fetched: ' + (new Date()).toLocaleString() + '\n\n' +
          'ERROR:\n' + msg
        );
        log('Corebridge pull data failed: ' + msg);
        if(opts.toastOnError !== false) showToast('Failed to pull Corebridge data.', {type: 'error', title: 'Corebridge'});
        throw err;
      }
    }
    if(corebridgePullData) {
      corebridgePullData.onclick = async () => {
        try {
          await executeCorebridgePullData({toastOnSuccess: true, toastOnError: true});
        } catch(_ePullClick) { }
      };
    }
    if(corebridgeOpenProof) {
      corebridgeOpenProof.onclick = () => {
        const proofPath = (corebridgeProofPath && corebridgeProofPath.value ? corebridgeProofPath.value : '').trim();
        if(!proofPath) {
          showToast('Enter a proof path first.', {type: 'warn', title: 'Corebridge'});
          return;
        }
        const safeProofPath = jsxEscapeDoubleQuoted(proofPath);
        runButtonJsxOperation('signarama_helper_corebridge_openProofPath("' + safeProofPath + '")', {logFn: log, toastTitle: 'Corebridge proof', toastMessage: 'Proof document opened.'});
      };
    }
    if(corebridgeCreateProofFromData) {
      async function runCorebridgeProofCreation(mode) {
        const criteriaNow = getCorebridgeCriteriaFromFields();
        if(corebridgeCriteriaChanged(criteriaNow) || !corebridgeHasFetchedData || !corebridgeLastFilteredData || !corebridgeLastFilteredData.length) {
          try {
            await executeCorebridgePullData({toastOnSuccess: false, toastOnError: true});
          } catch(_eAutoPull) {
            return;
          }
        }
        if(!corebridgeLastFilteredData || !corebridgeLastFilteredData.length) {
          showToast('No filtered rows to map.', {type: 'warn', title: 'Corebridge'});
          return;
        }
        const proofPath = (corebridgeProofPath && corebridgeProofPath.value ? corebridgeProofPath.value : '').trim();
        if(!proofPath) {
          showToast('Enter a proof path first.', {type: 'warn', title: 'Corebridge'});
          return;
        }
        const mappingText = (corebridgeProofMappings && corebridgeProofMappings.value ? corebridgeProofMappings.value : '').trim();
        if(!mappingText) {
          showToast('Add at least one mapping (source -> text frame name).', {type: 'warn', title: 'Corebridge'});
          return;
        }

        const safeProofPath = jsxEscapeDoubleQuoted(proofPath);
        const proofPayload = await buildCorebridgeProofPayload();
        const resolvedAddress = String(readValueAtPath(proofPayload, 'Derived.installAddress') || '');
        const hasQrSvg = !!String(readValueAtPath(proofPayload, 'DerivedAssets.addressQrSvg') || '').trim();
        const hasQrPng = !!String(readValueAtPath(proofPayload, 'DerivedAssets.addressQrPngDataUrl') || '').trim();
        const mediaText = String(readValueAtPath(proofPayload, 'Derived.mediaText') || '');
        const laminateText = String(readValueAtPath(proofPayload, 'Derived.laminateText') || '');
        const partsText = String(readValueAtPath(proofPayload, 'Derived.partsNumbered') || '');
        const notesText = String(readValueAtPath(proofPayload, 'Derived.notesAll') || '');
        const derivedDate = String(readValueAtPath(proofPayload, 'Derived.todayDate') || '');
        const addrSource = String(readValueAtPath(proofPayload, 'DerivedDebug.installAddressSource') || 'none');
        appendCorebridgeDataDump(
          '[Derived Debug] installAddress="' + resolvedAddress + '" | source=' + addrSource +
          ' | addressQrSvgGenerated=' + (hasQrSvg ? 'yes' : 'no') +
          ' | addressQrPngGenerated=' + (hasQrPng ? 'yes' : 'no') +
          ' | mediaText="' + mediaText + '"' +
          ' | laminateText="' + laminateText + '"' +
          ' | partsNumbered="' + partsText + '"' +
          ' | notesAll="' + notesText + '"' +
          ' | todayDate=' + derivedDate
        );
        log(
          'Corebridge install address debug: value="' + resolvedAddress + '"' +
          ' source=' + addrSource +
          ' qrSvg=' + (hasQrSvg ? 'yes' : 'no') +
          ' qrPng=' + (hasQrPng ? 'yes' : 'no') +
          ' mediaText="' + mediaText + '"' +
          ' laminateText="' + laminateText + '"' +
          ' partsNumbered="' + partsText + '"' +
          ' notesAll="' + notesText + '"'
        );
        const safeDataJson = jsxEscapeDoubleQuoted(JSON.stringify(proofPayload));
        const safeMappingText = jsxEscapeDoubleQuoted(mappingText);
        const flashFieldsText = (corebridgeFlashFields && corebridgeFlashFields.value ? corebridgeFlashFields.value : '').trim();
        const safeFlashFieldsText = jsxEscapeDoubleQuoted(flashFieldsText);
        const proofFnName = (mode === 'selected')
          ? 'signarama_helper_corebridge_createProofForSelected'
          : 'signarama_helper_corebridge_createProofFromData';
        const toastTitle = (mode === 'selected')
          ? 'Corebridge proof for selected'
          : 'Corebridge proof from data';
        if(mode === 'selected') {
          const a4Options = {
            rasterize: !!(a4Rasterize && a4Rasterize.checked),
            rasterizeQuality: (a4RasterizeQuality && a4RasterizeQuality.value) ? a4RasterizeQuality.value : 'high'
          };
          const safeA4Options = jsxEscapeDoubleQuoted(JSON.stringify(a4Options));
          runButtonJsxOperation(
            proofFnName + '("' + safeProofPath + '","' + safeDataJson + '","' + safeMappingText + '","' + safeA4Options + '","' + safeFlashFieldsText + '")',
            {logFn: log, toastTitle: toastTitle, onResult: (res) => {
              const txt = String(res || '').trim();
              if(!/^Error:/i.test(txt) && flashFieldsText) startCorebridgeFlashTickPolling();
              if(/^Error:/i.test(txt) || !flashFieldsText) stopCorebridgeFlashTickPolling('proof result error or no flash fields');
            }}
          );
        } else {
          runButtonJsxOperation(
            proofFnName + '("' + safeProofPath + '","' + safeDataJson + '","' + safeMappingText + '","' + safeFlashFieldsText + '")',
            {logFn: log, toastTitle: toastTitle, onResult: (res) => {
              const txt = String(res || '').trim();
              if(!/^Error:/i.test(txt) && flashFieldsText) startCorebridgeFlashTickPolling();
              if(/^Error:/i.test(txt) || !flashFieldsText) stopCorebridgeFlashTickPolling('proof result error or no flash fields');
            }}
          );
        }
      }
      corebridgeCreateProofFromData.onclick = async () => {
        await runCorebridgeProofCreation('item');
      };
      if(corebridgeCreateProofForSelected) {
        corebridgeCreateProofForSelected.onclick = async () => {
          await runCorebridgeProofCreation('selected');
        };
      }
    }
    const outlineAll = $('btnOutlineAllText');
    if(outlineAll) outlineAll.onclick = () => runButtonJsxOperation('signarama_helper_outlineAllText()', {logFn: log, toastTitle: 'Outline all text'});

    const setFillsStrokes = $('btnSetFillsStrokes');
    if(setFillsStrokes) setFillsStrokes.onclick = () => runButtonJsxOperation('signarama_helper_setAllFillsStrokes()', {logFn: log, toastTitle: 'Set fills/strokes'});

    wireDimensions();
    wireTransform();
    wireLightbox();
    wireLedLayout();
    wireLedDepiction();
    wireColours();

    const clear = $('btnClearLog');
    if(clear) clear.onclick = () => {const el = $('log'); if(el) el.textContent = ''; showToast('Console cleared.', {type: 'info', title: 'Log', duration: 2500});};
    const setGrad1090 = $('btnSetGrad1090');
    if(setGrad1090) {
      setGrad1090.onclick = () => loadJSX(() => runButtonJsxOperation(
        '((typeof signarama_helper_debugSetSelectedGradientStops1090 === "function") ? signarama_helper_debugSetSelectedGradientStops1090 : ((typeof $ !== "undefined" && $.global && typeof $.global.signarama_helper_debugSetSelectedGradientStops1090 === "function") ? $.global.signarama_helper_debugSetSelectedGradientStops1090 : function(){return "Error: set gradient 10/90 function not loaded.";}))()',
        {logFn: log, toastTitle: 'Set gradient 10/90'}
      ));
    }
    const debugGrad = $('btnDebugGrad1090');
    if(debugGrad) {
      debugGrad.onclick = () => loadJSX(() => runButtonJsxOperation(
        '((typeof signarama_helper_debugCreateGradientRect1090 === "function") ? signarama_helper_debugCreateGradientRect1090 : ((typeof $ !== "undefined" && $.global && typeof $.global.signarama_helper_debugCreateGradientRect1090 === "function") ? $.global.signarama_helper_debugCreateGradientRect1090 : function(){return "Error: debug gradient function not loaded.";}))()',
        {logFn: log, toastTitle: 'Debug gradient 10/90'}
      ));
    }
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
    const applyBtn = $('btnApplyColours');
    function refreshApplyButtonState() {
      if(!applyBtn) return;
      applyBtn.disabled = !coloursHasPendingChanges;
      applyBtn.classList.toggle('pulse-alert', !!coloursHasPendingChanges);
    }
    if(applyBtn) {
      applyBtn.onclick = () => {
        if(!coloursPendingApplyFns.length) return;
        const fns = coloursPendingApplyFns.slice(0);
        coloursPendingApplyFns = [];
        coloursHasPendingChanges = false;
        refreshApplyButtonState();
        let idx = 0;
        function runNext() {
          if(idx >= fns.length) {
            refreshColours();
            return;
          }
          const fn = fns[idx++];
          try {fn(runNext);} catch(_eApplyNext) {runNext();}
        }
        runNext();
      };
      refreshApplyButtonState();
    }
    window.refreshColoursApplyState = refreshApplyButtonState;
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
          if(s.length !== 6) return null;
          const r = parseInt(s.slice(0, 2), 16);
          const g = parseInt(s.slice(2, 4), 16);
          const b = parseInt(s.slice(4, 6), 16);
          if(![r, g, b].every(v => isFinite(v))) return null;
          return {r, g, b};
        }
        function rgbToCmykValues(r, g, b) {
          const rn = clamp(num(r), 0, 255) / 255;
          const gn = clamp(num(g), 0, 255) / 255;
          const bn = clamp(num(b), 0, 255) / 255;
          const k = 1 - Math.max(rn, gn, bn);
          if(k >= 0.9999) return {c: 0, m: 0, y: 0, k: 100};
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
          if([c, m, y, k].every(v => isFinite(v))) {
            return {c, m, y, k};
          }
          const rgb = hexToRgb(entry.hex);
          if(rgb) return rgbToCmykValues(rgb.r, rgb.g, rgb.b);
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

        coloursPendingApplyFns = [];
        coloursHasPendingChanges = false;
        if(typeof window.refreshColoursApplyState === 'function') window.refreshColoursApplyState();

        data.forEach((entry) => {
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
          row.className = 'row colour-row-shell';
          row.style.alignItems = 'center';
          row.style.gap = '10px';

          const swatch = document.createElement('div');
          swatch.style.width = '68px';
          swatch.style.height = '34px';
          swatch.style.minWidth = '68px';
          swatch.style.minHeight = '34px';
          swatch.style.maxWidth = '68px';
          swatch.style.maxHeight = '34px';
          swatch.style.flex = '0 0 68px';
          swatch.style.border = '1px solid #444';
          swatch.style.borderRadius = '6px';
          swatch.style.background = entry.hex || '#000000';
          const baseSwatchHex = entry.hex || '#000000';

          const cmyk = getDisplayCmyk(entry);
          const typeText = String(entry.type || 'fill').toUpperCase();
          const cmykLine = 'C ' + round1(cmyk.c) + '  M ' + round1(cmyk.m) + '  Y ' + round1(cmyk.y) + '  K ' + round1(cmyk.k);
          const showBlackHazard = isBlackLike(cmyk, entry.hex) && !isNearRichBlack(cmyk, 8);

          const label = document.createElement('div');
          label.style.display = 'flex';
          label.style.flexDirection = 'column';
          label.style.gap = '2px';
          label.style.minWidth = '76px';
          label.style.flex = '0 0 76px';
          const labelTop = document.createElement('div');
          labelTop.textContent = typeText;
          labelTop.style.fontSize = '12px';
          labelTop.style.fontWeight = '700';
          labelTop.style.color = '#d7d7d7';
          const labelBottom = document.createElement('div');
          labelBottom.textContent = '';
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
          const inputWraps = [];
          const floatLabels = colourEditState.mode === 'RGB' ? ['R', 'G', 'B'] : ['C', 'M', 'Y', 'K'];
          inputs.forEach((inp, idx) => {
            inp.style.width = '100%';
            inp.style.borderColor = '#252525';
            const w = document.createElement('div');
            w.className = 'colour-input-wrap';
            const fl = document.createElement('span');
            fl.className = 'colour-input-float';
            fl.textContent = floatLabels[idx] || '';
            w.appendChild(inp);
            w.appendChild(fl);
            inputWraps.push(w);
          });
          log('Colours: set inputs key=' + (entry.key || '') + ' type=' + (entry.type || '') +
            ' values=' + inputs.map(i => i.value).join(','));

          function previewSwatch() {
            if(!rowDirty) {
              swatch.style.background = baseSwatchHex;
              return;
            }
            if(colourEditState.mode === 'RGB') {
              const r = isFinite(parseFloat(inputs[0].value)) ? parseFloat(inputs[0].value) : num(entry.r);
              const g = isFinite(parseFloat(inputs[1].value)) ? parseFloat(inputs[1].value) : num(entry.g);
              const b = isFinite(parseFloat(inputs[2].value)) ? parseFloat(inputs[2].value) : num(entry.b);
              const nextHex = rgbToHex(r, g, b);
              swatch.style.background = 'linear-gradient(to right, ' + baseSwatchHex + ' 0 50%, ' + nextHex + ' 50% 100%)';
              return;
            }
            const c = isFinite(parseFloat(inputs[0].value)) ? parseFloat(inputs[0].value) : num(entry.c);
            const m = isFinite(parseFloat(inputs[1].value)) ? parseFloat(inputs[1].value) : num(entry.m);
            const y = isFinite(parseFloat(inputs[2].value)) ? parseFloat(inputs[2].value) : num(entry.y);
            const k = isFinite(parseFloat(inputs[3].value)) ? parseFloat(inputs[3].value) : num(entry.k);
            const nextHex = cmykToHex(c, m, y, k);
            swatch.style.background = 'linear-gradient(to right, ' + baseSwatchHex + ' 0 50%, ' + nextHex + ' 50% 100%)';
          }

          function applyValues(onDone) {
            const finish = () => { if(typeof onDone === 'function') onDone(); };
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
                log('Colours: apply RGB done toKey=' + colourEditState.lastEdit.toKey);
                row.classList.remove('colour-row-dirty');
                finish();
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
              if(!updated) { finish(); return; }
              swatch.style.background = cmykToHex(c, m, y, k);
              colourEditState.lastEdit = {
                mode: 'CMYK',
                type: entry.type || '',
                fromKey: entry.key || '',
                fromHex: entry.hex || '',
                toKey: cmykKey(c, m, y, k),
                c, m, y, k
              };
              log('Colours: apply CMYK done toKey=' + colourEditState.lastEdit.toKey);
              row.classList.remove('colour-row-dirty');
              finish();
            });
          }

          let rowDirty = false;
          function markRowDirty() {
            if(rowDirty) return;
            rowDirty = true;
            row.classList.add('colour-row-dirty');
            coloursHasPendingChanges = true;
            coloursPendingApplyFns.push((done) => applyValues(done));
            if(typeof window.refreshColoursApplyState === 'function') window.refreshColoursApplyState();
            previewSwatch();
          }

          inputs.forEach((inp, idx) => {
            inp.addEventListener('change', markRowDirty);
            inp.addEventListener('input', () => {
              markRowDirty();
              previewSwatch();
            });
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
          inputWraps.forEach(w => row.appendChild(w));
          if(showBlackHazard) {
            const hazard = document.createElement('span');
            hazard.textContent = '\u26A0';
            hazard.title = 'Black colour is not close to target rich black (C60 M60 Y60 K100).';
            hazard.style.color = '#ffba00';
            hazard.style.fontWeight = '800';
            hazard.style.fontSize = '16px';
            hazard.style.marginLeft = '2px';
            hazard.style.flex = '0 0 auto';
            row.appendChild(hazard);
          }
          list.appendChild(row);
        });

        if(debug) {
          const dbgText = 'Debug: total=' + (debug.totalItems || 0) +
            ', scanned=' + (debug.scanned || 0) +
            ', path=' + (debug.pathItems || 0) +
            ', text=' + (debug.textFrames || 0) +
            ', docPaths=' + (debug.totalPathItems || 0) +
            ', docText=' + (debug.totalTextFrames || 0) +
            ', fallback=' + (debug.fallbackUsed ? 'yes' : 'no') +
            ', samples=' + (debug.sampleTypes || []).join(', ');
          log('Colours debug: ' + dbgText);
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
    const dimSelectionHint = $('dimSelectionHint');
    const textColorInput = $('textColor');
    const lineColorInput = $('lineColor');
    let refreshDimensionSelectionHintTimer = null;

    function runDimensionOperation(fnCall, options) {
      const opts = options || {};
      const loadingToast = showToast(opts.loadingMessage || 'Working...', {
        type: 'info',
        title: opts.loadingTitle || 'Dimensions',
        persistent: true,
        spinner: true
      });
      const userOnResult = opts.onResult;
      const runOpts = {};
      Object.keys(opts).forEach((k) => {runOpts[k] = opts[k];});
      runOpts.onResult = function(res) {
        try {if(loadingToast && loadingToast.close) loadingToast.close();} catch(_eToastClose) { }
        if(userOnResult) userOnResult(res);
      };
      runButtonJsxOperation(fnCall, runOpts);
    }

    if(dimClear) {
      dimClear.onclick = () => runDimensionOperation('atlas_dimensions_clear()', {
        logFn: logDim,
        toastTitle: 'Dimensions',
        toastMessage: 'Finished clearing dimensions.',
        loadingTitle: 'Dimensions',
        loadingMessage: 'Clearing dimensions...',
        onResult: function() {
          if(scheduleDimensionSelectionHintRefresh) scheduleDimensionSelectionHintRefresh();
        }
      });
    }

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
      runDimensionOperation('atlas_dimensions_runMulti("' + json + '")', {
        logFn: logDim,
        toastTitle: 'Dimensions',
        toastMessage: 'Finished adding dimensions.',
        loadingTitle: 'Dimensions',
        loadingMessage: 'Adding dimensions...',
        onResult: function() {
          if(scheduleDimensionSelectionHintRefresh) scheduleDimensionSelectionHintRefresh();
        }
      });
    }

    function runLineMeasure() {
      const payload = buildDimensionPayload();
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      runDimensionOperation('atlas_dimensions_runLine("' + json + '")', {
        logFn: logDim,
        toastTitle: 'Dimensions',
        toastMessage: 'Finished measuring line/path.',
        loadingTitle: 'Dimensions',
        loadingMessage: 'Measuring line/path...',
        onResult: function() {
          if(scheduleDimensionSelectionHintRefresh) scheduleDimensionSelectionHintRefresh();
        }
      });
    }
    function runLineMeasureReplace() {
      const payload = buildDimensionPayload();
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      runDimensionOperation('atlas_dimensions_runLineReplace("' + json + '")', {
        logFn: logDim,
        toastTitle: 'Dimensions',
        toastMessage: 'Finished measuring and replacing line/path.',
        loadingTitle: 'Dimensions',
        loadingMessage: 'Measuring and replacing line/path...',
        onResult: function() {
          if(scheduleDimensionSelectionHintRefresh) scheduleDimensionSelectionHintRefresh();
        }
      });
    }
    function runInnerAngles() {
      const payload = buildDimensionPayload();
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      runDimensionOperation('atlas_dimensions_runAnglesInner("' + json + '")', {
        logFn: logDim,
        toastTitle: 'Dimensions',
        toastMessage: 'Finished measuring inner angles.',
        loadingTitle: 'Dimensions',
        loadingMessage: 'Measuring inner angles...',
        onResult: function() {
          if(scheduleDimensionSelectionHintRefresh) scheduleDimensionSelectionHintRefresh();
        }
      });
    }
    function runOuterAngles() {
      const payload = buildDimensionPayload();
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      runDimensionOperation('atlas_dimensions_runAnglesOuter("' + json + '")', {
        logFn: logDim,
        toastTitle: 'Dimensions',
        toastMessage: 'Finished measuring outer angles.',
        loadingTitle: 'Dimensions',
        loadingMessage: 'Measuring outer angles...',
        onResult: function() {
          if(scheduleDimensionSelectionHintRefresh) scheduleDimensionSelectionHintRefresh();
        }
      });
    }
    function runAreaMeasure() {
      const payload = buildDimensionPayload();
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      runDimensionOperation('atlas_dimensions_runArea("' + json + '")', {
        logFn: logDim,
        toastTitle: 'Dimensions',
        toastMessage: 'Finished measuring area.',
        loadingTitle: 'Dimensions',
        loadingMessage: 'Measuring area...',
        onResult: function() {
          if(scheduleDimensionSelectionHintRefresh) scheduleDimensionSelectionHintRefresh();
        }
      });
    }

    function scheduleRefreshDimensionSelectionHint() {
      if(refreshDimensionSelectionHintTimer) clearTimeout(refreshDimensionSelectionHintTimer);
      refreshDimensionSelectionHintTimer = setTimeout(function() {
        refreshDimensionSelectionHintTimer = null;
        log('Dimensions refresh tick.');
        refreshDimensionSelectionHint();
      }, 180);
    }
    scheduleDimensionSelectionHintRefresh = scheduleRefreshDimensionSelectionHint;

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
      btnLineMeasure: () => runLineMeasure(),
      btnLineMeasureReplace: () => runLineMeasureReplace(),
      btnAngleInner: () => runInnerAngles(),
      btnAngleOuter: () => runOuterAngles(),
      btnAreaMeasure: () => runAreaMeasure()
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

    document.addEventListener('visibilitychange', () => {
      if(!document.hidden) scheduleRefreshDimensionSelectionHint();
    });
    window.addEventListener('focus', () => {
      scheduleRefreshDimensionSelectionHint();
    });
    refreshDimensionSelectionHint();
  }

  function wireTransform() {
    const originField = $('transformOrigin');
    const originButtons = $all('#transformOriginGrid .origin-btn');
    const applyBtn = $('btnTransformApply');
    const modeField = $('transformMode');
    const widthField = $('transformWidthMm');
    const heightField = $('transformHeightMm');
    const excludeStrokeField = $('transformExcludeStroke');
    const artboardsWrap = $('transformArtboardsWrap');
    const artboardsList = $('transformArtboardsList');
    const refreshArtboardsBtn = $('btnTransformRefreshArtboards');
    const artboardIndicesField = $('transformArtboardIndices');

    function parseArtboardIndicesField() {
      const raw = String((artboardIndicesField && artboardIndicesField.value) || '').trim();
      if(!raw) return [];
      return raw.split(',')
        .map(v => parseInt(String(v).trim(), 10))
        .filter(v => Number.isFinite(v) && v >= 0);
    }
    function setArtboardIndicesField(indices) {
      const unique = [];
      indices.forEach((n) => {
        if(!Number.isFinite(n) || n < 0) return;
        if(unique.indexOf(n) === -1) unique.push(n);
      });
      unique.sort((a, b) => a - b);
      if(artboardIndicesField) artboardIndicesField.value = unique.join(',');
      if(schedulePanelSettingsSave) schedulePanelSettingsSave();
    }
    function readCheckedArtboardsFromList() {
      if(!artboardsList) return [];
      return Array.prototype.slice.call(artboardsList.querySelectorAll('input[type="checkbox"][data-artboard-index]'))
        .filter(cb => cb.checked)
        .map(cb => parseInt(cb.getAttribute('data-artboard-index'), 10))
        .filter(v => Number.isFinite(v) && v >= 0);
    }
    function renderArtboardList(items) {
      if(!artboardsList) return;
      artboardsList.innerHTML = '';
      const saved = parseArtboardIndicesField();
      const savedSet = {};
      saved.forEach(i => {savedSet[i] = true;});
      if(!items || !items.length) {
        const empty = document.createElement('div');
        empty.className = 'artboards-empty';
        empty.textContent = 'No artboards found.';
        artboardsList.appendChild(empty);
        return;
      }
      items.forEach((ab) => {
        const idx = parseInt(ab.index, 10);
        if(!Number.isFinite(idx) || idx < 0) return;
        const row = document.createElement('label');
        row.className = 'chk';
        row.style.marginBottom = '4px';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.setAttribute('data-artboard-index', String(idx));
        cb.checked = !!savedSet[idx];
        const text = document.createElement('span');
        const name = String(ab.name || '').trim() || ('Artboard ' + (idx + 1));
        const size = (Number.isFinite(ab.widthMm) && Number.isFinite(ab.heightMm))
          ? (' (' + (Math.round(ab.widthMm * 100) / 100) + ' x ' + (Math.round(ab.heightMm * 100) / 100) + ' mm)')
          : '';
        text.textContent = '#' + (idx + 1) + ' ' + name + size + (ab.isActive ? ' [active]' : '');
        row.appendChild(cb);
        row.appendChild(text);
        cb.addEventListener('change', () => {
          setArtboardIndicesField(readCheckedArtboardsFromList());
        });
        artboardsList.appendChild(row);
      });
      // First load convenience: if nothing saved, preselect active artboard entry.
      if(!parseArtboardIndicesField().length) {
        const active = Array.prototype.slice.call(artboardsList.querySelectorAll('input[type="checkbox"][data-artboard-index]')).find((cb) => {
          const t = cb.parentElement && cb.parentElement.textContent;
          return t && t.indexOf('[active]') >= 0;
        });
        if(active) {
          active.checked = true;
          setArtboardIndicesField(readCheckedArtboardsFromList());
        }
      }
    }
    function refreshTransformArtboards() {
      if(!artboardsList) return;
      function loadArtboardDebug(reason) {
        callJSX('((typeof signarama_helper_transform_debugArtboards === "function") ? signarama_helper_transform_debugArtboards : ((typeof $ !== "undefined" && $.global && typeof $.global.signarama_helper_transform_debugArtboards === "function") ? $.global.signarama_helper_transform_debugArtboards : function(){return "{\\"error\\":\\"debug function not loaded\\"}";}))()', (dbgRes) => {
          const dbgRaw = String(dbgRes || '');
          log('Transform artboard debug (' + reason + '): ' + dbgRaw);
          if(artboardsList && artboardsList.innerHTML.indexOf('No artboards found.') >= 0) {
            artboardsList.innerHTML = '<div class="artboards-empty">No artboards found. See panel log for debug details.</div>';
          }
        });
      }
      callJSX('((typeof signarama_helper_transform_listArtboards === "function") ? signarama_helper_transform_listArtboards : ((typeof $ !== "undefined" && $.global && typeof $.global.signarama_helper_transform_listArtboards === "function") ? $.global.signarama_helper_transform_listArtboards : function(){return "Error: artboard list function not loaded.";}))()', (res) => {
        const raw = String(res || '');
        if(raw.indexOf('Error:') === 0 || raw.indexOf('EvalScript error') === 0) {
          log('Transform artboard list failed: ' + raw);
          if(artboardsList) artboardsList.innerHTML = '<div class="artboards-empty">Failed to load artboards.</div>';
          loadArtboardDebug('list-error');
          return;
        }
        let items = [];
        try {
          const parsed = JSON.parse(raw || '[]');
          if(Array.isArray(parsed)) items = parsed;
        } catch(_eList) {
          try {
            const migrated = Function('return ' + raw)();
            if(Array.isArray(migrated)) items = migrated;
          } catch(_eListLegacy) {
            items = [];
          }
        }
        renderArtboardList(items);
        if(!items.length) loadArtboardDebug('empty-list');
      });
    }
    function updateTransformModeUi() {
      const modeVal = String((modeField && modeField.value) || 'selection');
      if(artboardsWrap) artboardsWrap.classList.toggle('hidden', modeVal !== 'artboards');
      if(modeVal === 'artboards') refreshTransformArtboards();
    }

    function syncOriginButtons(originCode) {
      const val = String(originCode || 'C');
      originButtons.forEach(btn => {
        btn.classList.toggle('active', btn.getAttribute('data-origin') === val);
      });
    }
    function setOrigin(originCode) {
      const val = String(originCode || 'C');
      if(originField) originField.value = val;
      syncOriginButtons(val);
      try {if(originField) originField.dispatchEvent(new Event('change', {bubbles: true}));} catch(_eOrgChg) { }
      if(schedulePanelSettingsSave) schedulePanelSettingsSave();
    }

    originButtons.forEach(btn => {
      btn.addEventListener('click', () => setOrigin(btn.getAttribute('data-origin')));
    });
    if(originField) {
      originField.addEventListener('change', () => syncOriginButtons(originField.value));
    }
    syncOriginButtons((originField && originField.value) || 'C');
    if(modeField) modeField.addEventListener('change', updateTransformModeUi);
    if(refreshArtboardsBtn) refreshArtboardsBtn.addEventListener('click', refreshTransformArtboards);
    updateTransformModeUi();

    if(!applyBtn) return;
    applyBtn.onclick = () => {
      const widthRaw = String((widthField && widthField.value) || '').trim();
      const heightRaw = String((heightField && heightField.value) || '').trim();
      if(!widthRaw && !heightRaw) {
        showToast('Enter width and/or height in mm.', {type: 'warn', title: 'Transform'});
        return;
      }

      const payload = {
        mode: (modeField && modeField.value) || 'selection',
        widthSpec: widthRaw || null,
        heightSpec: heightRaw || null,
        origin: (originField && originField.value) || 'C',
        excludeStroke: !!(excludeStrokeField && excludeStrokeField.checked),
        artboardIndices: parseArtboardIndicesField()
      };
      if(payload.mode === 'artboards' && (!payload.artboardIndices || !payload.artboardIndices.length)) {
        showToast('Select one or more target artboards in the list.', {type: 'warn', title: 'Transform'});
        return;
      }
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      loadJSX(function() {
        runButtonJsxOperation('((typeof atlas_transform_makeSize === "function") ? atlas_transform_makeSize : ((typeof $ !== "undefined" && $.global && typeof $.global.atlas_transform_makeSize === "function") ? $.global.atlas_transform_makeSize : function(){return "Error: Transform function not loaded.";}))("' + json + '")', {logFn: log, toastTitle: 'Transform'});
      });
    };
  }

  function loadJSX(done) {
    try {
      var extDir = cs.getSystemPath(SystemPath.EXTENSION).replace(/\\/g, '/');
      var jsxPath = extDir + '/jsx/hostscript.jsx';
      var cmd = 'try{' +
        'var f=new File("' + jsxPath + '");' +
        'if(!f.exists) {' +
        '"ERR: Missing JSX file at " + f.fsName;' +
        '} else {' +
        '$.evalFile(f);' +
        '"OK: fit=" + (typeof signarama_helper_fitArtboardToArtwork) + ' +
        '", settingsSave=" + ((typeof signarama_helper_panelSettingsSave==="function" || (typeof $!=="undefined" && $.global && typeof $.global.signarama_helper_panelSettingsSave==="function")) ? "function" : "undefined") + ' +
        '", transform=" + ((typeof atlas_transform_makeSize==="function" || (typeof $!=="undefined" && $.global && typeof $.global.atlas_transform_makeSize==="function")) ? "function" : "undefined") + ' +
        '", debugGrad=" + ((typeof signarama_helper_debugCreateGradientRect1090==="function" || (typeof $!=="undefined" && $.global && typeof $.global.signarama_helper_debugCreateGradientRect1090==="function")) ? "function" : "undefined") + ' +
        '", setGrad1090=" + ((typeof signarama_helper_debugSetSelectedGradientStops1090==="function" || (typeof $!=="undefined" && $.global && typeof $.global.signarama_helper_debugSetSelectedGradientStops1090==="function")) ? "function" : "undefined");' +
        '}' +
        '}catch(e){ "ERR: " + e; }';
      cs.evalScript(cmd, function(res) {
        var txt = String(res || '');
        log('JSX load result: ' + txt);
        if(done) done(txt);
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
