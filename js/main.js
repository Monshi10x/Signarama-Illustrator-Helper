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
    if(el) el.setAttribute('data-host-status', txt || '');
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
    const opts = options || {};
    const type = opts.type || 'info';
    // Keep the panel quiet during normal work. The log still records operation
    // results; only messages which need the user's attention become toasts.
    if(type !== 'error' && type !== 'warn') return null;
    ensureToastUi();
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

  function summarizeSvgChildrenFromText(svgText) {
    try {
      const parser = new DOMParser();
      const parsed = parser.parseFromString(String(svgText || ''), 'image/svg+xml');
      const root = parsed && parsed.documentElement;
      if(!root || !root.childNodes) return 'no-root';
      const summary = [];
      Array.prototype.slice.call(root.childNodes).forEach((child, index) => {
        if(!child || !child.tagName) return;
        const tag = String(child.tagName || '');
        const id = child.getAttribute ? String(child.getAttribute('id') || '') : '';
        const childCount = child.childNodes ? child.childNodes.length : 0;
        summary.push('#' + index + ':' + tag + (id ? ('#' + id) : '') + '[children=' + childCount + ']');
      });
      return summary.length ? summary.join(', ') : 'no-element-children';
    } catch(_eSvgSummary0) {
      return 'summary-failed: ' + (_eSvgSummary0 && _eSvgSummary0.message ? _eSvgSummary0.message : _eSvgSummary0);
    }
  }

  const colourEditState = {
    lastEdit: null,
    mode: 'CMYK'
  };
  let lightboxMeasureLiveTimer = null;
  let lightboxMeasureLiveInFlight = false;
  let lightboxMeasureLiveDocumentKey = '';
  let stopLightboxMeasureLiveFn = null;
  let corebridgeFlashTickPollTimer = null;
  let corebridgeFlashTickPollCount = 0;
  let corebridgeFlashLastTickCount = -1;
  let corebridgeFlashTickPollInFlight = false;
  let stopCorebridgeFlashTickPollingFn = null;
  let coloursPendingApplyFns = [];
  let coloursHasPendingChanges = false;
  let copiedColourValues = null;
  let isLargeArtboard = false;
  let refreshLightboxArtboardScaleNotice = null;
  let activateTabFn = null;
  let persistSettingsTimer = null;
  let isApplyingPanelSettings = false;
  let schedulePanelSettingsSave = null;
  let scheduleDimensionSelectionHintRefresh = null;
  let scheduleCorebridgeInitialFetch = null;
  let scheduleScriptListRefresh = null;
  const SRH_NEST_BACKUP_KEY = 'srhNestResultBackup';
  const SRH_NEST_FORCE_TAB_KEY = 'srhNestForceTabRestore';
  const SRH_NEST_RUNTIME_KEY = 'srhNestRuntimeState';

  function stopLiveHostTimers(reason) {
    if(stopCorebridgeFlashTickPollingFn) stopCorebridgeFlashTickPollingFn(reason || 'panel cleanup');
    if(stopLightboxMeasureLiveFn) stopLightboxMeasureLiveFn();
  }

  function saveNestRuntimeMarker(eventName, extra) {
    try {
      if(typeof localStorage === 'undefined') return;
      const current = readNestRuntimeMarker() || {};
      const isNestEvent = /^(nest-|result-)/.test(String(eventName || ''));
      const payload = Object.assign({
        event: String(eventName || ''),
        at: Date.now(),
        lastNestEvent: isNestEvent ? String(eventName || '') : (current.lastNestEvent || ''),
        lastNestAt: isNestEvent ? Date.now() : Number(current.lastNestAt || 0)
      }, extra || {});
      localStorage.setItem(SRH_NEST_RUNTIME_KEY, JSON.stringify(payload));
    } catch(_eNestRuntimeSave) { }
  }

  function readNestRuntimeMarker() {
    try {
      if(typeof localStorage === 'undefined') return null;
      const raw = localStorage.getItem(SRH_NEST_RUNTIME_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch(_eNestRuntimeRead) {
      return null;
    }
  }

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
      if(el.id === 'automaticUpdatesEnabled' || el.id === 'updateChannel') return;
      if(el.id === 'dimensionThemeProfileSelect') return;
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
        if(key === 'automaticUpdatesEnabled' || key === 'updateChannel') return;
        if(key === 'dimensionThemeProfileSelect') return;
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
      if(el.id === 'automaticUpdatesEnabled' || el.id === 'updateChannel') return;
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
    const raw = {
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
    return (typeof DimensionLogic !== 'undefined') ? DimensionLogic.normalizePayload(raw) : raw;
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
    saveNestRuntimeMarker('panel-boot');

    enqueueEvalScript('app.name', function(name) {
      setStatus(name ? ('Connected: ' + name) : 'Not connected');
      log(name ? ('Connected to: ' + name) : 'Could not query app.name');
    });

    loadJSX(() => {
      enqueueEvalScript('((typeof signarama_helper_fitArtboardToArtwork==="function" || (typeof $!=="undefined" && $.global && typeof $.global.signarama_helper_fitArtboardToArtwork==="function")) ? "function" : "undefined")', function(type) {
        log('JSX check: signarama_helper_fitArtboardToArtwork is ' + type);
        if(type !== 'function') log('ERROR: JSX not loaded (check path/case).');
        loadPanelSettings(function() {
          setupPanelSettingsPersistence();
          try {
            if(typeof localStorage !== 'undefined' && localStorage.getItem(SRH_NEST_FORCE_TAB_KEY) && activateTabFn) {
              activateTabFn('tab-nest');
            }
          } catch(_eNestTabRestore) { }
          if(typeof refreshLightboxArtboardScaleNotice === 'function') refreshLightboxArtboardScaleNotice();
        });
      });
    });
  });

  if(typeof window !== 'undefined') {
    window.addEventListener('error', function(evt) {
      const msg = evt && evt.message ? evt.message : 'unknown-error';
      saveNestRuntimeMarker('window-error', {message: String(msg)});
      log('Window error: ' + msg);
    });
    window.addEventListener('unhandledrejection', function(evt) {
      const reason = evt && evt.reason;
      const text = reason && reason.message ? reason.message : String(reason || 'unknown-rejection');
      saveNestRuntimeMarker('window-unhandledrejection', {message: text});
      log('Unhandled rejection: ' + text);
    });
    window.addEventListener('beforeunload', function() {
      stopLiveHostTimers('panel unload');
      saveNestRuntimeMarker('panel-beforeunload');
    });
    window.addEventListener('pagehide', function() {
      stopLiveHostTimers('panel pagehide');
      saveNestRuntimeMarker('panel-pagehide');
    });
  }

  function wireTabs() {
    const tabs = $all('.tab[data-tab]');
    const panels = $all('[role="tabpanel"]');

    function activate(tabId) {
      if(tabId !== 'tab-lightbox' && stopLightboxMeasureLiveFn) stopLightboxMeasureLiveFn();
      if(tabId !== 'tab-corebridge' && stopCorebridgeFlashTickPollingFn) stopCorebridgeFlashTickPollingFn('tab leave');
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
      if(tabId === 'tab-corebridge' && typeof scheduleCorebridgeInitialFetch === 'function') {
        scheduleCorebridgeInitialFetch();
        callJSX('signarama_helper_corebridge_updatePageNumbers()', function(res) {
          if(/^Error:/i.test(String(res || ''))) log('Page number refresh error: ' + res);
        });
      }
      if(tabId === 'tab-scripts' && typeof scheduleScriptListRefresh === 'function') {
        scheduleScriptListRefresh();
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


  function showPanelConfirm(title, message, okText, cancelText, onOk) {
    let overlay = document.getElementById('srhConfirmOverlay');
    if(overlay) overlay.remove();
    overlay = document.createElement('div');
    overlay.id = 'srhConfirmOverlay';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:2147482000;background:rgba(0,0,0,.48);display:flex;align-items:center;justify-content:center;padding:16px;box-sizing:border-box;';
    const dialog = document.createElement('div');
    dialog.style.cssText = 'width:min(460px,100%);max-height:calc(100vh - 32px);overflow:auto;background:#171717;color:#eee;border:1px solid #454545;border-radius:12px;box-shadow:0 16px 34px rgba(0,0,0,.55);';
    const head = document.createElement('div');
    head.style.cssText = 'padding:12px 14px;border-bottom:1px solid #333;font-weight:800;background:#111;';
    head.textContent = title || 'Confirm';
    const body = document.createElement('div');
    body.style.cssText = 'padding:14px;font-size:13px;line-height:1.45;white-space:pre-wrap;';
    body.textContent = message || '';
    const actions = document.createElement('div');
    actions.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;padding:0 14px 14px 14px;';
    const cancel = document.createElement('button');
    cancel.type = 'button';
    cancel.className = 'btn2';
    cancel.style.width = 'auto';
    cancel.textContent = cancelText || 'Cancel';
    const ok = document.createElement('button');
    ok.type = 'button';
    ok.className = 'btn2 danger';
    ok.style.width = 'auto';
    ok.textContent = okText || 'OK';
    function close(runOk) {
      if(overlay && overlay.isConnected) overlay.remove();
      if(runOk && typeof onOk === 'function') onOk();
    }
    cancel.addEventListener('click', () => close(false));
    ok.addEventListener('click', () => close(true));
    actions.appendChild(cancel);
    actions.appendChild(ok);
    dialog.appendChild(head);
    dialog.appendChild(body);
    dialog.appendChild(actions);
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);
    setTimeout(() => ok.focus(), 0);
  }

  function showPanelNotice(title, message, closeText, durationMs) {
    let overlay = document.getElementById('srhNoticeOverlay');
    if(overlay) overlay.remove();
    overlay = document.createElement('div');
    overlay.id = 'srhNoticeOverlay';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:2147482000;background:rgba(0,0,0,.32);display:flex;align-items:center;justify-content:center;padding:16px;box-sizing:border-box;pointer-events:auto;';
    const dialog = document.createElement('div');
    dialog.style.cssText = 'width:min(460px,100%);max-height:calc(100vh - 32px);overflow:auto;background:#171717;color:#eee;border:1px solid #454545;border-radius:12px;box-shadow:0 16px 34px rgba(0,0,0,.55);';
    const head = document.createElement('div');
    head.style.cssText = 'padding:12px 14px;border-bottom:1px solid #333;font-weight:800;background:#111;';
    head.textContent = title || 'Notice';
    const body = document.createElement('div');
    body.style.cssText = 'padding:14px;font-size:13px;line-height:1.45;white-space:pre-wrap;';
    body.textContent = message || '';
    const actions = document.createElement('div');
    actions.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;padding:0 14px 14px 14px;';
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'btn2';
    closeBtn.style.width = 'auto';
    closeBtn.textContent = closeText || 'Close';
    let timer = null;
    function close() {
      if(timer) window.clearTimeout(timer);
      if(overlay && overlay.isConnected) overlay.remove();
    }
    closeBtn.addEventListener('click', close);
    actions.appendChild(closeBtn);
    dialog.appendChild(head);
    dialog.appendChild(body);
    dialog.appendChild(actions);
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);
    timer = window.setTimeout(close, Math.max(1000, Number(durationMs) || 5000));
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

    const conceptDistort = $('btnConceptFourPointDistort');
    const conceptDistortAutoOutlineStroke = $('conceptDistortAutoOutlineStroke');
    const conceptDistortAutoOutlineText = $('conceptDistortAutoOutlineText');
    const conceptDistortAutoGroup = $('conceptDistortAutoGroup');
    if(conceptDistort) conceptDistort.onclick = () => {
      showPanelNotice(
        '4 Point Distort: Select 4 Target Points',
        'Click 4 locations on the Illustrator document to record the target corners.\n\nClick clockwise: top-left, top-right, bottom-right, then bottom-left. This message will close automatically while capture starts.',
        'Close',
        5000
      );
      callJSX('signarama_helper_concept_beginFourPointClickCapture()', function(beginRes) {
        const beginText = String(beginRes || '');
        log(beginText);
        if(/^Error:/i.test(beginText) || /^No\b/i.test(beginText)) {
          notifyOperationResult(beginText, {toastTitle: '4 point distort'});
          return;
        }
        showToast('Click 4 document locations with the Pen tool. Waiting for the fourth click...', {type: 'info', title: '4 point distort', duration: 9000});
        let captureAttempts = 0;
        const maxCaptureAttempts = 120;
        const capturePoll = window.setInterval(function() {
          captureAttempts++;
          callJSX('signarama_helper_concept_captureFourPointClickPath()', function(captureRes) {
            const captureText = String(captureRes || '');
            if(/^WAIT:/i.test(captureText) && captureAttempts < maxCaptureAttempts) return;
            window.clearInterval(capturePoll);
            if(/^WAIT:/i.test(captureText)) {
              log('Timed out waiting for 4 clicked points.');
              notifyOperationResult('Error: Timed out waiting for 4 clicked points.', {toastTitle: '4 point distort'});
              return;
            }
            if(/^Error:/i.test(captureText) || /^No\b/i.test(captureText)) {
              log(captureText);
              notifyOperationResult(captureText, {toastTitle: '4 point distort'});
              return;
            }
            log(captureText);
            showPanelNotice(
              '4 Point Distort: Select Artwork',
              'Select artwork to be distorted. The panel is already watching and will apply the 4 point distort as soon as it detects a selection.',
              'Close',
              5000
            );
            showToast('Waiting for artwork selection...', {type: 'info', title: '4 point distort', duration: 6500});
            let applyAttempts = 0;
            const maxApplyAttempts = 60;
            const applyPoll = window.setInterval(function() {
              applyAttempts++;
              const opts = {
                autoOutlineStroke: !!(conceptDistortAutoOutlineStroke && conceptDistortAutoOutlineStroke.checked),
                autoOutlineText: !!(conceptDistortAutoOutlineText && conceptDistortAutoOutlineText.checked),
                autoGroup: !!(conceptDistortAutoGroup && conceptDistortAutoGroup.checked)
              };
              callJSX('signarama_helper_concept_applyFourPointDistort(' + JSON.stringify(opts) + ')', function(applyRes) {
                const applyText = String(applyRes || '');
                if(/^No artwork selected/i.test(applyText) && applyAttempts < maxApplyAttempts) return;
                window.clearInterval(applyPoll);
                log(applyText);
                notifyOperationResult(applyText, {toastTitle: '4 point distort'});
              });
            }, 1000);
          });
        }, 500);
      });
    };

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
        const autoCloseOpenPaths = !!(($('pathBleedAutoCloseOpen') && $('pathBleedAutoCloseOpen').checked));
        const payload = JSON.stringify({
          offsetMm: amt,
          createCutline: cut,
          outlineText: outlineText,
          outlineStroke: outlineStroke,
          autoWeld: autoWeld,
          autoCloseOpenPaths: autoCloseOpenPaths
        });
        const safe = payload.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
        runButtonJsxOperation("signarama_helper_applyPathBleed('" + safe + "')", {
          logFn: log,
          toastTitle: 'Apply path bleed',
          toastMessage: 'Bleed Successfully Added',
          toastDuration: 2000
        });
      };
    }

    const path = $('btnAddPathText');
    if(path) path.onclick = () => runButtonJsxOperation('signarama_helper_addFilePathTextToArtboards()', {logFn: log, toastTitle: 'Add path labels'});

    function isCorebridgeTabActiveForPolling() {
      const panel = $('tab-corebridge');
      return !!(panel && !panel.classList.contains('hidden'));
    }
    const corebridgePullData = $('btnCorebridgePullData');
    const corebridgeOpenProof = $('btnCorebridgeOpenProof');
    const corebridgeCreateProofFromData = $('btnCorebridgeCreateProofFromData');
    const corebridgeCreateProofForSelected = $('btnCorebridgeCreateProofForSelected');
    const corebridgeDevMode = $('corebridgeDevMode');
    const corebridgeDevDumpWrap = $('corebridgeDevDumpWrap');
    const corebridgeProofMappingsWrap = $('corebridgeProofMappingsWrap');
    const corebridgeFlashFieldsWrap = $('corebridgeFlashFieldsWrap');
    const corebridgeDerivedMappingsWrap = $('corebridgeDerivedMappingsWrap');
    const corebridgeJobNumber = $('corebridgeJobNumber');
    const corebridgeItemNumber = $('corebridgeItemNumber');
    const corebridgeProofPath = $('corebridgeProofPath');
    if(corebridgeProofPath && !corebridgeProofPath.value) {
      const proofTemplateFilename = corebridgeProofPath.getAttribute('data-template-filename') || 'PROOF TEMPLATE - Landscape  v2.ai';
      const extensionPath = cs.getSystemPath(SystemPath.EXTENSION).replace(/[\\/]+$/, '');
      corebridgeProofPath.value = extensionPath + '/Proof Templates/' + proofTemplateFilename;
    }
    const corebridgeProofMappings = $('corebridgeProofMappings');
    const corebridgeFlashFields = $('corebridgeFlashFields');
    const corebridgeDerivedMappingsPreview = document.querySelector('textarea[data-corebridge-derived-mappings]');
    const corebridgeDumpHost = $('corebridgeDataDumpHost');
    const corebridgeFetchStatus = $('corebridgeFetchStatus');
    const corebridgeLookup = $('corebridgeLookup');
    const corebridgeLookupSearch = $('corebridgeLookupSearch');
    function stopCorebridgeFlashTickPolling(reason) {
      if(corebridgeFlashTickPollTimer) {
        clearInterval(corebridgeFlashTickPollTimer);
        corebridgeFlashTickPollTimer = null;
      }
      if(reason) log('Corebridge flash poll stop: ' + reason);
      corebridgeFlashLastTickCount = -1;
      corebridgeFlashTickPollInFlight = false;
    }
    stopCorebridgeFlashTickPollingFn = stopCorebridgeFlashTickPolling;
    function startCorebridgeFlashTickPolling() {
      stopCorebridgeFlashTickPolling('restart');
      corebridgeFlashTickPollCount = 0;
      log('Corebridge flash poll start (safe tick every 500ms).');
      corebridgeFlashTickPollTimer = setInterval(() => {
        if(!isCorebridgeTabActiveForPolling()) {
          stopCorebridgeFlashTickPolling('tab leave');
          return;
        }
        if(corebridgeFlashTickPollInFlight) return;
        corebridgeFlashTickPollInFlight = true;
        corebridgeFlashTickPollCount++;
        if(corebridgeFlashTickPollCount > 120) {
          corebridgeFlashTickPollInFlight = false;
          stopCorebridgeFlashTickPolling('hard timeout');
          return;
        }
        callJSX('((typeof signarama_helper_corebridge_flashTickTask === "function") ? signarama_helper_corebridge_flashTickTask : ((typeof $ !== "undefined" && $.global && typeof $.global.signarama_helper_corebridge_flashTickTask === "function") ? $.global.signarama_helper_corebridge_flashTickTask : function(){return "ERROR|flashTickTask missing";}))()', (tickRes) => {
          corebridgeFlashTickPollInFlight = false;
          const tickTxt = String(tickRes || '').trim();
          if(/^ERROR\|/i.test(tickTxt)) {
            log('Corebridge flash state error: ' + tickTxt);
            stopCorebridgeFlashTickPolling('state error');
            return;
          }
          const parts = tickTxt.split('|');
          if(parts.length >= 6 && /^STATE$/i.test(parts[0])) {
            const active = parseInt(parts[1], 10) || 0;
            const tickCount = parseInt(parts[3], 10) || 0;
            corebridgeFlashLastTickCount = tickCount;
            if(!active) {
              stopCorebridgeFlashTickPolling(parts[6] === 'DOC_CHANGED' ? 'document changed' : 'inactive');
            }
          }
        });
      }, 500);
    }

    function selectCorebridgeInputText(el) {
      if(!el || typeof el.select !== 'function') return;
      setTimeout(() => {
        try {el.select();} catch(_eSelectCorebridgeInput) { }
      }, 0);
    }
    [corebridgeJobNumber, corebridgeItemNumber].forEach((el) => {
      if(!el) return;
      el.addEventListener('focus', () => selectCorebridgeInputText(el));
      el.addEventListener('click', () => selectCorebridgeInputText(el));
    });
    function corebridgeFirstValue(row, keys) {
      if(!row || !keys || !keys.length) return '';
      for(let i = 0; i < keys.length; i++) {
        const key = keys[i];
        if(row[key] != null && String(row[key]).trim() !== '') return String(row[key]).trim();
      }
      return '';
    }
    function corebridgeLookupRowValues(row) {
      const jobNumber = normalizeCorebridgeInvoiceNumber(corebridgeFirstValue(row, ['OrderInvoiceNumber', 'InvoiceNumber', 'JobNumber', 'OrderNumber']));
      const itemNumber = normalizeCorebridgeItemNumber(corebridgeFirstValue(row, ['LineItemOrder', 'lineItemOrder', 'ItemNumber', 'LineItemNumber']));
      const companyName = corebridgeFirstValue(row, ['CompanyName', 'companyName', 'AccountName', 'CustomerName', 'CustomerCompanyName']);
      const productName = corebridgeFirstValue(row, ['ProductName', 'productName', 'ProductDescription', 'Description', 'Name', 'PartName', 'ItemName']);
      return {jobNumber: jobNumber, itemNumber: itemNumber, companyName: companyName, productName: productName};
    }
    function corebridgeLookupSearchMatches(vals, query) {
      const q = String(query == null ? '' : query).trim().toLowerCase();
      if(!q) return true;
      const haystack = [vals.jobNumber, vals.itemNumber, vals.companyName, vals.productName].join(' ').toLowerCase();
      const tokens = q.split(/\s+/).filter(Boolean);
      for(let i = 0; i < tokens.length; i++) {
        if(haystack.indexOf(tokens[i]) === -1) return false;
      }
      return true;
    }
    function renderCorebridgeLookup(rows) {
      if(!corebridgeLookup) return;
      const source = Array.isArray(rows) ? rows.slice() : [];
      const query = corebridgeLookupSearch ? corebridgeLookupSearch.value : '';
      const list = source.filter((row) => corebridgeLookupSearchMatches(corebridgeLookupRowValues(row), query));
      corebridgeLookup.innerHTML = '';
      if(!source.length || !list.length) {
        const empty = document.createElement('div');
        empty.className = 'corebridge-lookup-empty';
        empty.textContent = !source.length ? 'Fetched jobs will appear here.' : 'No lookup rows match your search.';
        corebridgeLookup.appendChild(empty);
        return;
      }
      list.sort((a, b) => {
        const av = corebridgeLookupRowValues(a);
        const bv = corebridgeLookupRowValues(b);
        return [av.companyName, av.jobNumber, av.itemNumber, av.productName].join(' ').localeCompare([bv.companyName, bv.jobNumber, bv.itemNumber, bv.productName].join(' '), undefined, {numeric: true, sensitivity: 'base'});
      });
      list.forEach((row) => {
        const vals = corebridgeLookupRowValues(row);
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'corebridge-lookup-row';
        btn.title = [vals.jobNumber, vals.itemNumber, vals.companyName, vals.productName].filter(Boolean).join(' | ');
        [vals.jobNumber || '-', vals.itemNumber || '-', vals.companyName || '-', vals.productName || '-'].forEach((text) => {
          const cell = document.createElement('span');
          cell.className = 'corebridge-lookup-cell';
          cell.textContent = text;
          btn.appendChild(cell);
        });
        btn.addEventListener('click', () => {
          if(corebridgeJobNumber) corebridgeJobNumber.value = vals.jobNumber;
          if(corebridgeItemNumber) corebridgeItemNumber.value = vals.itemNumber;
          $all('.corebridge-lookup-row.selected').forEach((el) => el.classList.remove('selected'));
          btn.classList.add('selected');
          if(corebridgeJobNumber) corebridgeJobNumber.dispatchEvent(new Event('input', {bubbles: true}));
          if(corebridgeItemNumber) corebridgeItemNumber.dispatchEvent(new Event('input', {bubbles: true}));
          useCachedCorebridgeDataForCriteria(getCorebridgeCriteriaFromFields());
        });
        corebridgeLookup.appendChild(btn);
      });
    }
    if(corebridgeLookupSearch) {
      corebridgeLookupSearch.addEventListener('input', () => renderCorebridgeLookup(corebridgeLastAllData));
    }
    function invalidateCorebridgeFetchCache() {
      corebridgeHasFetchedData = false;
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
    function corebridgeProxyBaseUrl() {
      const select = $('corebridgeProxyBaseUrl');
      return select ? select.value : 'https://signschedulerapp.ts.r.appspot.com';
    }
    function corebridgePrimaryDataUrl() {return corebridgeProxyBaseUrl() + '/CB_DesignBoard_Data';}
    function corebridgePartSearchEntriesUrl() {return corebridgeProxyBaseUrl() + '/CB_OrderEntryProducts_PartSearchEntries';}
    const corebridgeProxySelect = $('corebridgeProxyBaseUrl');
    if(corebridgeProxySelect) corebridgeProxySelect.addEventListener('change', () => {invalidateCorebridgeFetchCache(); preloadCorebridgePartSearchEntries();});
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
        const res = await fetch(corebridgePartSearchEntriesUrl() + '?_ts=' + Date.now(), {
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
    let corebridgeLastAllData = [];
    let corebridgeLastFilteredData = [];
    let corebridgeInitialFetchStarted = false;
    let corebridgeLastSecondaryFetchResults = null;
    let corebridgeHasFetchedData = false;
    let corebridgeLastFetchCriteria = {jobNumber: "", itemNumber: ""};
    let corebridgeFetchPromise = null;
    const corebridgeDerivedMappingsPreviewText = [
      'Derived.installAddress -> Address Text',
      'Derived.todayDate -> Date Text',
      'Derived.lineItemNumber -> Line Item Number',
      'Derived.itemNumber -> Item Number',
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
    function useCachedCorebridgeDataForCriteria(criteria) {
      const next = criteria || getCorebridgeCriteriaFromFields();
      if(!Array.isArray(corebridgeLastAllData) || !corebridgeLastAllData.length) return false;
      const cachedMatches = corebridgeLastAllData.filter((row) => {
        const invoice = normalizeCorebridgeInvoiceNumber(row && row.OrderInvoiceNumber);
        const item = normalizeCorebridgeItemNumber(row && (row.LineItemOrder != null ? row.LineItemOrder : row.lineItemOrder));
        return (!next.jobNumber || invoice === next.jobNumber) && (!next.itemNumber || item === next.itemNumber);
      });
      if(!cachedMatches.length) return false;
      const criteriaChanged = corebridgeCriteriaChanged(next);
      corebridgeLastFilteredData = cachedMatches;
      corebridgeHasFetchedData = true;
      corebridgeLastFetchCriteria = {jobNumber: next.jobNumber, itemNumber: next.itemNumber};
      if(criteriaChanged) corebridgeLastSecondaryFetchResults = null;
      return true;
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
        corebridgeProxyBaseUrl() +
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
        corebridgeProxyBaseUrl() +
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
      const orderProductId = String(primaryRow && primaryRow.Id != null ? primaryRow.Id : (primaryRow && primaryRow.OrderProductId != null ? primaryRow.OrderProductId : (primaryRow && primaryRow.OrderProductID != null ? primaryRow.OrderProductID : ''))).trim();
      const corebridgeLineItemOrderRaw = primaryRow && primaryRow.lineItemOrder != null
        ? primaryRow.lineItemOrder
        : (primaryRow && primaryRow.LineItemOrder != null ? primaryRow.LineItemOrder : '');
      const lineItemOrder = parseInt(corebridgeLineItemOrderRaw !== '' ? corebridgeLineItemOrderRaw : 1, 10) || 1;

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

      function readQuoteLineItems(data) {
        const direct = readValueAtPath(data, 'OrderInformation.OrderInformation.H2');
        if(Array.isArray(direct)) return direct;
        const wrapped = readValueAtPath(data, 'data.OrderInformation.OrderInformation.H2');
        if(Array.isArray(wrapped)) return wrapped;
        return [];
      }
      const quoteItems = readQuoteLineItems(quoteData);
      const quoteItemIndex = lineItemOrder - 1;
      const quoteItem = (Array.isArray(quoteItems) && quoteItemIndex >= 0 && quoteItems.length > quoteItemIndex) ? quoteItems[quoteItemIndex] : null;
      const productQty = quoteItem && quoteItem.B0 != null ? String(quoteItem.B0) : '';
      const lineItemNumber = String(corebridgeLineItemOrderRaw !== '' ? corebridgeLineItemOrderRaw : lineItemOrder);
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
      function titleCaseWords(value) {
        return String(value == null ? '' : value)
          .toLowerCase()
          .replace(/\b([a-z])/g, (m) => m.toUpperCase())
          .replace(/\bPvc\b/g, 'PVC')
          .replace(/\bHdpe\b/g, 'HDPE')
          .replace(/\bAcm\b/g, 'ACM');
      }
      function escapeRegExp(value) {
        return String(value == null ? '' : value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      }
      function normalizeLooseKey(value) {
        return String(value == null ? '' : value)
          .replace(/\u00a0/g, ' ')
          .toLowerCase()
          .replace(/&/g, ' and ')
          .replace(/[^a-z0-9.]+/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();
      }
      function cleanThicknessValue(value) {
        let out = String(value == null ? '' : value).replace(/\s+/g, '').trim();
        out = out.replace(/mm$/i, '').replace(/[^0-9.]/g, '');
        return out;
      }
      function removeWordsForSubstrateFinish(value, materialName, thicknessClean) {
        let finish = String(value == null ? '' : value).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
        finish = finish
          .replace(/\([^)]*\)/g, ' ')
          .replace(/\b\d+(?:\.\d+)?(?:\s*x\s*\d+(?:\.\d+)?)+\b/ig, ' ')
          .replace(/\b(?:sqm|sq\s*m|m2)\b/ig, ' ');
        if(thicknessClean) {
          const escapedThickness = escapeRegExp(thicknessClean);
          finish = finish
            .replace(new RegExp('(?:^|\\s|x)' + escapedThickness + '\\s*mm\\b', 'ig'), ' ')
            .replace(new RegExp('(?:^|\\s|x)' + escapedThickness + '\\b', 'ig'), ' ');
        }
        const materialWords = String(materialName || '').split(/\s+/).filter(Boolean);
        materialWords.forEach((word) => {
          finish = finish.replace(new RegExp('\\b' + escapeRegExp(word) + '\\b', 'ig'), ' ');
        });
        finish = finish.replace(/[-_]+/g, ' ').replace(/\s+/g, ' ').trim();
        return finish;
      }
      function findSubstrateThickness(raw, materialName) {
        const byName = srhGlobalState && srhGlobalState.corebridgePartSearchEntriesByName
          ? srhGlobalState.corebridgePartSearchEntriesByName
          : {};
        const rawNorm = normalizeLooseKey(raw);
        const materialNorm = normalizeLooseKey(materialName);
        const lookupKeys = [];
        if(rawNorm) lookupKeys.push(rawNorm);
        if(materialNorm && rawNorm) lookupKeys.push((materialNorm + ' ' + rawNorm).trim());
        if(materialNorm && rawNorm) lookupKeys.push((materialNorm + ' - ' + rawNorm).trim());
        for(let i = 0; i < lookupKeys.length; i++) {
          const key = lookupKeys[i];
          if(byName[key] && byName[key].thickness) return String(byName[key].thickness || '').trim();
        }
        const keys = Object.keys(byName);
        for(let k = 0; k < keys.length; k++) {
          const key = normalizeLooseKey(keys[k]);
          if(!key || !rawNorm) continue;
          const hasRaw = key.indexOf(rawNorm) >= 0 || rawNorm.indexOf(key) >= 0;
          const hasMaterial = !materialNorm || key.indexOf(materialNorm) >= 0;
          if(hasRaw && hasMaterial && byName[keys[k]] && byName[keys[k]].thickness) {
            return String(byName[keys[k]].thickness || '').trim();
          }
        }
        const compact = String(raw == null ? '' : raw).replace(/\s+/g, '');
        const xParts = compact.split('x');
        if(xParts.length > 1) return String(xParts[xParts.length - 1] || '').trim();
        const m = compact.match(/(\d+(?:\.\d+)?)\s*(?:mm)?$/i);
        return (m && m[1]) ? m[1] : '';
      }
      function deriveSubstrateShortText(prefix, rawValue) {
        const materialName = String(prefix || '').replace(/\s*-\s*$/, '').trim();
        const raw = String(rawValue == null ? '' : rawValue).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
        if(!materialName) return null;

        const thicknessClean = cleanThicknessValue(findSubstrateThickness(raw, materialName));
        const finish = removeWordsForSubstrateFinish(raw, materialName, thicknessClean);
        const materialText = titleCaseWords(materialName);
        const finishText = finish ? titleCaseWords(finish) : '';
        const parts = [];
        if(thicknessClean) parts.push(thicknessClean + 'mm');
        if(finishText) parts.push(finishText);
        parts.push(materialText);
        return {
          text: parts.join(' '),
          key: normalizeLooseKey((finishText ? finishText + ' ' : '') + materialText),
          hasThickness: !!thicknessClean
        };
      }
      const substrateRows = [];
      const substrateSeen = {};
      substratePrefixes.forEach((prefix) => {
        const vals = Array.isArray(groupedPartTypes[prefix]) ? groupedPartTypes[prefix] : [];
        vals.forEach((v) => {
          const substrateInfo = deriveSubstrateShortText(prefix, v);
          if(!substrateInfo || !substrateInfo.text || !substrateInfo.key) return;
          const existingIndex = substrateSeen[substrateInfo.key];
          if(existingIndex == null) {
            substrateSeen[substrateInfo.key] = substrateRows.length;
            substrateRows.push(substrateInfo);
          } else if(substrateInfo.hasThickness && !substrateRows[existingIndex].hasThickness) {
            substrateRows[existingIndex] = substrateInfo;
          }
        });
      });
      const substrateText = substrateRows.map((row) => row.text).join('\n');

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
          lineItemNumber: lineItemNumber,
          lineItemOrder: lineItemNumber,
          itemNumber: orderProductId,
          orderProductId: orderProductId,
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
          installAddressTrace: installAddressDetails.debug,
          selectedRow: {
            lineItemOrder: primaryRow && primaryRow.lineItemOrder != null ? primaryRow.lineItemOrder : null,
            LineItemOrder: primaryRow && primaryRow.LineItemOrder != null ? primaryRow.LineItemOrder : null,
            Id: primaryRow && primaryRow.Id != null ? primaryRow.Id : null,
            OrderProductId: primaryRow && primaryRow.OrderProductId != null ? primaryRow.OrderProductId : null,
            OrderProductID: primaryRow && primaryRow.OrderProductID != null ? primaryRow.OrderProductID : null
          },
          quoteLineItemSelection: {
            corebridgeLineItemOrderRaw: corebridgeLineItemOrderRaw,
            lineItemOrder: lineItemOrder,
            quoteItemIndex: quoteItemIndex,
            quoteItemsCount: Array.isArray(quoteItems) ? quoteItems.length : 0,
            quoteItemFound: !!quoteItem,
            quoteItemB0: quoteItem && quoteItem.B0 != null ? String(quoteItem.B0) : '',
            productQty: productQty
          }
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
    function corebridgeExtractPrimaryRows(parsed) {
      if(Array.isArray(parsed)) return parsed;
      if(!parsed || typeof parsed !== 'object') return [];

      const arrayKeys = ['data', 'Data', 'rows', 'Rows', 'items', 'Items', 'results', 'Results'];
      for(let i = 0; i < arrayKeys.length; i++) {
        const rows = parsed[arrayKeys[i]];
        if(Array.isArray(rows)) return rows;
      }

      if(parsed.OrderInvoiceNumber != null || parsed.LineItemOrder != null || parsed.OrderId != null || parsed.Id != null) {
        return [parsed];
      }

      return [];
    }
    function corebridgeParsePrimaryResponseText(text) {
      const rawText = String(text == null ? '' : text);
      let parsed = JSON.parse(rawText);

      if(typeof parsed === 'string') {
        const trimmed = parsed.trim();
        if(trimmed && (/^[\[{]/).test(trimmed)) {
          parsed = JSON.parse(trimmed);
        } else if(trimmed.indexOf('[object Object]') !== -1) {
          throw new Error(
            'Corebridge primary endpoint returned a stringified object list instead of JSON rows: "' +
            trimmed.slice(0, 160) +
            (trimmed.length > 160 ? '…' : '') +
            '". The localhost proxy/server must return the raw array/object as JSON, not String(rows) + " Entries".'
          );
        }
      }

      return parsed;
    }
    function normalizeCorebridgeItemNumber(value) {
      return String(value == null ? '' : value).trim();
    }
    function corebridgeDebugLog(label, value) {
      let output = '';
      if(typeof value === 'string') {
        output = value;
      } else {
        try {
          output = JSON.stringify(value, null, 2);
        } catch(_eStringify) {
          output = String(value);
        }
      }
      log('[Corebridge debug] ' + label + ': ' + output);
    }
    function corebridgeRowMatchesItemNumber(row, itemNumber) {
      const item = normalizeCorebridgeItemNumber(itemNumber);
      if(!item) return true;
      const lineItemOrder = normalizeCorebridgeItemNumber(row && (row.LineItemOrder != null ? row.LineItemOrder : row.lineItemOrder));
      return lineItemOrder === item;
    }
    async function fetchCorebridgeFilteredData(options) {
      if(corebridgeFetchPromise) return corebridgeFetchPromise;
      const opts = options || {};
      const url = corebridgePrimaryDataUrl() + '?_ts=' + Date.now();
      const criteria = getCorebridgeCriteriaFromFields();
      const jobNumber = criteria.jobNumber;
      const itemNumber = criteria.itemNumber;
      corebridgeDebugLog('fetch start', {
        url: url,
        rawJobNumber: corebridgeJobNumber && corebridgeJobNumber.value ? corebridgeJobNumber.value : '',
        normalizedJobNumber: jobNumber,
        rawItemNumber: corebridgeItemNumber && corebridgeItemNumber.value ? corebridgeItemNumber.value : '',
        normalizedItemNumber: itemNumber
      });
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
        corebridgeDebugLog('fetch response status', res.status + ' ' + res.statusText);
        if(!res.ok) throw new Error('HTTP ' + res.status + ' ' + res.statusText);

        const parsed = corebridgeParsePrimaryResponseText(text);
        const list = corebridgeExtractPrimaryRows(parsed);
        if(!list.length) {
          corebridgeDebugLog('empty row extraction warning', {
            parsedType: parsed === null ? 'null' : typeof parsed,
            parsedKeys: parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? Object.keys(parsed) : [],
            message: 'No Corebridge rows were found in the primary endpoint response.'
          });
        }
        const filteredData = list.filter((row) => {
          const rowInvoiceRaw = row && row.OrderInvoiceNumber;
          const rowInvoice = normalizeCorebridgeInvoiceNumber(rowInvoiceRaw);
          const rowLineItemOrder = normalizeCorebridgeItemNumber(row && (row.LineItemOrder != null ? row.LineItemOrder : row.lineItemOrder));
          const invoiceMatches = !jobNumber || rowInvoice === jobNumber;
          const itemMatches = !itemNumber || rowLineItemOrder === itemNumber;
          return invoiceMatches && itemMatches;
        });

        corebridgeLastAllData = list;
        renderCorebridgeLookup(list);
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
          'URL: ' + corebridgePrimaryDataUrl() + '\n' +
          'Fetched: ' + (new Date()).toLocaleString() + '\n\n' +
          'ERROR:\n' + msg
        );
        log('Corebridge pull data failed: ' + msg);
        if(opts.toastOnError !== false) showToast('Failed to pull Corebridge data.', {type: 'error', title: 'Corebridge'});
        throw err;
      }
    }
    async function executeCorebridgeInitialFetch() {
      if(corebridgeInitialFetchStarted) return;
      corebridgeInitialFetchStarted = true;
      try {
        await executeCorebridgePullData({toastOnSuccess: false, toastOnError: false});
        log('Corebridge proof data refreshed.');
      } catch(err) {
        log('Corebridge proof data refresh failed: ' + ((err && err.message) ? err.message : err));
      } finally {
        corebridgeInitialFetchStarted = false;
      }
    }
    scheduleCorebridgeInitialFetch = function() {
      setTimeout(executeCorebridgeInitialFetch, 0);
    };
    if(document.querySelector('.tab[data-tab="tab-corebridge"].active') || (document.querySelector('.tab[data-tab="tab-corebridge"]') && !document.getElementById('tab-corebridge').classList.contains('hidden'))) {
      scheduleCorebridgeInitialFetch();
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
        if((corebridgeCriteriaChanged(criteriaNow) || !corebridgeHasFetchedData) && !useCachedCorebridgeDataForCriteria(criteriaNow)) {
          showToast('Click Pull Data to load the current job and item first.', {type: 'warn', title: 'Corebridge'});
          return;
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

        if(!corebridgeLastSecondaryFetchResults) {
          const secondaryPlan = buildCorebridgeSecondaryFetchPlan(corebridgeLastFilteredData);
          corebridgeLastSecondaryFetchResults = await executeCorebridgeSecondaryFetches({plan: secondaryPlan});
          appendCorebridgeDataDump(buildCorebridgeSecondaryFetchResultsLog(corebridgeLastSecondaryFetchResults));
        }

        const safeProofPath = jsxEscapeDoubleQuoted(proofPath);
        const proofPayload = await buildCorebridgeProofPayload();
        const resolvedAddress = String(readValueAtPath(proofPayload, 'Derived.installAddress') || '');
        const hasQrSvg = !!String(readValueAtPath(proofPayload, 'DerivedAssets.addressQrSvg') || '').trim();
        const hasQrPng = !!String(readValueAtPath(proofPayload, 'DerivedAssets.addressQrPngDataUrl') || '').trim();
        const mediaText = String(readValueAtPath(proofPayload, 'Derived.mediaText') || '');
        const laminateText = String(readValueAtPath(proofPayload, 'Derived.laminateText') || '');
        const partsText = String(readValueAtPath(proofPayload, 'Derived.partsNumbered') || '');
        const lineItemText = String(readValueAtPath(proofPayload, 'Derived.lineItemNumber') || '');
        const itemNumberText = String(readValueAtPath(proofPayload, 'Derived.itemNumber') || '');
        const quantityText = String(readValueAtPath(proofPayload, 'Derived.productQty') || '');
        const substrateText = String(readValueAtPath(proofPayload, 'Derived.substrateText') || '');
        const notesText = String(readValueAtPath(proofPayload, 'Derived.notesAll') || '');
        const derivedDate = String(readValueAtPath(proofPayload, 'Derived.todayDate') || '');
        const addrSource = String(readValueAtPath(proofPayload, 'DerivedDebug.installAddressSource') || 'none');
        const selectedRowDebug = readValueAtPath(proofPayload, 'DerivedDebug.selectedRow') || {};
        const quoteLineItemDebug = readValueAtPath(proofPayload, 'DerivedDebug.quoteLineItemSelection') || {};
        appendCorebridgeDataDump(
          '[Derived Debug] installAddress="' + resolvedAddress + '" | source=' + addrSource +
          ' | addressQrSvgGenerated=' + (hasQrSvg ? 'yes' : 'no') +
          ' | addressQrPngGenerated=' + (hasQrPng ? 'yes' : 'no') +
          ' | mediaText="' + mediaText + '"' +
          ' | laminateText="' + laminateText + '"' +
          ' | lineItemNumber="' + lineItemText + '"' +
          ' | itemNumber="' + itemNumberText + '"' +
          ' | productQty="' + quantityText + '"' +
          ' | substrateText="' + substrateText + '"' +
          ' | partsNumbered="' + partsText + '"' +
          ' | notesAll="' + notesText + '"' +
          ' | todayDate=' + derivedDate
        );
        appendCorebridgeDataDump(
          '[Line/Qty Debug] selectedRow=' + JSON.stringify(selectedRowDebug) +
          ' | quoteLineItemSelection=' + JSON.stringify(quoteLineItemDebug) +
          ' | filteredRows=' + (Array.isArray(corebridgeLastFilteredData) ? corebridgeLastFilteredData.length : 0)
        );
        log(
          'Corebridge install address debug: value="' + resolvedAddress + '"' +
          ' source=' + addrSource +
          ' qrSvg=' + (hasQrSvg ? 'yes' : 'no') +
          ' qrPng=' + (hasQrPng ? 'yes' : 'no') +
          ' mediaText="' + mediaText + '"' +
          ' laminateText="' + laminateText + '"' +
          ' lineItemNumber="' + lineItemText + '"' +
          ' itemNumber="' + itemNumberText + '"' +
          ' productQty="' + quantityText + '"' +
          ' substrateText="' + substrateText + '"' +
          ' partsNumbered="' + partsText + '"' +
          ' notesAll="' + notesText + '"'
        );
        log('Corebridge line/qty debug: selectedRow=' + JSON.stringify(selectedRowDebug) + ' quoteLineItemSelection=' + JSON.stringify(quoteLineItemDebug));
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
    function refreshCutfilePathLabels() {
      runButtonJsxOperation('signarama_helper_cutfile_refreshFilePathLabels()', {logFn: log, toastTitle: 'Refresh file path label'});
    }
    function makeCutfile(fnCall, title) {
      runButtonJsxOperation(fnCall, {logFn: log, toastTitle: title, onResult: function(res) {
        const text = String(res || '');
        if(!/^Error:/i.test(text) && !/^No\b/i.test(text)) refreshCutfilePathLabels();
      }});
    }
    const makeRouterCutfile = $('btnMakeRouterCutfile');
    if(makeRouterCutfile) makeRouterCutfile.onclick = () => makeCutfile('signarama_helper_makeRouterCutfile()', 'Make Router Cutfile');

    const makeLaserCutfile = $('btnMakeLaserCutfile');
    if(makeLaserCutfile) makeLaserCutfile.onclick = () => makeCutfile('signarama_helper_makeLaserCutfile()', 'Make Laser Cutfile');

    const refreshCutfilePath = $('btnRefreshCutfilePathLabel');
    if(refreshCutfilePath) refreshCutfilePath.onclick = refreshCutfilePathLabels;

    const outlineAll = $('btnOutlineAllText');
    if(outlineAll) outlineAll.onclick = () => runButtonJsxOperation('signarama_helper_outlineAllText()', {logFn: log, toastTitle: 'Outline all text'});

    const setFillsStrokes = $('btnSetFillsStrokes');
    if(setFillsStrokes) setFillsStrokes.onclick = () => runButtonJsxOperation('signarama_helper_setAllFillsStrokes()', {logFn: log, toastTitle: 'Set fills/strokes'});

    wireDimensions();
    wireTransform();
    wireFixings();
    wirePreflight();
    wireLightbox();
    wireLedLayout();
    wireLedDepiction();
    wireLedLetterSpecs();
    wireLedCentreline();
    wireColours();
    wireNest();
    wireScripts();

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

  function wireScripts() {
    const list = $('predefinedScriptsList');
    const search = $('predefinedScriptsSearch');
    const refresh = $('btnRefreshScripts');
    const runFile = $('btnRunScriptFile');
    const runCode = $('btnRunScriptCode');
    const code = $('scriptCode');
    const clear = $('btnClearScriptLog');
    const output = $('scriptsLog');

    function scriptLog(message) {
      if(!output) return;
      output.textContent += String(message == null ? '' : message) + '\n';
      output.scrollTop = output.scrollHeight;
    }
    function reportResult(result, title) {
      const text = String(result == null ? '' : result);
      scriptLog(text || 'Completed (no return value).');
      notifyOperationResult(text || 'Completed (no return value).', {toastTitle: title || 'Script'});
    }
    function runScriptPath(path, title) {
      scriptLog('Running: ' + path);
      callJSX('signarama_helper_runScriptFile("' + jsxEscapeDoubleQuoted(path) + '")', function(result) {
        reportResult(result, title || 'Run script');
      });
    }
    function filterScripts() {
      if(!list) return;
      const query = String((search && search.value) || '').trim().toLowerCase();
      const rows = list.querySelectorAll('.script-list-row');
      let visible = 0;
      Array.prototype.forEach.call(rows, function(row) {
        const matches = !query || String(row.getAttribute('data-script-name') || '').toLowerCase().indexOf(query) !== -1;
        row.style.display = matches ? '' : 'none';
        if(matches) visible += 1;
      });
      let empty = list.querySelector('.script-filter-empty');
      if(rows.length && !visible) {
        if(!empty) {
          empty = document.createElement('div');
          empty.className = 'small script-filter-empty';
          empty.textContent = 'No scripts match your search.';
          list.appendChild(empty);
        }
        empty.style.display = '';
      } else if(empty) {
        empty.style.display = 'none';
      }
    }
    function refreshList() {
      if(!list) return;
      list.innerHTML = '<div class="small">Loading scripts...</div>';
      const extensionPath = cs.getSystemPath(SystemPath.EXTENSION).replace(/\\/g, '/');
      const scriptsPath = extensionPath + '/jsx/scripts';
      scriptLog('Scanning predefined scripts: ' + scriptsPath);
      callJSX('signarama_helper_listPredefinedScripts("' + jsxEscapeDoubleQuoted(scriptsPath) + '")', function(result) {
        let scripts = [];
        let diagnostics = null;
        try {
          const parsed = JSON.parse(String(result || '{}'));
          scripts = Array.isArray(parsed) ? parsed : (Array.isArray(parsed.scripts) ? parsed.scripts : []);
          diagnostics = parsed && parsed.diagnostics ? parsed.diagnostics : null;
        } catch(e) {
          scriptLog('Script scan returned invalid JSON: ' + String(result || '(empty)'));
          scriptLog('Script scan parse error: ' + (e && e.message ? e.message : e));
        }
        if(diagnostics) {
          scriptLog('Resolved scripts folder: ' + diagnostics.folder + ' (exists: ' + diagnostics.exists + ')');
          scriptLog('Folder entries: ' + (diagnostics.entries && diagnostics.entries.length ? diagnostics.entries.join(', ') : '(none)'));
          if(diagnostics.error) scriptLog('Script scan error: ' + diagnostics.error);
        }
        list.innerHTML = '';
        if(!scripts.length) {
          list.innerHTML = '<div class="small">No .jsx or .js scripts found in jsx/scripts.</div>';
          return;
        }
        scripts.forEach(function(script) {
          const row = document.createElement('div');
          row.className = 'script-list-row';
          row.setAttribute('data-script-name', script.name || '');
          const name = document.createElement('div');
          name.className = 'script-list-name';
          name.textContent = script.name;
          name.title = script.path;
          const button = document.createElement('button');
          button.type = 'button';
          button.className = 'btn2';
          button.textContent = 'Run';
          button.addEventListener('click', function() {runScriptPath(script.path, script.name);});
          row.appendChild(name);
          row.appendChild(button);
          list.appendChild(row);
        });
        filterScripts();
      });
    }

    scheduleScriptListRefresh = function() {setTimeout(refreshList, 0);};
    if(refresh) refresh.onclick = refreshList;
    if(search) search.addEventListener('input', filterScripts);
    if(runFile) runFile.onclick = function() {
      callJSX('signarama_helper_chooseAndRunScriptFile()', function(result) {reportResult(result, 'Run script file');});
    };
    if(runCode) runCode.onclick = function() {
      const source = String((code && code.value) || '');
      if(!source.trim()) {
        showToast('Paste some script code first.', {type: 'warn', title: 'Scripts'});
        return;
      }
      scriptLog('Running pasted code...');
      callJSX('signarama_helper_runScriptCode("' + jsxEscapeDoubleQuoted(source) + '")', function(result) {
        reportResult(result, 'Run pasted code');
      });
    };
    if(clear) clear.onclick = function() {if(output) output.textContent = '';};
  }

  function wireNest() {
    const partsBtn = $('btnNestLoadParts');
    const placeBtn = $('btnNestPlaceResult');
    const startBtn = $('btnNestStart');
    const stopBtn = $('btnNestStop');
    const clearBtn = $('btnNestClear');
    const sizeModeField = $('nestSizeMode');
    const manualSizeWrap = $('nestManualSizeFields');
    const selectionSizeWrap = $('nestSelectionSizeFields');
    const selectionShapeWrap = $('nestSelectionShapeFields');
    const partsSummary = $('nestPartsSummary');
    const selectionSummary = $('nestSelectionSizeSummary');
    const selectionShapeSummary = $('nestSelectionShapeSummary');
    const statusText = $('nestStatusText');
    const progressBar = $('nestProgressBar');
    const logEl = $('nestLog');
    const iterationsEl = $('nestIterations');
    const efficiencyEl = $('nestEfficiency');
    const placedEl = $('nestPlaced');
    const sheetSizeEl = $('nestSheetSize');

    if(!partsBtn || !startBtn) return;

    const nestState = {
      partsSvg: '',
      partsMeta: null,
      selectionSize: null,
      selectionBin: null,
      working: false,
      iterations: 0,
      latestPlacement: null,
      latestResultSvgList: null,
      latestOutputSvg: '',
      latestOutputSvgPath: '',
      latestBinSize: null,
      latestSheetMode: 'manual',
      startTimeMs: 0,
      heartbeatTimer: null,
      latestProgress: 0,
      lastHeartbeatLogMs: 0
    };

    let lastSavedProgressPct = -1;

    function mmToPt(mm) {
      return num(mm) * 2.834645669291339;
    }

    function logNest(line) {
      log('[Nest] ' + String(line));
      if(logEl) {
        logEl.textContent += String(line) + '\n';
        logEl.scrollTop = logEl.scrollHeight;
      }
    }

    function saveNestBackup(payload) {
      try {
        if(typeof localStorage === 'undefined') return false;
        localStorage.setItem(SRH_NEST_BACKUP_KEY, JSON.stringify(payload || {}));
        localStorage.setItem(SRH_NEST_FORCE_TAB_KEY, '1');
        return true;
      } catch(_eNestBackupSave) {
        logNest('Backup save failed: ' + (_eNestBackupSave && _eNestBackupSave.message ? _eNestBackupSave.message : _eNestBackupSave));
        return false;
      }
    }

    function writeNestOutputToTempFile(svgText) {
      try {
        if(typeof require !== 'function') return '';
        const fs = require('fs');
        const os = require('os');
        const path = require('path');
        const outDir = path.join(os.tmpdir(), 'SignaramaIllustratorHelper');
        if(!fs.existsSync(outDir)) fs.mkdirSync(outDir, {recursive: true});
        const filePath = path.join(outDir, 'nest-result-' + Date.now() + '.svg');
        fs.writeFileSync(filePath, String(svgText || ''), 'utf8');
        return filePath;
      } catch(_eNestTempWrite) {
        logNest('Result file save failed: ' + (_eNestTempWrite && _eNestTempWrite.message ? _eNestTempWrite.message : _eNestTempWrite));
        return '';
      }
    }

    function readNestOutputFromTempFile(filePath) {
      try {
        if(!filePath || typeof require !== 'function') return '';
        const fs = require('fs');
        if(!fs.existsSync(filePath)) return '';
        return String(fs.readFileSync(filePath, 'utf8') || '');
      } catch(_eNestTempRead) {
        logNest('Result file read failed: ' + (_eNestTempRead && _eNestTempRead.message ? _eNestTempRead.message : _eNestTempRead));
        return '';
      }
    }

    function clearNestBackup() {
      try {
        if(typeof localStorage === 'undefined') return;
        localStorage.removeItem(SRH_NEST_BACKUP_KEY);
        localStorage.removeItem(SRH_NEST_FORCE_TAB_KEY);
      } catch(_eNestBackupClear) { }
    }

    function copyTextToClipboard(text) {
      const value = String(text || '');
      if(!value) return Promise.reject(new Error('Nothing to copy.'));
      if(typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
        return navigator.clipboard.writeText(value);
      }
      return new Promise((resolve, reject) => {
        try {
          const ta = document.createElement('textarea');
          ta.value = value;
          ta.setAttribute('readonly', 'readonly');
          ta.style.position = 'fixed';
          ta.style.opacity = '0';
          ta.style.pointerEvents = 'none';
          document.body.appendChild(ta);
          ta.focus();
          ta.select();
          const ok = document.execCommand('copy');
          ta.remove();
          if(ok) resolve();
          else reject(new Error('Clipboard copy was rejected.'));
        } catch(err) {
          reject(err);
        }
      });
    }

    function buildNestClipboardPayload(details) {
      const info = details || {};
      const lines = [
        'SVGnest completed',
        'Iterations: ' + String(info.iterations || 0),
        'Efficiency: ' + String(info.efficiencyText || 'n/a'),
        'Placed: ' + String(info.placedText || 'n/a'),
        'Bins: ' + String(info.binCount || 0)
      ];
      if(info.sheetText) lines.push('Sheet: ' + String(info.sheetText));
      if(info.outputFilePath) lines.push('Saved file: ' + String(info.outputFilePath));
      if(info.savedAtText) lines.push('Saved at: ' + String(info.savedAtText));
      return lines.join('\n');
    }

    function restoreNestBackup() {
      try {
        if(typeof localStorage === 'undefined') return;
        const raw = localStorage.getItem(SRH_NEST_BACKUP_KEY);
        if(!raw) return;
        const saved = JSON.parse(raw);
        if(!saved) return;
        nestState.latestPlacement = saved.rawPlacement || null;
        nestState.latestOutputSvg = String(saved.outputSvg || '');
        nestState.latestOutputSvgPath = String(saved.outputFilePath || '');
        if(!nestState.latestOutputSvg && nestState.latestOutputSvgPath) {
          nestState.latestOutputSvg = readNestOutputFromTempFile(nestState.latestOutputSvgPath);
        }
        if(!nestState.latestOutputSvg && !nestState.latestOutputSvgPath && !nestState.latestPlacement) return;
        nestState.latestResultSvgList = [];
        setMetricText(iterationsEl, saved.iterations || 0);
        setMetricText(efficiencyEl, saved.efficiencyText || 'n/a');
        setMetricText(placedEl, saved.placedText || 'n/a');
        if(saved.sheetText) setMetricText(sheetSizeEl, saved.sheetText);
        setNestStatus('Recovered the last nest result after a panel reload.');
        logNest('Recovered a saved nest result from local storage.');
        updateButtons();
      } catch(_eNestRestore) {
        logNest('Nest backup restore failed: ' + (_eNestRestore && _eNestRestore.message ? _eNestRestore.message : _eNestRestore));
      }
    }

    if(typeof window !== 'undefined') {
      window.__srhNestLogHook = function(message) {
        if(!message) return;
        logNest(String(message));
      };
      if(!window.__srhNestConsoleMirrorInstalled && typeof console !== 'undefined' && console && typeof console.error === 'function') {
        const originalConsoleError = console.error.bind(console);
        console.error = function() {
          const args = Array.prototype.slice.call(arguments);
          try {
            originalConsoleError.apply(console, args);
          } catch(_eNestConsoleOriginal) { }
          try {
            if(window.__srhNestLogHook) {
              const text = args.map((entry) => {
                if(entry == null) return '';
                if(typeof entry === 'string') return entry;
                try {return JSON.stringify(entry);} catch(_eNestConsoleJson) {return String(entry);}
              }).join(' ');
              if(text) window.__srhNestLogHook('Console error: ' + text);
            }
          } catch(_eNestConsoleMirror) { }
        };
        window.__srhNestConsoleMirrorInstalled = true;
      }
    }

    function setNestStatus(text) {
      if(statusText) statusText.textContent = String(text || '');
    }

    function setNestProgress(percent) {
      if(!progressBar) return;
      const clamped = Math.max(0, Math.min(1, Number(percent) || 0));
      progressBar.style.width = (clamped * 100).toFixed(1) + '%';
      nestState.latestProgress = clamped;
    }

    function parseNestJson(raw) {
      const text = String(raw || '');
      try {
        return JSON.parse(text);
      } catch(_eNestParse) {
        try {
          return Function('return ' + text)();
        } catch(_eNestLegacyParse) {
          return {
            ok: false,
            error: text || 'Invalid response from Illustrator.'
          };
        }
      }
    }

    function setMetricText(el, text) {
      if(el) el.textContent = String(text);
    }

    function updateButtons() {
      if(stopBtn) stopBtn.disabled = !nestState.working;
      if(startBtn) startBtn.disabled = !nestState.partsSvg || nestState.working;
      if(placeBtn) placeBtn.disabled = (!nestState.latestOutputSvg && !nestState.latestOutputSvgPath && !nestState.latestPlacement && (!nestState.latestResultSvgList || !nestState.latestResultSvgList.length)) || nestState.working;
    }

    function clearNestHeartbeat() {
      if(nestState.heartbeatTimer) {
        clearInterval(nestState.heartbeatTimer);
        nestState.heartbeatTimer = null;
      }
    }

    function formatNestElapsed(ms) {
      const totalSeconds = Math.max(0, Math.floor((Number(ms) || 0) / 1000));
      const minutes = Math.floor(totalSeconds / 60);
      const seconds = totalSeconds % 60;
      return minutes + 'm ' + String(seconds).padStart(2, '0') + 's';
    }

    function updateNestHeartbeat() {
      if(!nestState.working) return;
      const now = Date.now();
      const elapsed = nestState.startTimeMs ? (now - nestState.startTimeMs) : 0;
      const progressPct = Math.round((nestState.latestProgress || 0) * 100);
      const resultState = nestState.iterations
        ? ('best result after ' + nestState.iterations + ' iteration(s)')
        : 'no result yet';
      setNestStatus('Nesting... ' + formatNestElapsed(elapsed) + ' elapsed, pass ' + progressPct + '%, ' + resultState + '.');
      if(!nestState.lastHeartbeatLogMs || (now - nestState.lastHeartbeatLogMs) >= 5000) {
        logNest('Still nesting... ' + formatNestElapsed(elapsed) + ' elapsed, pass ' + progressPct + '%, ' + resultState + '.');
        nestState.lastHeartbeatLogMs = now;
      }
    }

    function startNestHeartbeat() {
      clearNestHeartbeat();
      nestState.startTimeMs = Date.now();
      nestState.lastHeartbeatLogMs = 0;
      nestState.latestProgress = 0;
      nestState.heartbeatTimer = setInterval(updateNestHeartbeat, 1000);
      updateNestHeartbeat();
    }

    function updatePartsSummary() {
      if(!partsSummary) return;
      if(!nestState.partsMeta) {
        partsSummary.textContent = 'No parts loaded.';
        return;
      }
      const widthMm = Number(nestState.partsMeta.boundsMm && nestState.partsMeta.boundsMm.width) || 0;
      const heightMm = Number(nestState.partsMeta.boundsMm && nestState.partsMeta.boundsMm.height) || 0;
      const itemCount = parseInt(nestState.partsMeta.itemCount, 10) || 0;
      partsSummary.textContent = itemCount + ' part(s) loaded. Source bounds: ' + widthMm.toFixed(2) + ' x ' + heightMm.toFixed(2) + ' mm.';
    }

    function updateSelectionSummary() {
      if(!selectionSummary) return;
      if(!nestState.selectionSize) {
        selectionSummary.textContent = 'No selection size captured.';
        return;
      }
      selectionSummary.textContent = 'Captured selection bounds: ' + nestState.selectionSize.widthMm.toFixed(2) + ' x ' + nestState.selectionSize.heightMm.toFixed(2) + ' mm (' + nestState.selectionSize.itemCount + ' item(s)).';
    }

    function updateSelectionShapeSummary() {
      if(!selectionShapeSummary) return;
      if(!nestState.selectionBin) {
        selectionShapeSummary.textContent = 'No selection shape captured.';
        return;
      }
      selectionShapeSummary.textContent = 'Captured bin shape: ' + nestState.selectionBin.widthMm.toFixed(2) + ' x ' + nestState.selectionBin.heightMm.toFixed(2) + ' mm (' + nestState.selectionBin.itemCount + ' item selected).';
    }

    function updateSheetMetric() {
      const size = getBinSize();
      if(!size) {
        setMetricText(sheetSizeEl, '-');
        return;
      }
      setMetricText(sheetSizeEl, size.widthMm.toFixed(0) + ' x ' + size.heightMm.toFixed(0) + ' mm');
    }

    function updateSizeModeUi() {
      const mode = String((sizeModeField && sizeModeField.value) || 'manual');
      if(manualSizeWrap) manualSizeWrap.classList.toggle('hidden', mode !== 'manual');
      if(selectionSizeWrap) selectionSizeWrap.classList.toggle('hidden', mode !== 'selection');
      if(selectionShapeWrap) selectionShapeWrap.classList.toggle('hidden', mode !== 'shape');
      updateSheetMetric();
    }

    function getBinSize() {
      const mode = String((sizeModeField && sizeModeField.value) || 'manual');
      if(mode === 'shape') {
        if(!nestState.selectionBin) return null;
        return {
          widthPt: Number(nestState.selectionBin.widthPt) || 0,
          heightPt: Number(nestState.selectionBin.heightPt) || 0,
          widthMm: Number(nestState.selectionBin.widthMm) || 0,
          heightMm: Number(nestState.selectionBin.heightMm) || 0
        };
      }
      if(mode === 'selection') {
        if(!nestState.selectionSize) return null;
        return {
          widthPt: Number(nestState.selectionSize.widthPt) || 0,
          heightPt: Number(nestState.selectionSize.heightPt) || 0,
          widthMm: Number(nestState.selectionSize.widthMm) || 0,
          heightMm: Number(nestState.selectionSize.heightMm) || 0
        };
      }
      const widthMm = num(($('nestBinWidthMm') && $('nestBinWidthMm').value) || 0);
      const heightMm = num(($('nestBinHeightMm') && $('nestBinHeightMm').value) || 0);
      if(!(widthMm > 0) || !(heightMm > 0)) return null;
      return {
        widthPt: mmToPt(widthMm),
        heightPt: mmToPt(heightMm),
        widthMm: widthMm,
        heightMm: heightMm
      };
    }

    function stopNest() {
      if(window.SvgNest && typeof window.SvgNest.stop === 'function') {
        try {window.SvgNest.stop();} catch(_eNestStop) { }
      }
      nestState.working = false;
      saveNestRuntimeMarker('nest-stop');
      clearNestHeartbeat();
      updateButtons();
    }

    function measureSvgBounds(svgNode, targetNode) {
      if(!svgNode || !targetNode || !document.body) return null;
      const host = document.createElement('div');
      host.style.position = 'fixed';
      host.style.left = '-100000px';
      host.style.top = '-100000px';
      host.style.width = '1px';
      host.style.height = '1px';
      host.style.opacity = '0';
      host.style.pointerEvents = 'none';
      try {
        document.body.appendChild(host);
        host.appendChild(svgNode);
        const box = targetNode.getBBox ? targetNode.getBBox() : null;
        if(!box) return null;
        return {
          x: Number(box.x) || 0,
          y: Number(box.y) || 0,
          width: Number(box.width) || 0,
          height: Number(box.height) || 0
        };
      } catch(_eNestMeasure) {
        return null;
      } finally {
        if(host.parentNode) host.parentNode.removeChild(host);
      }
    }

    function buildNestSource() {
      if(!nestState.partsSvg) return null;
      const mode = String((sizeModeField && sizeModeField.value) || 'manual');
      const binSize = getBinSize();
      if(!binSize || !(binSize.widthPt > 0) || !(binSize.heightPt > 0)) return null;

      const parser = new DOMParser();
      const parsed = parser.parseFromString(nestState.partsSvg, 'image/svg+xml');
      const partSvg = parsed.documentElement;
      if(!partSvg || !partSvg.childNodes) return null;

      const ns = 'http://www.w3.org/2000/svg';
      const out = document.createElementNS(ns, 'svg');
      out.setAttribute('xmlns', ns);
      out.setAttribute('version', '1.1');
      const collectedStyleTexts = [];

      function collectStyleText(rootNode) {
        if(!rootNode || !rootNode.querySelectorAll) return;
        Array.prototype.slice.call(rootNode.querySelectorAll('style')).forEach((styleNode) => {
          const text = String((styleNode && styleNode.textContent) || '').trim();
          if(text) collectedStyleTexts.push(text);
        });
      }

      const partsScaleGroup = document.createElementNS(ns, 'g');
      const partsOriginGroup = document.createElementNS(ns, 'g');
      partsScaleGroup.appendChild(partsOriginGroup);
      collectStyleText(partSvg);

      const sourceWidthPt = Math.max(
        binSize.widthPt + mmToPt(10),
        Number(nestState.partsMeta && nestState.partsMeta.boundsPt && nestState.partsMeta.boundsPt.width) || 0
      );
      const sourceHeightPt = Math.max(
        binSize.heightPt + mmToPt(10),
        Number(nestState.partsMeta && nestState.partsMeta.boundsPt && nestState.partsMeta.boundsPt.height) || 0
      );

      out.setAttribute('viewBox', '0 0 ' + sourceWidthPt + ' ' + sourceHeightPt);
      out.setAttribute('width', String(sourceWidthPt));
      out.setAttribute('height', String(sourceHeightPt));

      function stampPartIdentity(node, sourcePartId) {
        if(!node || !sourcePartId) return;
        if(node.setAttribute) {
          node.setAttribute('data-srh-part-id', sourcePartId);
          const existingClass = String(node.getAttribute('class') || '').trim();
          if(existingClass.split(/\s+/).indexOf(sourcePartId) < 0) {
            node.setAttribute('class', (existingClass ? (existingClass + ' ') : '') + sourcePartId);
          }
        }
        if(node.childNodes && node.childNodes.length) {
          Array.prototype.slice.call(node.childNodes).forEach((childNode) => {
            if(childNode && childNode.tagName) stampPartIdentity(childNode, sourcePartId);
          });
        }
      }

      let rawPartIndex = 0;
      const topLevelPartNodes = [];
      Array.prototype.slice.call(partSvg.childNodes).forEach((child) => {
        if(!child || !child.tagName) return;
        const cloned = child.cloneNode(true);
        if(cloned && cloned.tagName && cloned.setAttribute) {
          const sourcePartId = 'srh-nest-part-' + rawPartIndex;
          cloned.setAttribute('data-srh-part-id', sourcePartId);
          if(!cloned.getAttribute('id')) cloned.setAttribute('id', sourcePartId);
          const existingClass = String(cloned.getAttribute('class') || '').trim();
          cloned.setAttribute('class', (existingClass ? (existingClass + ' ') : '') + sourcePartId);
          stampPartIdentity(cloned, sourcePartId);
          topLevelPartNodes.push(cloned);
          rawPartIndex += 1;
        }
        partsOriginGroup.appendChild(cloned);
      });
      out.appendChild(partsScaleGroup);

      const measuredBounds = measureSvgBounds(out, partsOriginGroup);
      const expectedWidthPt = Number(nestState.partsMeta && nestState.partsMeta.boundsPt && nestState.partsMeta.boundsPt.width) || 0;
      const expectedHeightPt = Number(nestState.partsMeta && nestState.partsMeta.boundsPt && nestState.partsMeta.boundsPt.height) || 0;
      let normalizationNote = 'none';
      if(measuredBounds && measuredBounds.width > 0 && measuredBounds.height > 0) {
        let scaleRatio = 1;
        const widthRatio = expectedWidthPt > 0 ? (expectedWidthPt / measuredBounds.width) : 1;
        const heightRatio = expectedHeightPt > 0 ? (expectedHeightPt / measuredBounds.height) : 1;
        const widthRatioOk = isFinite(widthRatio) && widthRatio > 0;
        const heightRatioOk = isFinite(heightRatio) && heightRatio > 0;
        const ratioSpread = (widthRatioOk && heightRatioOk)
          ? (Math.abs(widthRatio - heightRatio) / Math.max(widthRatio, heightRatio))
          : 0;
        if(widthRatioOk && heightRatioOk && ratioSpread <= 0.08) {
          scaleRatio = (widthRatio + heightRatio) / 2;
          normalizationNote = 'uniform-scale widthRatio=' + widthRatio.toFixed(4) + ', heightRatio=' + heightRatio.toFixed(4);
        } else if(widthRatioOk && !heightRatioOk) {
          scaleRatio = widthRatio;
          normalizationNote = 'width-only-scale widthRatio=' + widthRatio.toFixed(4);
        } else if(heightRatioOk && !widthRatioOk) {
          scaleRatio = heightRatio;
          normalizationNote = 'height-only-scale heightRatio=' + heightRatio.toFixed(4);
        } else if(widthRatioOk && heightRatioOk) {
          scaleRatio = 1;
          normalizationNote = 'scale-skipped ratio-mismatch widthRatio=' + widthRatio.toFixed(4) + ', heightRatio=' + heightRatio.toFixed(4);
        }
        const normalizeTransform = (isFinite(scaleRatio) && scaleRatio > 0 && Math.abs(scaleRatio - 1) > 0.001)
          ? ('scale(' + scaleRatio + ') translate(' + (-measuredBounds.x) + ' ' + (-measuredBounds.y) + ')')
          : ('translate(' + (-measuredBounds.x) + ' ' + (-measuredBounds.y) + ')');
        if(partsScaleGroup.parentNode === out) out.removeChild(partsScaleGroup);
        topLevelPartNodes.forEach((partNode) => {
          const wrapper = document.createElementNS(ns, 'g');
          const sourcePartId = partNode.getAttribute('data-srh-part-id') || partNode.getAttribute('id') || '';
          if(sourcePartId) {
            wrapper.setAttribute('data-srh-part-id', sourcePartId);
            wrapper.setAttribute('id', sourcePartId);
            wrapper.setAttribute('class', sourcePartId);
            stampPartIdentity(partNode, sourcePartId);
          }
          wrapper.setAttribute('transform', normalizeTransform);
          wrapper.appendChild(partNode);
          out.appendChild(wrapper);
        });
        out.setAttribute('data-srh-part-scale', String(scaleRatio));
        out.setAttribute('data-srh-part-bounds', measuredBounds.width.toFixed(3) + 'x' + measuredBounds.height.toFixed(3));
        out.setAttribute('data-srh-part-normalization-note', normalizationNote);
      } else {
        if(partsScaleGroup.parentNode === out) out.removeChild(partsScaleGroup);
        topLevelPartNodes.forEach((partNode) => {
          const sourcePartId = partNode.getAttribute && (partNode.getAttribute('data-srh-part-id') || partNode.getAttribute('id') || '');
          if(sourcePartId) stampPartIdentity(partNode, sourcePartId);
          out.appendChild(partNode);
        });
      }

      if(mode === 'shape') {
        if(!nestState.selectionBin || !nestState.selectionBin.svgText) return null;
        const parsedBin = parser.parseFromString(String(nestState.selectionBin.svgText || ''), 'image/svg+xml');
        const binSvg = parsedBin.documentElement;
        if(!binSvg) return null;
        collectStyleText(binSvg);
        const binElement = binSvg.querySelector ? binSvg.querySelector('path, polygon, polyline, rect, circle, ellipse') : null;
        if(!binElement) return null;
        const binGroup = document.createElementNS(ns, 'g');
        const binClone = binElement.cloneNode(true);
        binClone.setAttribute('id', 'nest-bin');
        const originalAppearance = nestState.selectionBin.originalAppearance || null;
        if(originalAppearance) {
          if(originalAppearance.filled) {
            binClone.setAttribute('fill', String(originalAppearance.fillCss || 'none'));
          } else {
            binClone.setAttribute('fill', 'none');
          }
          if(originalAppearance.stroked) {
            binClone.setAttribute('stroke', String(originalAppearance.strokeCss || 'none'));
            if(Number(originalAppearance.strokeWidthPt || 0) > 0) {
              binClone.setAttribute('stroke-width', String(Number(originalAppearance.strokeWidthPt || 0)));
            }
          } else {
            binClone.setAttribute('stroke', 'none');
          }
          if(isFinite(Number(originalAppearance.opacity))) {
            binClone.setAttribute('opacity', String(Number(originalAppearance.opacity)));
          }
        }
        binGroup.appendChild(binClone);
        out.appendChild(binGroup);
        const binBounds = measureSvgBounds(out, binGroup);
        if(!binBounds || !(binBounds.width > 0) || !(binBounds.height > 0)) return null;
        binGroup.setAttribute('transform', 'translate(' + (-binBounds.x) + ' ' + (-binBounds.y) + ')');
      } else {
        const binRect = document.createElementNS(ns, 'rect');
        binRect.setAttribute('id', 'nest-bin');
        binRect.setAttribute('x', '0');
        binRect.setAttribute('y', '0');
        binRect.setAttribute('width', String(binSize.widthPt));
        binRect.setAttribute('height', String(binSize.heightPt));
        binRect.setAttribute('fill', 'none');
        binRect.setAttribute('stroke', '#c8102e');
        binRect.setAttribute('stroke-width', String(Math.max(1, mmToPt(0.2))));
        out.appendChild(binRect);
      }

      if(collectedStyleTexts.length) {
        const mergedStyle = document.createElementNS(ns, 'style');
        mergedStyle.setAttribute('type', 'text/css');
        mergedStyle.textContent = collectedStyleTexts.join('\n');
        if(out.firstChild) out.insertBefore(mergedStyle, out.firstChild);
        else out.appendChild(mergedStyle);
      }

      const serialized = (new XMLSerializer()).serializeToString(out);
      const normalizedChildSummary = Array.prototype.slice.call(out.childNodes).map((child, index) => {
        if(!child || !child.tagName) return null;
        return '#' + index + ':' + child.tagName + (child.getAttribute && child.getAttribute('id') ? ('#' + child.getAttribute('id')) : '') + '[children=' + (child.childNodes ? child.childNodes.length : 0) + ']';
      }).filter(Boolean).join(', ');
      return {
        svgText: serialized,
        binSize: binSize,
        partScale: Number(out.getAttribute('data-srh-part-scale') || 1),
        partBoundsText: String(out.getAttribute('data-srh-part-bounds') || ''),
        partNormalizationNote: String(out.getAttribute('data-srh-part-normalization-note') || normalizationNote),
        normalizedChildSummary: normalizedChildSummary,
        rawChildSummary: summarizeSvgChildrenFromText(nestState.partsSvg)
      };
    }

    function serializeOutputSvgList(svgList) {
      if(!svgList || !svgList.length) return '';
      const ns = 'http://www.w3.org/2000/svg';
      const root = document.createElementNS(ns, 'svg');
      root.setAttribute('xmlns', ns);
      root.setAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
      root.setAttribute('version', '1.1');
      let offsetY = 0;
      let maxWidth = 0;

      if(window.SvgNest && window.SvgNest.style) {
        try {root.appendChild(window.SvgNest.style.cloneNode(true));} catch(_eNestStyle) { }
      }

      svgList.forEach((svgNode) => {
        const width = parseFloat(svgNode.getAttribute('width')) || ((svgNode.viewBox && svgNode.viewBox.baseVal && svgNode.viewBox.baseVal.width) || 0);
        const height = parseFloat(svgNode.getAttribute('height')) || ((svgNode.viewBox && svgNode.viewBox.baseVal && svgNode.viewBox.baseVal.height) || 0);
        const group = document.createElementNS(ns, 'g');
        group.setAttribute('transform', 'translate(0 ' + offsetY + ')');
        Array.prototype.slice.call(svgNode.childNodes).forEach((child) => {
          group.appendChild(child.cloneNode(true));
        });
        root.appendChild(group);
        maxWidth = Math.max(maxWidth, width);
        offsetY += height + (height * 0.1);
      });

      const finalHeight = Math.max(0, offsetY);
      root.setAttribute('viewBox', '0 0 ' + maxWidth + ' ' + finalHeight);
      root.setAttribute('width', String(maxWidth));
      root.setAttribute('height', String(finalHeight));
      return '<?xml version="1.0" encoding="UTF-8"?>\n' + (new XMLSerializer()).serializeToString(root);
    }

    function resetNestMetrics() {
      nestState.iterations = 0;
      nestState.latestProgress = 0;
      setMetricText(iterationsEl, '0');
      setMetricText(efficiencyEl, '0%');
      setMetricText(placedEl, '0/0');
      setNestProgress(0);
      updateSheetMetric();
    }

    function callNestJsx(functionName, argString, cb) {
      loadJSX(function() {
        const fallback = '{\\"ok\\":false,\\"error\\":\\"' + functionName + ' not loaded.\\"}';
        const fnCall = '((typeof ' + functionName + ' === "function") ? ' + functionName + ' : ((typeof $ !== "undefined" && $.global && typeof $.global.' + functionName + ' === "function") ? $.global.' + functionName + ' : function(){return "' + fallback + '";}))(' + (argString || '') + ')';
        callJSX(fnCall, cb);
      });
    }

    function startNest() {
      if(!window.SvgNest) {
        showToast('SVGnest failed to load in this panel.', {type: 'error', title: 'Nest'});
        return;
      }
      const source = buildNestSource();
      if(!source) {
        showToast('Load parts and provide a valid sheet size before nesting.', {type: 'warn', title: 'Nest'});
        return;
      }

      stopNest();
      nestState.latestPlacement = null;
      nestState.latestResultSvgList = null;
      nestState.latestOutputSvg = '';
      nestState.latestOutputSvgPath = '';
      nestState.latestBinSize = source.binSize || null;
      nestState.latestSheetMode = String((sizeModeField && sizeModeField.value) || 'manual');
      lastSavedProgressPct = -1;
      clearNestBackup();
      resetNestMetrics();
      updateButtons();

      const spacingPt = mmToPt(($('nestSpacingMm') && $('nestSpacingMm').value) || 0);
      const curveTolerancePt = mmToPt(($('nestCurveToleranceMm') && $('nestCurveToleranceMm').value) || 0.5);
      const config = {
        spacing: spacingPt,
        curveTolerance: curveTolerancePt,
        rotations: parseInt((($('nestRotations') && $('nestRotations').value) || 4), 10) || 4,
        populationSize: parseInt((($('nestPopulationSize') && $('nestPopulationSize').value) || 10), 10) || 10,
        mutationRate: parseInt((($('nestMutationRate') && $('nestMutationRate').value) || 10), 10) || 10,
        useHoles: !!($('nestUseHoles') && $('nestUseHoles').checked),
        exploreConcave: !!($('nestExploreConcave') && $('nestExploreConcave').checked)
      };

      try {
        if(typeof window !== 'undefined') {
          window.__srhSvgNestUsePlacementOnlyCallback = true;
          window.__srhSvgNestMaxWorkers = 1;
          window.__srhSvgNestDisableWorkers = true;
        }
        window.SvgNest.config(config);
        const parsedSvg = window.SvgNest.parsesvg(source.svgText);
        const binNode = parsedSvg ? parsedSvg.querySelector('#nest-bin') : null;
        if(!parsedSvg || !binNode) {
          showToast('Could not prepare the nesting SVG input.', {type: 'error', title: 'Nest'});
          return;
        }
        window.SvgNest.setbin(binNode);
      } catch(e) {
        showToast('SVGnest setup failed: ' + (e && e.message ? e.message : e), {type: 'error', title: 'Nest'});
        logNest('SVGnest setup failed: ' + (e && e.message ? e.message : e));
        return;
      }

      nestState.working = true;
      saveNestRuntimeMarker('nest-start', {
        partsLoaded: !!nestState.partsSvg
      });
      startNestHeartbeat();
      updateButtons();
      setNestStatus('Searching for a better nesting layout...');
      logNest(
        'Nest started. Sheet ' + source.binSize.widthMm.toFixed(2) + ' x ' + source.binSize.heightMm.toFixed(2) +
        ' mm. Spacing ' + (($('nestSpacingMm') && $('nestSpacingMm').value) || 0) +
        ' mm, rotations ' + config.rotations + ', population ' + config.populationSize + ', mutation ' + config.mutationRate + ', workers disabled=' + (!!(window && window.__srhSvgNestDisableWorkers)) + '.'
      );
      if(source.partBoundsText) {
        logNest('Imported SVG geometry measured at ' + source.partBoundsText + ' pt with normalization scale ' + source.partScale.toFixed(4) + '.');
        if(source.partNormalizationNote) {
          logNest('Normalization detail: ' + source.partNormalizationNote + '.');
        }
      }
      if(source.rawChildSummary) {
        logNest('Nest source raw SVG children: ' + source.rawChildSummary + '.');
      }
      if(source.normalizedChildSummary) {
        logNest('Nest source normalized children: ' + source.normalizedChildSummary + '.');
      }

      window.SvgNest.start(function(percent) {
        setNestProgress(percent);
        if(nestState.working) {
          const elapsed = nestState.startTimeMs ? (Date.now() - nestState.startTimeMs) : 0;
          const progressPct = Math.round((nestState.latestProgress || 0) * 100);
          if(progressPct !== lastSavedProgressPct && progressPct % 5 === 0) {
            lastSavedProgressPct = progressPct;
            saveNestRuntimeMarker('nest-progress-' + progressPct, {progressPct: progressPct});
          }
          const resultState = nestState.iterations
            ? ('best result after ' + nestState.iterations + ' iteration(s)')
            : 'still searching for first result';
          setNestStatus('Nesting... ' + formatNestElapsed(elapsed) + ' elapsed, pass ' + progressPct + '%, ' + resultState + '.');
        }
      }, function(svgList, efficiency, placed, total) {
        const placementPayload = (!Array.isArray(svgList) && svgList && svgList.placementOnly && svgList.placements)
          ? svgList
          : null;
        const hasSvgListResult = Array.isArray(svgList) && svgList.length > 0;
        const hasPlacementResult = !!(placementPayload && placementPayload.placements && placementPayload.placements.length);
        const resultBinCount = hasPlacementResult
          ? placementPayload.placements.length
          : (hasSvgListResult ? svgList.length : 0);

        saveNestRuntimeMarker('result-callback-enter', {
          binCount: resultBinCount
        });
        logNest('Result callback entered. Bins=' + resultBinCount + '.');

        if(!hasSvgListResult && !hasPlacementResult) {
          logNest('No improved placement in this pass.');
          return;
        }

        nestState.iterations += 1;
        setMetricText(iterationsEl, nestState.iterations);

        if(typeof efficiency === 'number' && isFinite(efficiency)) {
          setMetricText(efficiencyEl, Math.round(efficiency * 100) + '%');
        }
        if(typeof placed !== 'undefined' && typeof total !== 'undefined') {
          setMetricText(placedEl, String(placed) + '/' + String(total));
        }

        if(hasSvgListResult || hasPlacementResult) {
          try {
            const binCount = resultBinCount;
            nestState.latestPlacement = placementPayload ? placementPayload.placements : null;
            nestState.latestResultSvgList = placementPayload ? null : svgList;
            const efficiencyText = (typeof efficiency === 'number' && isFinite(efficiency)) ? (Math.round(efficiency * 100) + '%') : 'n/a';
            const placedText = (typeof placed !== 'undefined' && typeof total !== 'undefined') ? (String(placed) + '/' + String(total)) : 'n/a';
            const savedAtText = (new Date()).toLocaleString();
            saveNestBackup({
              rawPlacement: nestState.latestPlacement,
              iterations: nestState.iterations,
              efficiencyText: efficiencyText,
              placedText: placedText,
              sheetText: sheetSizeEl ? sheetSizeEl.textContent : '',
              binCount: binCount,
              savedAt: Date.now(),
              savedAtText: savedAtText
            });
            copyTextToClipboard(buildNestClipboardPayload({
              iterations: nestState.iterations,
              efficiencyText: efficiencyText,
              placedText: placedText,
              binCount: binCount,
              sheetText: sheetSizeEl ? sheetSizeEl.textContent : '',
              savedAtText: savedAtText
            }))
              .then(() => {
                saveNestRuntimeMarker('result-summary-copied');
                showToast('Nest finished. Result summary copied to clipboard.', {type: 'success', title: 'Nest', duration: 20000});
                logNest('Copied nest result summary to clipboard.');
              })
              .catch((copyErr) => {
                logNest('Clipboard summary copy failed: ' + (copyErr && copyErr.message ? copyErr.message : copyErr));
              });
            setNestStatus('Best result updated after ' + nestState.iterations + ' iteration(s).');
            logNest('Result #' + nestState.iterations + ': efficiency ' + efficiencyText + ', placed ' + placedText + ', bins ' + binCount + '.');
            if(placementPayload) {
              logNest('Stored raw placement data. Final SVG generation is deferred until you place or export it.');
              updateButtons();
              return;
            }
            updateButtons();
            setTimeout(() => {
              try {
                saveNestRuntimeMarker('result-svg-save-begin');
                const outputSvg = serializeOutputSvgList(svgList);
                const outputFilePath = writeNestOutputToTempFile(outputSvg);
                nestState.latestOutputSvg = outputSvg;
                nestState.latestOutputSvgPath = outputFilePath;
                saveNestBackup({
                  outputSvg: '',
                  outputFilePath: outputFilePath,
                  iterations: nestState.iterations,
                  efficiencyText: efficiencyText,
                  placedText: placedText,
                  sheetText: sheetSizeEl ? sheetSizeEl.textContent : '',
                  binCount: svgList.length,
                  savedAt: Date.now(),
                  savedAtText: savedAtText
                });
                if(outputFilePath) {
                  saveNestRuntimeMarker('result-svg-saved', {outputFilePath: outputFilePath});
                  logNest('Saved final nest SVG to ' + outputFilePath + '.');
                  copyTextToClipboard(buildNestClipboardPayload({
                    iterations: nestState.iterations,
                    efficiencyText: efficiencyText,
                    placedText: placedText,
                    binCount: svgList.length,
                    sheetText: sheetSizeEl ? sheetSizeEl.textContent : '',
                    outputFilePath: outputFilePath,
                    savedAtText: savedAtText
                  }))
                    .then(() => {
                      saveNestRuntimeMarker('result-filepath-copied', {outputFilePath: outputFilePath});
                      logNest('Copied nest result file path to clipboard.');
                    })
                    .catch((copyErr) => {
                      logNest('Clipboard file-path copy failed: ' + (copyErr && copyErr.message ? copyErr.message : copyErr));
                    });
                }
                updateButtons();
              } catch(renderErr) {
                logNest('Deferred result save error: ' + (renderErr && renderErr.message ? renderErr.message : renderErr));
              }
            }, 0);
          } catch(renderErr) {
            logNest('Result handling error: ' + (renderErr && renderErr.message ? renderErr.message : renderErr));
            setNestStatus('Result found, but preview rendering failed. You can still try placing it.');
            updateButtons();
          }
        }
      });
    }

    partsBtn.addEventListener('click', () => {
      const loadingToast = showToast('Reading selected Illustrator artwork...', {type: 'info', title: 'Nest', spinner: true, persistent: true});
      stopNest();
      callNestJsx('signarama_helper_nest_captureSelectionAsSvg', '', (res) => {
        if(loadingToast) loadingToast.close();
        const payload = parseNestJson(res);
        if(!payload.ok) {
          showToast(payload.error || 'Failed to load selection as SVG.', {type: 'error', title: 'Nest'});
          logNest('Load parts failed: ' + (payload.error || res));
          if(payload.debug && payload.debug.length) {
            payload.debug.forEach((line) => logNest('Debug: ' + line));
          }
          return;
        }
        nestState.partsSvg = String(payload.svgText || '');
        nestState.partsMeta = payload;
        nestState.latestPlacement = null;
        nestState.latestResultSvgList = null;
        nestState.latestOutputSvg = '';
        nestState.latestOutputSvgPath = '';
        nestState.latestBinSize = null;
        nestState.latestSheetMode = 'manual';
        clearNestBackup();
        updatePartsSummary();
        resetNestMetrics();
        updateButtons();
        setNestStatus('Parts loaded. Ready to nest.');
        logNest(
          'Loaded ' + (payload.itemCount || 0) + ' part(s) from Illustrator selection. Bounds ' +
          Number(payload.boundsMm && payload.boundsMm.width || 0).toFixed(2) + ' x ' +
          Number(payload.boundsMm && payload.boundsMm.height || 0).toFixed(2) + ' mm' +
          ' (scaleFactor=' + Number(payload.scaleFactor || 1) + ').'
        );
        logNest('Loaded-parts raw SVG children: ' + summarizeSvgChildrenFromText(nestState.partsSvg) + '.');
        if(payload.debug && payload.debug.length) {
          payload.debug.forEach((line) => logNest('Debug: ' + line));
        }
      });
    });

    function captureSelectionSize(options) {
      const opts = options || {};
      const loadingToast = opts.silent ? null : showToast('Reading selection bounds from Illustrator...', {type: 'info', title: 'Nest', spinner: true, persistent: true});
      callNestJsx('signarama_helper_nest_captureSelectionBounds', '', (res) => {
        if(loadingToast) loadingToast.close();
        const payload = parseNestJson(res);
        if(!payload.ok) {
          if(!opts.silent) showToast(payload.error || 'Failed to capture selection bounds.', {type: 'error', title: 'Nest'});
          logNest('Capture size failed: ' + (payload.error || res));
          if(opts.done) opts.done(false, payload);
          return;
        }
        nestState.selectionSize = payload;
        updateSelectionSummary();
        updateSheetMetric();
        setNestStatus('Selection size captured.');
        logNest('Captured sheet size from selection: ' + payload.widthMm.toFixed(2) + ' x ' + payload.heightMm.toFixed(2) + ' mm (scaleFactor=' + Number(payload.scaleFactor || 1) + ').');
        if(opts.done) opts.done(true, payload);
      });
    }

    function captureSelectionShape(options) {
      const opts = options || {};
      const loadingToast = opts.silent ? null : showToast('Reading selection shape from Illustrator...', {type: 'info', title: 'Nest', spinner: true, persistent: true});
      callNestJsx('signarama_helper_nest_captureSelectionShapeAsSvg', '', (res) => {
        if(loadingToast) loadingToast.close();
        const payload = parseNestJson(res);
        if(!payload.ok) {
          if(!opts.silent) showToast(payload.error || 'Failed to capture selection shape.', {type: 'error', title: 'Nest'});
          logNest('Capture shape failed: ' + (payload.error || res));
          if(payload.debug && payload.debug.length) {
            payload.debug.forEach((line) => logNest('Debug: ' + line));
          }
          if(opts.done) opts.done(false, payload);
          return;
        }
        nestState.selectionBin = payload;
        updateSelectionShapeSummary();
        updateSheetMetric();
        setNestStatus('Selection shape captured.');
        logNest('Captured bin shape from selection: ' + payload.widthMm.toFixed(2) + ' x ' + payload.heightMm.toFixed(2) + ' mm (scaleFactor=' + Number(payload.scaleFactor || 1) + ').');
        if(opts.done) opts.done(true, payload);
      });
    }

    function captureCurrentSheetSourceThenStart() {
      const mode = String((sizeModeField && sizeModeField.value) || 'manual');
      if(mode === 'selection') {
        const loadingToast = showToast('Reading current selection bounds from Illustrator...', {type: 'info', title: 'Nest', spinner: true, persistent: true});
        captureSelectionSize({
          silent: true,
          done: (ok, payload) => {
            if(loadingToast) loadingToast.close();
            if(!ok) {
              showToast((payload && payload.error) || 'Select artwork to use as the sheet bounds before nesting.', {type: 'error', title: 'Nest'});
              return;
            }
            startNest();
          }
        });
        return;
      }
      if(mode === 'shape') {
        const loadingToast = showToast('Reading current selection shape from Illustrator...', {type: 'info', title: 'Nest', spinner: true, persistent: true});
        captureSelectionShape({
          silent: true,
          done: (ok, payload) => {
            if(loadingToast) loadingToast.close();
            if(!ok) {
              showToast((payload && payload.error) || 'Select one shape to use as the sheet before nesting.', {type: 'error', title: 'Nest'});
              return;
            }
            startNest();
          }
        });
        return;
      }
      startNest();
    }

    startBtn.addEventListener('click', captureCurrentSheetSourceThenStart);
    if(stopBtn) {
      stopBtn.addEventListener('click', () => {
        stopNest();
        setNestStatus('Nesting stopped.');
        logNest('Nest stopped.');
      });
    }
    if(clearBtn) {
      clearBtn.addEventListener('click', () => {
        stopNest();
        nestState.latestPlacement = null;
        nestState.latestResultSvgList = null;
        nestState.latestOutputSvg = '';
        nestState.latestOutputSvgPath = '';
        nestState.latestBinSize = null;
        nestState.latestSheetMode = 'manual';
        clearNestBackup();
        resetNestMetrics();
        updateButtons();
        setNestStatus('Result cleared.');
      });
    }
    if(placeBtn) {
      placeBtn.addEventListener('click', () => {
        if(!nestState.latestOutputSvg && !nestState.latestOutputSvgPath && !nestState.latestPlacement && (!nestState.latestResultSvgList || !nestState.latestResultSvgList.length)) {
          showToast('Run a nest first so there is a result to place.', {type: 'warn', title: 'Nest'});
          return;
        }
        const loadingToast = showToast('Placing nested result into Illustrator...', {type: 'info', title: 'Nest', spinner: true, persistent: true});
        const placementItems = nestState.latestPlacement ? nestState.latestPlacement.reduce((sum, bin) => sum + (bin && bin.length ? bin.length : 0), 0) : 0;
        const snapshotIds = (nestState.partsMeta && nestState.partsMeta.snapshotIds) || [];
        const canPlaceNativeOriginals = !!(nestState.latestPlacement && nestState.latestPlacement.length && snapshotIds.length);
        const placementDebug = {
          placementBins: nestState.latestPlacement ? nestState.latestPlacement.length : 0,
          placementItems: placementItems,
          snapshotIds: snapshotIds,
          usingNativeOriginals: canPlaceNativeOriginals,
          usingPlacementSvg: !canPlaceNativeOriginals
        };
        logNest('Place debug: ' + JSON.stringify(placementDebug));
        function copyNestDebugLog() {
          const debugText = logEl ? String(logEl.textContent || '') : '';
          if(!debugText) return;
          copyTextToClipboard('NEST DEBUG LOG\n' + debugText).then(() => {
            logNest('Copied nest debug log to clipboard.');
          }).catch((copyErr) => {
            logNest('Debug log clipboard copy failed: ' + (copyErr && copyErr.message ? copyErr.message : copyErr));
          });
        }
        function finishPlaceResult(res, successFallback) {
          if(loadingToast) loadingToast.close();
          const text = String(res || '');
          const isError = /^error:/i.test(text);
          copyNestDebugLog();
          if(isError) {
            showToast(text, {type: 'error', title: 'Nest'});
            logNest('Place result failed: ' + text);
            return false;
          }
          showToast(text || successFallback, {type: 'success', title: 'Nest'});
          logNest(text || successFallback);
          return true;
        }
        if(canPlaceNativeOriginals) {
          const nativePayload = {
            placements: nestState.latestPlacement,
            snapshotLayerName: (nestState.partsMeta && nestState.partsMeta.snapshotLayerName) || 'SRH_NEST_SOURCE_SNAPSHOT',
            snapshotIds: snapshotIds,
            binSize: (nestState.latestBinSize || {}),
            sheetMode: nestState.latestSheetMode || 'manual',
            binSourceIds: ((nestState.latestSheetMode === 'shape' && nestState.selectionBin && nestState.selectionBin.binSourceIds)
              ? nestState.selectionBin.binSourceIds
              : ((nestState.latestSheetMode === 'selection' && nestState.selectionSize && nestState.selectionSize.binSourceIds) ? nestState.selectionSize.binSourceIds : []))
          };
          const safePayload = jsxEscapeDoubleQuoted(JSON.stringify(nativePayload));
          logNest('Placing original loaded artwork from raw placement data.');
          callNestJsx('signarama_helper_nest_placeNativeNestPlacement', '"' + safePayload + '"', (res) => {
            finishPlaceResult(res, 'Nested original artwork placed on the active artboard.');
          });
          return;
        }
        logNest('Native original placement metadata was not available; falling back to generated SVG placement.');
        let outputSvg = String(nestState.latestOutputSvg || '');
        try {
          if(!outputSvg && nestState.latestOutputSvgPath) {
            outputSvg = readNestOutputFromTempFile(nestState.latestOutputSvgPath);
            nestState.latestOutputSvg = outputSvg;
            if(outputSvg) logNest('Loaded the saved final SVG output from ' + nestState.latestOutputSvgPath + '.');
          }
          if(!outputSvg && nestState.latestPlacement && window.SvgNest && typeof window.SvgNest.applyPlacement === 'function') {
            logNest('Generating final SVG from saved raw placement data.');
            nestState.latestResultSvgList = window.SvgNest.applyPlacement(nestState.latestPlacement);
          }
          if(!outputSvg) {
            outputSvg = serializeOutputSvgList(nestState.latestResultSvgList);
            nestState.latestOutputSvg = outputSvg;
            logNest('Prepared final SVG output (' + outputSvg.length + ' chars) for Illustrator placement.');
          } else {
            logNest('Using the saved final SVG output (' + outputSvg.length + ' chars) for Illustrator placement.');
          }
        } catch(serializeErr) {
          if(loadingToast) loadingToast.close();
          showToast('Could not prepare the final SVG output: ' + (serializeErr && serializeErr.message ? serializeErr.message : serializeErr), {type: 'error', title: 'Nest'});
          logNest('Place result failed during serialization: ' + (serializeErr && serializeErr.message ? serializeErr.message : serializeErr));
          return;
        }
        const safeSvg = jsxEscapeDoubleQuoted(outputSvg);
        callNestJsx('signarama_helper_nest_placeSvgOnActiveArtboard', '"' + safeSvg + '"', (res) => {
          finishPlaceResult(res, 'Nested SVG placed on the active artboard.');
        });
      });
    }

    if(sizeModeField) sizeModeField.addEventListener('change', updateSizeModeUi);
    ['nestBinWidthMm', 'nestBinHeightMm', 'nestSpacingMm', 'nestRotations', 'nestCurveToleranceMm', 'nestPopulationSize', 'nestMutationRate'].forEach((id) => {
      const el = $(id);
      if(!el) return;
      el.addEventListener('input', () => {
        updateSheetMetric();
      });
      el.addEventListener('change', () => {
        updateSheetMetric();
      });
    });

    updatePartsSummary();
    updateSelectionSummary();
    updateSelectionShapeSummary();
    resetNestMetrics();
    updateButtons();
    updateSizeModeUi();
    restoreNestBackup();
    setNestStatus('Ready.');
  }

  function wireLightbox() {
    function refreshArtboardScaleNotice() {
      callJSX('app.documents.length ? app.activeDocument.scaleFactor : 1', sfRawDoc => {
        const sf = Number(sfRawDoc);
        const sfText = isFinite(sf) && sf > 0 ? sf : 1;
        const invScale = 1 / sfText;
        const large = sfText > 1.0001;
        const chipLabel = Math.abs(sfText - 10) < 0.0001 ? 'x10!' : ('x' + String(Math.round(sfText * 1000) / 1000));
        isLargeArtboard = large;
        if(typeof window !== 'undefined') window.isLargeArtboard = large;
        const chip = $('artboardScaleButton');
        const tooltip = $('artboardScaleTooltip');
        if(!chip || !tooltip) return;
        log('app.activeDocument.scaleFactor: ' + sfText);
        if(large) {
          chip.classList.add('large');
          chip.textContent = chipLabel;
          tooltip.textContent = 'Large artboard detected (scaleFactor=' + sfText + '). Dimension/lightbox geometry uses x' + invScale.toFixed(4) + ' (1/' + sfText + ').';
        } else {
          chip.classList.remove('large');
          chip.textContent = 'x1';
          tooltip.textContent = 'Standard artboard scale detected (scaleFactor=1). Dimension/lightbox geometry uses x1.0000.';
        }
        chip.setAttribute('aria-label', tooltip.textContent);
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
        expectedDocumentKey: force ? '' : lightboxMeasureLiveDocumentKey,
        includeDocumentIdentity: true,
        measureOptions: buildDimensionPayload(),
        isLargeArtboard: isLargeArtboard
      };
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      callJSX('signarama_helper_updateLightboxMeasures("' + json + '")', res => {
        lightboxMeasureLiveInFlight = false;
        const parts = String(res || '').split('|||');
        const documentKey = parts.length > 1 ? parts.shift() : '';
        const result = parts.length ? parts.join('|||') : String(res || '');
        if(documentKey) lightboxMeasureLiveDocumentKey = documentKey;
        if(result === 'Document changed.' || result === 'No live lightbox found.') stopLightboxMeasureLive();
        if(result && result !== 'No live measure changes.' && result !== 'No live lightbox found.') log(result);
      });
    }

    function stopLightboxMeasureLive() {
      if(lightboxMeasureLiveTimer) {
        clearInterval(lightboxMeasureLiveTimer);
        lightboxMeasureLiveTimer = null;
      }
      lightboxMeasureLiveDocumentKey = '';
      const liveCheckbox = $('lightboxUpdateMeasuresLive');
      if(liveCheckbox) liveCheckbox.checked = false;
    }
    stopLightboxMeasureLiveFn = stopLightboxMeasureLive;

    function startLightboxMeasureLive() {
      if(lightboxMeasureLiveTimer) return;
      lightboxMeasureLiveTimer = setInterval(() => {
        const panel = $('tab-lightbox');
        if(!panel || panel.classList.contains('hidden')) {
          stopLightboxMeasureLive();
          return;
        }
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
    window.addEventListener('focus', refreshArtboardScaleNotice);

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

    const letterW = $('letterLedWidthMm');
    const letterH = $('letterLedHeightMm');
    const letterRect = $('letterLedDepictionRect');
    const letterLine = $('letterLedDepictionLine');
    const letterText = $('letterLedDepictionText');
    function updateLetter() {
      const width = num((letterW && letterW.value) || 0) || 0;
      const height = num((letterH && letterH.value) || 0) || 0;
      if(letterRect) {
        const maxW = 100;
        const maxH = 40;
        let scale = 1;
        if(width > 0 && height > 0) {
          scale = Math.min(maxW / width, maxH / height, 1);
        }
        letterRect.style.width = Math.max(10, width * scale) + 'px';
        letterRect.style.height = Math.max(6, height * scale) + 'px';
      }
      if(letterLine) {
        letterLine.style.left = '0';
        letterLine.style.right = '0';
        letterLine.style.top = '50%';
        letterLine.style.bottom = 'auto';
        letterLine.style.width = '100%';
        letterLine.style.height = '1px';
      }
      if(letterText) {
        letterText.textContent = 'LED ' + width + 'mm Ã— ' + height + 'mm';
      }
    }
    if(letterW) letterW.addEventListener('input', updateLetter);
    if(letterH) letterH.addEventListener('input', updateLetter);
    updateLetter();
  }

  function wireLedCentreline() {
    const create = $('ledCreateLayout'), guidesOnly = $('ledCreateGuides'), clear = $('ledClearGenerated'), repopulate = $('ledRepopulateGuides'), cancel = $('ledCancelLayout');
    const status = $('ledLayoutStatus'), progress = $('ledProgressBar');
    function ledDebug(message, extra) {
      let line = '[SRH][LED] ' + String(message || '');
      if(extra !== undefined) {
        try {line += ' | ' + (typeof extra === 'string' ? extra : JSON.stringify(extra));}
        catch(_eLedDebug) {line += ' | [unserializable detail]';}
      }
      log(line);
    }
    ledDebug('wire start', {button: !!create, engine: !!window.LedCentreline});
    if(!create) {ledDebug('wire failed: Create Layout button not found'); return;}
    if(!window.LedCentreline) {
      ledDebug('wire failed: LedCentreline browser global is unavailable; check js/led-centreline.js loading');
      if(status) status.textContent = 'LED engine failed to load. See Testing console.';
      return;
    }
    let cancelled = false;
    function setProgress(text, fraction) {if(status) status.textContent=text;if(progress) progress.style.width=Math.max(0,Math.min(100,(fraction||0)*100))+'%';ledDebug('status', {text:String(text||''),progress:Math.round((fraction||0)*100)});}
    function queuedHost(functionCall, cb) {
      ledDebug('queue host call', String(functionCall || '').split('(')[0]);
      loadJSX(function() {callJSX(functionCall, function(result) {ledDebug('host response', String(result || '').slice(0, 500));if(cb) cb(result);});});
    }
    function contoursFromSvg(svgText, boundsMm) {
      const root = SvgParser.load(String(svgText || ''));
      SvgParser.config({tolerance: Math.max(0.05,num(($('ledCurveToleranceMm')&&$('ledCurveToleranceMm').value)||0.5))});
      const clean = SvgParser.clean();
      const nodes = clean.querySelectorAll('path,polygon,polyline,rect,circle,ellipse');
      const raw=[];
      Array.prototype.forEach.call(nodes,function(node){try{const p=SvgParser.polygonify(node);if(p&&p.length>2)raw.push({points:p.map(function(q){return{x:Number(q.x!=null?q.x:q.X),y:Number(q.y!=null?q.y:q.Y)};})});}catch(_eLedPoly){}});
      if(!raw.length) throw new Error('No filled contours could be read from the selection SVG.');
      let minX=Infinity,minY=Infinity,maxX=-Infinity,maxY=-Infinity;
      raw.forEach(function(c){c.points.forEach(function(p){if(p.x<minX)minX=p.x;if(p.x>maxX)maxX=p.x;if(p.y<minY)minY=p.y;if(p.y>maxY)maxY=p.y;});});
      const sx=Number(boundsMm.width)/(maxX-minX),sy=Number(boundsMm.height)/(maxY-minY);
      const contours = raw.map(function(c){return{points:c.points.map(function(p){return{x:(p.x-minX)*sx,y:(maxY-p.y)*sy};})};});
      ledDebug('SVG contours prepared', {contours:contours.length,widthMm:Number(boundsMm.width),heightMm:Number(boundsMm.height)});
      return contours;
    }
    function settings(guidesOnlyFlag) {return{cellMm:num($('ledRasterPrecisionMm').value)||1,maxCells:250000,widthMm:num($('letterLedWidthMm').value),heightMm:num($('letterLedHeightMm').value),clearanceMm:num($('ledEdgeClearanceMm').value),pruneMm:num($('ledBranchPruneMm').value),maxSpacingMm:num($('ledMaxCentreSpacingMm').value),endpointInsetMm:Math.max(num($('letterLedWidthMm').value),num($('letterLedHeightMm').value))/2,guidesOnly:!!guidesOnlyFlag};}
    function toDocument(result,payload) {
      const b=payload.sourceBounds,w=Number(payload.boundsMm.width),h=Number(payload.boundsMm.height);
      function map(p){return{x:Number(b.left)+(p.x/w)*(Number(b.right)-Number(b.left)),y:Number(b.bottom)+(p.y/h)*(Number(b.top)-Number(b.bottom)),angle:-Number(p.angle||0),sequence:p.sequence};}
      return{guides:result.guides.map(function(g){return g.map(map);}),placements:result.placements.map(function(r){return r.map(map);})};
    }
    function output(mapped, opts, preserveGuides) {
      let modules=0,min=Infinity,max=0,sum=0,spacingCount=0,series=0,unreachable=0;
      mapped.placements.forEach(function(run){modules+=run.length;const s=window.LedCentreline.splitSeries(run,parseInt($('letterMaxLedsInSeries').value,10)||50,num($('ledMaxWireReachMm').value));series+=s.series.length;unreachable+=s.unreachable.length;for(let i=1;i<run.length;i++){const dx=run[i].x-run[i-1].x,dy=run[i].y-run[i-1].y,d=Math.sqrt(dx*dx+dy*dy);min=Math.min(min,d);max=Math.max(max,d);sum+=d;spacingCount++;}});
      const anchor=mapped.guides.length&&mapped.guides[0].length?mapped.guides[0][0]:{x:0,y:0};
      const payload={layoutId:'led-'+Date.now(),sourceKey:'bounds-'+anchor.x.toFixed(3)+'-'+anchor.y.toFixed(3),profileCode:String($('letterLedCode').value||'generic'),settingsVersion:1,guides:mapped.guides,placements:opts.guidesOnly?mapped.placements.map(function(){return[];}):mapped.placements,moduleWidthMm:opts.widthMm,moduleHeightMm:opts.heightMm,moduleSvg:String($('letterLedSvgOverride').value||''),drawGuides:!!$('ledDrawGuides').checked,drawModules:!!$('ledDrawModules').checked,drawWiring:!!$('ledDrawWiring').checked,drawStats:!!$('ledDrawStats').checked,replacePrevious:!!$('ledReplacePrevious').checked,preserveGuides:!!preserveGuides,statsX:anchor.x,statsY:anchor.y,summary:'LED layout | '+modules+' modules | '+series+' series | spacing '+(spacingCount?min.toFixed(1)+'–'+max.toFixed(1)+' mm':'n/a')+' | estimated '+(modules*num($('letterLedWatt').value)).toFixed(2)+' W | unreachable gaps '+unreachable};
      ledDebug('output prepared', {guides:mapped.guides.length,modules:modules,series:series,unreachable:unreachable,preserveGuides:!!preserveGuides});
      const encoded=JSON.stringify(payload).replace(/\\/g,'\\\\').replace(/"/g,'\\"');queuedHost('signarama_helper_led_drawLayout("'+encoded+'")',function(res){setProgress(String(res||payload.summary),1);if(cancel)cancel.disabled=true;});
    }
    function run(guidesOnlyFlag) {cancelled=false;if(cancel)cancel.disabled=false;ledDebug('run requested',{guidesOnly:!!guidesOnlyFlag});setProgress('Extracting selected filled geometry…',.05);queuedHost('signarama_helper_led_extractSelectionGeometry()',function(res){let p;try{p=JSON.parse(String(res||'{}'));}catch(e){ledDebug('geometry response JSON parse failed',e&&e.message?e.message:e);p=null;}if(!p||!p.ok){ledDebug('geometry extraction failed',p||String(res||''));setProgress((p&&p.error)||'Geometry extraction failed.',0);return;}ledDebug('geometry extracted',{items:Number(p.itemCount||0),svgChars:String(p.svgText||'').length,boundsMm:p.boundsMm,scaleFactor:p.scaleFactor});setProgress('Computing bounded medial axis…',.25);setTimeout(function(){if(cancelled){setProgress('Cancelled before geometry computation.',0);return;}try{const contours=contoursFromSvg(p.svgText,p.boundsMm),opts=settings(guidesOnlyFlag);ledDebug('engine generate start',opts);const started=Date.now(),result=window.LedCentreline.generate(contours,opts),mapped=toDocument(result,p);ledDebug('engine generate complete',{elapsedMs:Date.now()-started,guides:result.guides.length,cells:result.grid&&result.grid.cells,cellMm:result.grid&&result.grid.cellMm,modules:result.placements.reduce(function(n,r){return n+r.length;},0)});localStorage.setItem('srhLedLastContours',JSON.stringify({contours:contours,payload:p,options:opts}));setProgress('Drawing editable guides and modules…',.8);output(mapped,opts,false);}catch(e){ledDebug('layout exception',e&&e.stack?e.stack:(e&&e.message?e.message:e));setProgress('LED layout failed: '+(e.message||e),0);}},0);});}
    create.onclick=function(){ledDebug('Create Layout clicked');run(false);};guidesOnly.onclick=function(){ledDebug('Create Guide Paths Only clicked');run(true);};if(cancel)cancel.onclick=function(){ledDebug('Cancel clicked');cancelled=true;cancel.disabled=true;setProgress('Cancellation requested; stopping between phases…',0);};
    if(clear)clear.onclick=function(){queuedHost('signarama_helper_led_clearGenerated()',function(res){setProgress(String(res||'Cleared.'),0);});};
    if(repopulate)repopulate.onclick=function(){setProgress('Reading edited LED guides…',.1);queuedHost('signarama_helper_led_captureGuides()',function(res){let p;try{p=JSON.parse(String(res||'{}'));}catch(e){p=null;}if(!p||!p.ok){setProgress((p&&p.error)||'Could not read guides.',0);return;}const saved=JSON.parse(localStorage.getItem('srhLedLastContours')||'null');if(!saved){setProgress('Original fill geometry is unavailable; create a layout before repopulating.',0);return;}const b=saved.payload.sourceBounds,w=Number(saved.payload.boundsMm.width),h=Number(saved.payload.boundsMm.height),local=p.guides.map(function(g){return g.map(function(q){return{x:(q.x-b.left)/(b.right-b.left)*w,y:(q.y-b.bottom)/(b.top-b.bottom)*h};});}),placements=local.map(function(g){return window.LedCentreline.placeModules(g,saved.contours,settings(false));}),mapped={guides:p.guides,placements:placements.map(function(run){return run.map(function(q){return{x:b.left+q.x/w*(b.right-b.left),y:b.bottom+q.y/h*(b.top-b.bottom),angle:q.angle,sequence:q.sequence};});})};output(mapped,settings(false),true);});};
    ledDebug('wire complete: LED actions ready');
  }

  function wireLedLetterSpecs() {
    const options = $('letterLedSpecOptions');
    const details = $('letterLedSpecDropdown');
    if(!options || !details) return;

    function svgUrl(svg) {
      return 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(String(svg || ''));
    }
    function setValue(id, value) {
      const el = $(id);
      if(el) el.value = value == null ? '' : value;
    }
    let currentSpec = null;
    function choose(spec, button) {
      currentSpec = spec;
      setValue('ledProfileBrand', spec.manufacturer || 'Generic example — verify');
      setValue('ledProfileSeries', spec.series || 'Example');
      setValue('letterLedCode', spec.code);
      setValue('letterLedVoltage', spec.voltage);
      setValue('letterLedWatt', spec.watt);
      setValue('letterLedWidthMm', spec.widthMm);
      setValue('letterLedHeightMm', spec.heightMm);
      setValue('letterLedSvgOverride', spec.svg);
      setValue('ledMaxCentreSpacingMm', spec.maxCentreSpacingMm || 100);
      setValue('ledMaxWireReachMm', spec.maxWireReachMm || 150);
      setValue('letterMaxLedsInSeries', spec.maxModulesPerSeries || 50);
      const name = $('letterLedSpecName');
      const meta = $('letterLedSpecMeta');
      const icon = $('letterLedSpecIcon');
      const summary = $('letterLedSpecSummary');
      if(name) name.textContent = spec.name || spec.code || 'LED module';
      if(meta) meta.textContent = (spec.code || '') + ' · ' + spec.widthMm + ' × ' + spec.heightMm + ' mm · ' + spec.watt + ' W · ' + spec.voltage + ' V';
      if(icon) icon.src = svgUrl(spec.svg);
      if(summary) summary.textContent = spec.name || spec.code || 'LED module';
      Array.prototype.forEach.call(options.querySelectorAll('.led-spec-option'), function(el) {el.setAttribute('aria-selected', el === button ? 'true' : 'false');});
      details.open = false;
    }
    let bundledSpecs = [];
    let localSpecs = [];
    function localCatalogPath() {
      const path = require('path');
      return path.join(cs.getSystemPath(SystemPath.USER_DATA), 'Signarama Helper', 'led-specs.json');
    }
    function readLocalSpecs() {
      try {
        const fs = require('fs');
        const parsed = JSON.parse(fs.readFileSync(localCatalogPath(), 'utf8'));
        return parsed && Array.isArray(parsed.leds) ? parsed.leds : [];
      } catch(_eLedLocalRead) {return [];}
    }
    function writeLocalSpecs(specs) {
      const fs = require('fs');
      const path = require('path');
      const filePath = localCatalogPath();
      fs.mkdirSync(path.dirname(filePath), {recursive: true});
      fs.writeFileSync(filePath, JSON.stringify({version: 1, leds: specs}, null, 2), 'utf8');
      return filePath;
    }
    function render() {
      const specs = bundledSpecs.concat(localSpecs);
      options.innerHTML = '';
      specs.forEach(function(spec, index) {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'led-spec-option';
        button.setAttribute('role', 'option');
        button.setAttribute('aria-selected', 'false');
        const icon = document.createElement('img');
        icon.alt = '';
        icon.src = svgUrl(spec.svg);
        const copy = document.createElement('div');
        const title = document.createElement('strong');
        title.textContent = spec.name || spec.code || 'LED module';
        const meta = document.createElement('div');
        meta.className = 'small';
        meta.textContent = (spec.code || '') + ' · ' + spec.widthMm + ' × ' + spec.heightMm + ' mm · ' + spec.watt + ' W · ' + spec.voltage + ' V';
        copy.appendChild(title);
        copy.appendChild(meta);
        button.appendChild(icon);
        button.appendChild(copy);
        button.addEventListener('click', function() {choose(spec, button);});
        options.appendChild(button);
        if(index === 0) choose(spec, button);
      });
      if(!specs.length) options.innerHTML = '<div class="small">No LED specifications found.</div>';
    }

    function showAddSpecModal(svgPayload) {
      const old = $('letterLedAddOverlay');
      if(old) old.remove();
      const bounds = svgPayload.boundsMm || {};
      const overlay = document.createElement('div');
      overlay.id = 'letterLedAddOverlay';
      overlay.style.cssText = 'position:fixed;inset:0;z-index:2147482000;background:rgba(0,0,0,.55);display:flex;align-items:center;justify-content:center;padding:16px;box-sizing:border-box;';
      overlay.innerHTML = '<div class="card" style="width:min(460px,100%);max-height:calc(100vh - 32px);overflow:auto;margin:0;">' +
        '<div class="cardTitle">Add Selection to LED Library</div><div class="small" style="margin:5px 0 12px;">The selected Illustrator artwork was converted to SVG. Enter its LED specifications.</div>' +
        '<div class="grid2">' +
        '<label class="fld"><span>Name</span><input id="addLedName" value="Custom LED Module"></label>' +
        '<label class="fld"><span>LED Code</span><input id="addLedCode" value="CUSTOM-LED"></label>' +
        '<label class="fld"><span>Width (mm)</span><input id="addLedWidth" type="number" step="0.1" value="' + Number(bounds.width || 0).toFixed(2) + '"></label>' +
        '<label class="fld"><span>Height (mm)</span><input id="addLedHeight" type="number" step="0.1" value="' + Number(bounds.height || 0).toFixed(2) + '"></label>' +
        '<label class="fld"><span>LED Watt</span><input id="addLedWatt" type="number" step="0.01" value="1.5"></label>' +
        '<label class="fld"><span>LED Voltage</span><input id="addLedVoltage" type="number" step="0.1" value="12"></label></div>' +
        '<div class="row" style="margin-top:14px;justify-content:flex-end;"><button id="cancelAddLedSpec" class="btn2" type="button">Cancel</button><button id="saveAddLedSpec" class="btn2" type="button">Save LED</button></div></div>';
      document.body.appendChild(overlay);
      $('cancelAddLedSpec').onclick = function() {overlay.remove();};
      $('saveAddLedSpec').onclick = function() {
        const spec = {
          name: String($('addLedName').value || '').trim(), code: String($('addLedCode').value || '').trim(),
          manufacturer: 'User profile', series: 'Local',
          widthMm: num($('addLedWidth').value), heightMm: num($('addLedHeight').value),
          watt: num($('addLedWatt').value), voltage: num($('addLedVoltage').value), svg: String(svgPayload.svgText || ''),
          maxCentreSpacingMm: 100, maxWireReachMm: 150, maxModulesPerSeries: 50
        };
        if(!spec.name || !spec.code || !(spec.widthMm > 0) || !(spec.heightMm > 0) || !(spec.watt > 0) || !(spec.voltage > 0)) {
          showToast('Complete all LED specification fields with valid values.', {type: 'warn', title: 'LED Letters'}); return;
        }
        try {
          localSpecs.push(spec);
          const savedPath = writeLocalSpecs(localSpecs);
          render();
          const buttons = options.querySelectorAll('.led-spec-option');
          choose(spec, buttons[buttons.length - 1]);
          overlay.remove();
          log('Saved custom LED specification to: ' + savedPath);
        } catch(e) {showToast('Could not save the LED library: ' + (e.message || e), {type: 'error', title: 'LED Letters'});}
      };
    }

    const addSelection = $('btnAddSelectionLedSpec');
    if(addSelection) addSelection.onclick = function() {
      loadJSX(function() {
        callJSX('signarama_helper_led_extractSelectionGeometry()', function(result) {
          let payload = null;
          try {payload = JSON.parse(String(result || '{}'));} catch(_eLedCaptureJson) {payload = null;}
          if(!payload || !payload.ok || !payload.svgText) {
            showToast((payload && payload.error) || 'Could not convert the current selection to SVG.', {type: 'error', title: 'LED Letters'}); return;
          }
          showAddSpecModal(payload);
        });
      });
    };

    const duplicate = $('ledProfileDuplicate');
    if(duplicate) duplicate.onclick = function() {
      if(!currentSpec) return;
      const copy = JSON.parse(JSON.stringify(currentSpec));
      copy.name = String(copy.name || copy.code || 'LED') + ' Copy'; copy.code = String(copy.code || 'CUSTOM') + '-COPY'; copy.manufacturer = 'User profile';
      localSpecs.push(copy);writeLocalSpecs(localSpecs);render();
    };
    const saveProfile = $('ledProfileSave');
    if(saveProfile) saveProfile.onclick = function() {
      const target = currentSpec && localSpecs.indexOf(currentSpec) >= 0 ? currentSpec : {};
      target.manufacturer=String($('ledProfileBrand').value||'User profile');target.series=String($('ledProfileSeries').value||'Local');target.name=String($('letterLedCode').value||'Custom LED');target.code=String($('letterLedCode').value||'CUSTOM');target.widthMm=num($('letterLedWidthMm').value);target.heightMm=num($('letterLedHeightMm').value);target.watt=num($('letterLedWatt').value);target.voltage=num($('letterLedVoltage').value);target.maxCentreSpacingMm=num($('ledMaxCentreSpacingMm').value);target.maxWireReachMm=num($('ledMaxWireReachMm').value);target.maxModulesPerSeries=parseInt($('letterMaxLedsInSeries').value,10)||50;target.svg=String($('letterLedSvgOverride').value||'');
      if(localSpecs.indexOf(target)<0)localSpecs.push(target);writeLocalSpecs(localSpecs);currentSpec=target;render();
    };
    const remove = $('ledProfileDelete');
    if(remove) remove.onclick = function() {
      if(!currentSpec) return;
      const index = localSpecs.indexOf(currentSpec);
      if(index < 0) {showToast('Bundled generic examples cannot be deleted; duplicate one to create a local profile.', {type:'warn',title:'LEDs'});return;}
      localSpecs.splice(index,1);writeLocalSpecs(localSpecs);render();
    };

    document.addEventListener('click', function(event) {
      if(details.open && !details.contains(event.target)) details.open = false;
    });

    const request = new XMLHttpRequest();
    request.open('GET', 'data/led-specs.json', true);
    request.onreadystatechange = function() {
      if(request.readyState !== 4) return;
      if(request.status === 0 || (request.status >= 200 && request.status < 300)) {
        try {
          const parsed = JSON.parse(request.responseText);
          bundledSpecs = parsed && Array.isArray(parsed.leds) ? parsed.leds : [];
          localSpecs = readLocalSpecs();
          render();
          return;
        } catch(_eLedSpecJson) { }
      }
      localSpecs = readLocalSpecs();
      render();
      if(!localSpecs.length) options.innerHTML = '<div class="small">Could not load the bundled or local LED library.</div>';
    };
    request.send();
  }

  function wireColours() {
    const refreshBtn = $('btnRefreshColours');
    if(refreshBtn) refreshBtn.onclick = () => refreshColours({showToastOnComplete: true});
    const settingsBtn = $('btnColourSettings');
    const settingsPanel = $('colourSettingsPanel');
    if(settingsBtn && settingsPanel) settingsBtn.onclick = () => {
      const show = settingsPanel.classList.contains('hidden');
      settingsPanel.classList.toggle('hidden', !show);
      settingsBtn.setAttribute('aria-expanded', show ? 'true' : 'false');
    };
    ['richBlackC', 'richBlackM', 'richBlackY', 'richBlackK'].forEach(id => {const input = $(id); if(input) input.addEventListener('change', () => refreshColours());});
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

      const progressWrap = $('coloursScanProgressWrap');
      const progressBar = $('coloursScanProgress');
      const progressText = $('coloursScanProgressText');
      if(progressWrap) progressWrap.classList.remove('hidden');
      function scanStep(command) { callJSX(command, function(res) {
        if(!res) {
          log('Colours: no response from JSX.');
          if(showToastOnComplete) showToast('No response from colour scan.', {type: 'warn', title: 'Refresh colours'});
          if(progressWrap) progressWrap.classList.add('hidden'); return;
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
          if(progressWrap) progressWrap.classList.add('hidden'); return;
        }
        if(data && typeof data.position === 'number') {
          if(progressBar) {progressBar.max = Math.max(1, data.total || 0); progressBar.value = data.position || 0;}
          if(progressText) progressText.textContent = 'Scanning objects: ' + (data.position || 0) + ' / ' + (data.total || 0);
          if(!data.done) {setTimeout(() => scanStep('signarama_helper_stepDocumentColorScan(125)'), 0); return;}
          if(progressWrap) progressWrap.classList.add('hidden');
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
          const target = getRichBlackTarget();
          return Math.abs(cmyk.c - target.c) <= tol &&
            Math.abs(cmyk.m - target.m) <= tol &&
            Math.abs(cmyk.y - target.y) <= tol &&
            Math.abs(cmyk.k - target.k) <= 3;
        }
        function getRichBlackTarget() {
          return {c:clamp(num(($('richBlackC') && $('richBlackC').value) || 60),0,100), m:clamp(num(($('richBlackM') && $('richBlackM').value) || 60),0,100), y:clamp(num(($('richBlackY') && $('richBlackY').value) || 60),0,100), k:clamp(num(($('richBlackK') && $('richBlackK').value) || 100),0,100)};
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
          label.style.minWidth = '46px';
          label.style.flex = '0 0 46px';
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
          const copyButton = document.createElement('button');
          copyButton.type = 'button'; copyButton.className = 'btn2'; copyButton.textContent = 'Copy';
          copyButton.title = 'Copy this row colour values'; copyButton.style.padding = '3px 7px'; copyButton.style.flex = '0 0 auto';
          copyButton.addEventListener('click', () => {
            copiedColourValues = {mode: colourEditState.mode, values: inputs.map(inp => inp.value)};
            Array.prototype.forEach.call(list.querySelectorAll('[data-colour-paste]'), btn => {btn.disabled = btn.dataset.colourMode !== copiedColourValues.mode || Number(btn.dataset.colourCount) !== copiedColourValues.values.length;});
            showToast('Colour values copied.', {type:'success', title:'Colours'});
          });
          const pasteButton = document.createElement('button');
          pasteButton.type = 'button'; pasteButton.className = 'btn2'; pasteButton.textContent = 'Paste';
          pasteButton.title = 'Paste copied colour values into this row'; pasteButton.style.padding = '3px 7px'; pasteButton.style.flex = '0 0 auto';
          pasteButton.setAttribute('data-colour-paste', 'true');
          pasteButton.dataset.colourMode = colourEditState.mode; pasteButton.dataset.colourCount = String(inputs.length);
          pasteButton.disabled = !copiedColourValues || copiedColourValues.mode !== colourEditState.mode || copiedColourValues.values.length !== inputs.length;
          pasteButton.addEventListener('click', () => {
            if(!copiedColourValues || copiedColourValues.mode !== colourEditState.mode) {showToast('Copy a colour from this document mode first.', {type:'warn', title:'Colours'}); return;}
            inputs.forEach((inp, idx) => {if(copiedColourValues.values[idx] != null) inp.value = copiedColourValues.values[idx];});
            markRowDirty(); previewSwatch();
            showToast('Colour values pasted. Click Apply to update the document.', {type:'success', title:'Colours'});
          });
          row.appendChild(copyButton);
          row.appendChild(pasteButton);
          if(showBlackHazard) {
            const richBlackTarget = getRichBlackTarget();
            const richBlackText = 'C' + round1(richBlackTarget.c) + ' M' + round1(richBlackTarget.m) + ' Y' + round1(richBlackTarget.y) + ' K' + round1(richBlackTarget.k);
            const hazard = document.createElement('span');
            hazard.textContent = '\u26A0';
            hazard.title = 'Black colour is not close to target rich black (' + richBlackText + ').';
            hazard.style.color = '#ffba00';
            hazard.style.fontWeight = '800';
            hazard.style.fontSize = '16px';
            hazard.style.marginLeft = '2px';
            hazard.style.flex = '0 0 auto';
            row.appendChild(hazard);
            if(colourEditState.mode === 'CMYK') {
              const fix = document.createElement('button');
              fix.type = 'button'; fix.className = 'btn2'; fix.textContent = '\u2692';
              fix.title = 'Fix rich black to ' + richBlackText;
              fix.style.padding = '2px 6px'; fix.style.minWidth = '28px'; fix.style.flex = '0 0 auto';
              fix.addEventListener('click', () => {
                inputs[0].value = richBlackTarget.c; inputs[1].value = richBlackTarget.m; inputs[2].value = richBlackTarget.y; inputs[3].value = richBlackTarget.k;
                markRowDirty(); previewSwatch();
                showToast('Rich black values populated (' + richBlackText + '). Click Apply to update the document.', {type:'success', title:'Colours'});
              });
              row.appendChild(fix);
            }
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
      }); }
      scanStep('signarama_helper_beginDocumentColorScan()');
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
    const profileSelect = $('dimensionThemeProfileSelect');
    const profilesField = $('dimensionThemeProfilesJson');
    const saveProfileBtn = $('btnDimensionThemeSave');
    const deleteProfileBtn = $('btnDimensionThemeDelete');
    const themeGrid = $('dimensionThemeGrid');
    let refreshDimensionSelectionHintTimer = null;
    let customProfiles = [];

    function sanitizeHex(hex, fallback) {
      const raw = String(hex || '').trim();
      if(/^#[0-9a-f]{6}$/i.test(raw)) return raw.toLowerCase();
      return String(fallback || '#000000').toLowerCase();
    }

    function buildBuiltinProfiles() {
      return [
        {id: 'builtin:0', name: 'Black / Black', textColor: '#000000', lineColor: '#000000', builtin: true},
        {id: 'builtin:1', name: 'Pink / Pink', textColor: '#ff00ff', lineColor: '#ff00ff', builtin: true},
        {id: 'builtin:2', name: 'Red / Red', textColor: '#ff0000', lineColor: '#ff0000', builtin: true},
        {id: 'builtin:3', name: 'Red / Black', textColor: '#ff0000', lineColor: '#000000', builtin: true},
        {id: 'builtin:4', name: 'White / White', textColor: '#ffffff', lineColor: '#ffffff', builtin: true}
      ];
    }

    function parseCustomProfiles() {
      if(!profilesField) return [];
      const raw = String(profilesField.value || '').trim();
      if(!raw) return [];
      try {
        const parsed = JSON.parse(raw);
        if(!Array.isArray(parsed)) return [];
        return parsed
          .map((entry, idx) => {
            const name = String((entry && entry.name) || '').trim();
            if(!name) return null;
            return {
              id: 'custom:' + name.toLowerCase(),
              name: name,
              textColor: sanitizeHex(entry.textColor, '#000000'),
              lineColor: sanitizeHex(entry.lineColor, '#000000'),
              builtin: false,
              order: idx
            };
          })
          .filter(Boolean);
      } catch(_eParseProfiles) {
        return [];
      }
    }

    function persistCustomProfiles(nextProfiles) {
      customProfiles = (nextProfiles || []).slice(0);
      if(!profilesField) return;
      profilesField.value = JSON.stringify(customProfiles.map((entry) => ({
        name: entry.name,
        textColor: entry.textColor,
        lineColor: entry.lineColor
      })));
      try {profilesField.dispatchEvent(new Event('input', {bubbles: true}));} catch(_eProfInp) { }
      try {profilesField.dispatchEvent(new Event('change', {bubbles: true}));} catch(_eProfChg) { }
    }

    function getAllProfiles() {
      return buildBuiltinProfiles().concat(customProfiles);
    }

    function updateDeleteButtonState() {
      if(!deleteProfileBtn) return;
      const selectedId = profileSelect ? String(profileSelect.value || '') : '';
      deleteProfileBtn.disabled = !/^custom:/i.test(selectedId);
    }

    function syncProfileSelectionToCurrentColours() {
      if(!profileSelect || !textColorInput || !lineColorInput) return;
      const textHex = sanitizeHex(textColorInput.value, '#000000');
      const lineHex = sanitizeHex(lineColorInput.value, '#000000');
      const match = getAllProfiles().find((entry) => entry.textColor === textHex && entry.lineColor === lineHex);
      profileSelect.value = match ? match.id : '';
      updateDeleteButtonState();
    }

    function renderProfileOptions(preferredId) {
      if(!profileSelect) return;
      const currentId = String(preferredId || profileSelect.value || '');
      const allProfiles = getAllProfiles();
      profileSelect.innerHTML = '';

      const currentOption = document.createElement('option');
      currentOption.value = '';
      currentOption.textContent = 'Current colours';
      profileSelect.appendChild(currentOption);

      const builtins = allProfiles.filter((entry) => entry.builtin);
      const customs = allProfiles.filter((entry) => !entry.builtin);
      if(builtins.length) {
        const builtinsGroup = document.createElement('optgroup');
        builtinsGroup.label = 'Built-in';
        builtins.forEach((entry) => {
          const opt = document.createElement('option');
          opt.value = entry.id;
          opt.textContent = entry.name;
          builtinsGroup.appendChild(opt);
        });
        profileSelect.appendChild(builtinsGroup);
      }
      if(customs.length) {
        const customGroup = document.createElement('optgroup');
        customGroup.label = 'Custom';
        customs.forEach((entry) => {
          const opt = document.createElement('option');
          opt.value = entry.id;
          opt.textContent = entry.name;
          customGroup.appendChild(opt);
        });
        profileSelect.appendChild(customGroup);
      }

      const hasPreferred = currentId && allProfiles.some((entry) => entry.id === currentId);
      if(hasPreferred) profileSelect.value = currentId;
      else syncProfileSelectionToCurrentColours();
      updateDeleteButtonState();
    }

    function applyProfileSelection(profileId) {
      if(!profileId || !textColorInput || !lineColorInput) {
        updateDeleteButtonState();
        return;
      }
      const profile = getAllProfiles().find((entry) => entry.id === profileId);
      if(!profile) {
        syncProfileSelectionToCurrentColours();
        return;
      }
      textColorInput.value = profile.textColor;
      lineColorInput.value = profile.lineColor;
      try {textColorInput.dispatchEvent(new Event('change', {bubbles: true}));} catch(_eDimTc) { }
      try {lineColorInput.dispatchEvent(new Event('change', {bubbles: true}));} catch(_eDimLc) { }
      profileSelect.value = profile.id;
      updateDeleteButtonState();
      if(schedulePanelSettingsSave) schedulePanelSettingsSave();
    }

    function deleteCustomProfile(profileId) {
      if(!/^custom:/i.test(String(profileId || ''))) {
        showToast('Only custom colour profiles can be deleted.', {type: 'warn', title: 'Dimensions'});
        updateDeleteButtonState();
        return;
      }
      const selectedProfile = customProfiles.find((entry) => entry.id === profileId);
      if(!selectedProfile) {
        updateDeleteButtonState();
        return;
      }
      persistCustomProfiles(customProfiles.filter((entry) => entry.id !== profileId));
      renderProfileOptions('');
      renderQuickThemeTiles();
      syncProfileSelectionToCurrentColours();
      showToast('Deleted dimension colour profile "' + selectedProfile.name + '".', {type: 'success', title: 'Dimensions'});
    }

    function renderQuickThemeTiles() {
      if(!themeGrid) return;
      themeGrid.innerHTML = '';
      getAllProfiles().forEach((profile) => {
        const tile = document.createElement('button');
        tile.className = 'theme-tile';
        tile.type = 'button';
        tile.title = profile.name;
        tile.setAttribute('data-profile-id', profile.id);
        tile.setAttribute('data-text-color', profile.textColor);
        tile.setAttribute('data-line-color', profile.lineColor);

        const inner = document.createElement('span');
        inner.className = 'theme-tile-inner';
        const textHalf = document.createElement('span');
        textHalf.className = 'theme-half';
        textHalf.style.background = profile.textColor;
        const lineHalf = document.createElement('span');
        lineHalf.className = 'theme-half';
        lineHalf.style.background = profile.lineColor;
        inner.appendChild(textHalf);
        inner.appendChild(lineHalf);

        const label = document.createElement('div');
        label.className = 'theme-tile-label';
        label.textContent = profile.name;

        tile.appendChild(inner);
        tile.appendChild(label);
        tile.addEventListener('click', () => {
          if(textColorInput) textColorInput.value = profile.textColor;
          if(lineColorInput) lineColorInput.value = profile.lineColor;
          try {if(textColorInput) textColorInput.dispatchEvent(new Event('change', {bubbles: true}));} catch(_eTC) { }
          try {if(lineColorInput) lineColorInput.dispatchEvent(new Event('change', {bubbles: true}));} catch(_eLC) { }
          if(profileSelect) profileSelect.value = profile.id;
          syncProfileSelectionToCurrentColours();
          if(schedulePanelSettingsSave) schedulePanelSettingsSave();
        });

        if(!profile.builtin) {
          const deleteBtn = document.createElement('button');
          deleteBtn.type = 'button';
          deleteBtn.className = 'theme-tile-delete';
          deleteBtn.title = 'Delete "' + profile.name + '"';
          deleteBtn.setAttribute('aria-label', 'Delete ' + profile.name);
          deleteBtn.textContent = 'x';
          deleteBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            deleteCustomProfile(profile.id);
          });
          tile.appendChild(deleteBtn);
        }

        themeGrid.appendChild(tile);
      });
    }

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

    function runBetweenMeasure(axis, centers) {
      const payload = buildDimensionPayload();
      payload.axis = axis;
      payload.centers = !!centers;
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      runDimensionOperation('atlas_dimensions_runBetween("' + json + '")', {
        logFn: logDim,
        toastTitle: 'Dimensions',
        onResult: function() {
          if(scheduleDimensionSelectionHintRefresh) scheduleDimensionSelectionHintRefresh();
        }
      });
    }

    function scheduleRefreshDimensionSelectionHint() {
      if(refreshDimensionSelectionHintTimer) clearTimeout(refreshDimensionSelectionHintTimer);
      refreshDimensionSelectionHintTimer = setTimeout(function() {
        refreshDimensionSelectionHintTimer = null;
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
      btnAreaMeasure: () => runAreaMeasure(),
      btnBetweenWidth: () => runBetweenMeasure('horizontal', false),
      btnBetweenHeight: () => runBetweenMeasure('vertical', false),
      btnCentersWidth: () => runBetweenMeasure('horizontal', true),
      btnCentersHeight: () => runBetweenMeasure('vertical', true)
    };

    Object.keys(map).forEach(id => {
      const el = $(id);
      if(el) el.onclick = map[id];
    });

    customProfiles = parseCustomProfiles();
    renderProfileOptions();
    renderQuickThemeTiles();

    if(profilesField) {
      const refreshProfilesFromField = () => {
        customProfiles = parseCustomProfiles();
        renderProfileOptions();
        renderQuickThemeTiles();
      };
      profilesField.addEventListener('input', refreshProfilesFromField);
      profilesField.addEventListener('change', refreshProfilesFromField);
    }

    if(profileSelect) {
      profileSelect.addEventListener('change', () => {
        applyProfileSelection(String(profileSelect.value || ''));
      });
    }

    function handleColourInputChanged() {
      syncProfileSelectionToCurrentColours();
      if(schedulePanelSettingsSave) schedulePanelSettingsSave();
    }
    if(textColorInput) {
      textColorInput.addEventListener('input', handleColourInputChanged);
      textColorInput.addEventListener('change', handleColourInputChanged);
    }
    if(lineColorInput) {
      lineColorInput.addEventListener('input', handleColourInputChanged);
      lineColorInput.addEventListener('change', handleColourInputChanged);
    }

    if(saveProfileBtn) {
      saveProfileBtn.addEventListener('click', () => {
        const selectedId = profileSelect ? String(profileSelect.value || '') : '';
        const existingProfile = getAllProfiles().find((entry) => entry.id === selectedId);
        const defaultName = existingProfile ? existingProfile.name : 'New profile';
        const rawName = window.prompt('Save current colours as profile name:', defaultName);
        const name = String(rawName || '').trim();
        if(!name) return;
        const normalizedId = 'custom:' + name.toLowerCase();
        const nextEntry = {
          id: normalizedId,
          name: name,
          textColor: sanitizeHex(textColorInput && textColorInput.value, '#000000'),
          lineColor: sanitizeHex(lineColorInput && lineColorInput.value, '#000000'),
          builtin: false
        };
        const nextProfiles = customProfiles.slice(0);
        const existingIdx = nextProfiles.findIndex((entry) => entry.id === normalizedId);
        if(existingIdx >= 0) nextProfiles[existingIdx] = nextEntry;
        else nextProfiles.push(nextEntry);
        persistCustomProfiles(nextProfiles);
        renderProfileOptions(nextEntry.id);
        renderQuickThemeTiles();
        if(profileSelect) profileSelect.value = nextEntry.id;
        updateDeleteButtonState();
        showToast('Saved dimension colour profile "' + name + '".', {type: 'success', title: 'Dimensions'});
      });
    }

    if(deleteProfileBtn) {
      deleteProfileBtn.addEventListener('click', () => {
        const selectedId = profileSelect ? String(profileSelect.value || '') : '';
        deleteCustomProfile(selectedId);
      });
    }

    syncProfileSelectionToCurrentColours();

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
    const moveModeField = $('transformMoveMode');
    const moveXField = $('transformMoveX');
    const moveYField = $('transformMoveY');
    const moveBtn = $('btnTransformMove');

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

    if(moveBtn) moveBtn.onclick = () => {
      const xRaw = String((moveXField && moveXField.value) || '').trim();
      const yRaw = String((moveYField && moveYField.value) || '').trim();
      if(!xRaw && !yRaw) {
        showToast('Enter an X and/or Y value in mm.', {type: 'warn', title: 'Transform'});
        return;
      }
      const payload = {
        mode: (modeField && modeField.value) || 'selection',
        moveMode: (moveModeField && moveModeField.value) || 'amount',
        xMm: xRaw || null,
        yMm: yRaw || null,
        excludeStroke: !!(excludeStrokeField && excludeStrokeField.checked),
        artboardIndices: parseArtboardIndicesField()
      };
      if(payload.mode === 'artboards' && (!payload.artboardIndices || !payload.artboardIndices.length)) {
        showToast('Select one or more target artboards in the list.', {type: 'warn', title: 'Transform'});
        return;
      }
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      loadJSX(function() {
        runButtonJsxOperation('atlas_transform_move("' + json + '")', {logFn: log, toastTitle: 'Transform'});
      });
    };
  }

  function wireFixings() {
    const createBtn = $('btnCreateFixingHoles');
    const modeField = $('fixingSpacingMode');
    const spacingField = $('fixingSpacingMm');
    const quantityField = $('fixingQuantity');
    const cornersField = $('fixingIncludeCorners');
    const logField = $('fixingsLog');
    function updateModeUi() {
      const mode = (modeField && modeField.value) || 'corners';
      if(spacingField) spacingField.disabled = mode === 'quantity' || mode === 'corners';
      if(quantityField) quantityField.disabled = mode !== 'quantity';
      if(cornersField) {
        if(mode === 'corners') cornersField.checked = true;
        cornersField.disabled = mode === 'corners';
      }
    }
    if(modeField) modeField.addEventListener('change', updateModeUi);
    updateModeUi();
    if(!createBtn) return;
    createBtn.onclick = () => {
      const payload = {
        diameterMm: num(($('fixingDiameterMm') && $('fixingDiameterMm').value) || 0),
        insetMm: num(($('fixingInsetMm') && $('fixingInsetMm').value) || 0),
        includeCorners: !!($('fixingIncludeCorners') && $('fixingIncludeCorners').checked),
        spacingMode: (modeField && modeField.value) || 'maximum',
        spacingMm: num((spacingField && spacingField.value) || 0),
        quantity: parseInt((quantityField && quantityField.value) || 0, 10)
      };
      if(payload.spacingMode === 'corners') payload.includeCorners = true;
      if(!(payload.diameterMm > 0)) {
        showToast('Enter a hole diameter greater than zero.', {type: 'warn', title: 'Fixings'});
        return;
      }
      if(payload.spacingMode !== 'quantity' && payload.spacingMode !== 'corners' && !(payload.spacingMm > 0)) {
        showToast('Enter a spacing greater than zero.', {type: 'warn', title: 'Fixings'});
        return;
      }
      const json = JSON.stringify(payload).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      loadJSX(function() {
        runButtonJsxOperation('signarama_helper_createFixingHoles("' + json + '")', {
          logFn: function(message) {
            if(logField) logField.textContent = String(message || '');
            log(message);
          },
          toastTitle: 'Fixings'
        });
      });
    };
  }

  function wirePreflight() {
    const runBtn = $('btnRunPreflight');
    const firstStep = document.querySelector('[data-preflight-step="1"]');
    const selectBtn = $('btnPreflightSelectIssues');
    const highlightBtn = $('btnPreflightHighlightIssues');
    const retryBtn = $('btnPreflightRetryDoubleCuts');
    const issueActions = $('preflightDoubleCutActions');
    let lastIssues = [];
    function setStep(state) {
      if(!firstStep) return;
      firstStep.classList.toggle('running', state === 'running');
      firstStep.classList.toggle('complete', state === 'complete');
      firstStep.classList.toggle('failed', state === 'failed');
      const icon = firstStep.querySelector('.preflight-step-icon');
      if(icon) icon.textContent = state === 'complete' ? '✓' : (state === 'failed' ? '×' : (state === 'running' ? '…' : '1'));
    }
    function issuePathIndices() {
      const values = [];
      lastIssues.forEach(issue => (issue.pathIndices || []).forEach(index => {if(values.indexOf(index) < 0) values.push(index);}));
      return values;
    }
    function runDoubleCutCheck() {
      if(!runBtn || runBtn.disabled) return;
      if(typeof PreflightLogic === 'undefined') {
        showToast('Preflight comparison engine did not load.', {type: 'error', title: 'Preflight'});
        return;
      }
      const colourName = String(($('preflightCutColourName') && $('preflightCutColourName').value) || 'CutContour').trim();
      const toleranceMm = num(($('preflightToleranceMm') && $('preflightToleranceMm').value) || 0.02);
      const minimumMm = num(($('preflightMinimumOverlapMm') && $('preflightMinimumOverlapMm').value) || 0.1);
      const gridMm = num(($('preflightGridCellMm') && $('preflightGridCellMm').value) || 5);
      if(!colourName || !(toleranceMm > 0) || !(minimumMm > 0) || !(gridMm > 0)) {
        showToast('Enter a colour name and positive preflight geometry settings.', {type: 'warn', title: 'Preflight'});
        return;
      }
      runBtn.disabled = true; setStep('running'); lastIssues = [];
      if(issueActions) issueActions.classList.add('hidden');
      if(selectBtn) selectBtn.disabled = true;
      if(highlightBtn) highlightBtn.disabled = true;
      loadJSX(function() {
        const geometry = {colourName: colourName, pathCount: 0, paths: []};
        function failGeometry(message) {
          runBtn.disabled = false; setStep('idle');
          showToast(String(message || 'Could not read cut path geometry.'), {type: 'error', title: 'Preflight'});
        }
        function collectGeometryPage(startIndex) {
          const request = jsxEscapeDoubleQuoted(JSON.stringify({colourName: colourName, startIndex: startIndex, batchSize: 25}));
          const command = '((typeof signarama_helper_preflight_extractCutGeometry === "function") ? signarama_helper_preflight_extractCutGeometry : ((typeof $ !== "undefined" && $.global && typeof $.global.signarama_helper_preflight_extractCutGeometry === "function") ? $.global.signarama_helper_preflight_extractCutGeometry : function(){return "{\\"error\\":\\"Preflight geometry function not loaded.\\"}";}))("' + request + '")';
          callJSX(command, function(raw) {
            let page;
            const rawText = String(raw || '');
            try {page = JSON.parse(rawText);} catch(_ePfJson) {
              // Some Illustrator/CEP combinations serialize ExtendScript objects
              // as parenthesized JavaScript object literals rather than strict JSON.
              try {page = Function('return ' + rawText)();} catch(_ePfObjectLiteral) {
                failGeometry('Could not read cut path geometry. Illustrator returned: ' + String(raw || '(empty response)').slice(0, 240));
                return;
              }
            }
            if(page.error) {failGeometry('Could not read cut path geometry: ' + page.error); return;}
            geometry.paths = geometry.paths.concat(page.paths || []);
            geometry.pathCount = geometry.paths.length;
            if(!page.done && Number(page.nextIndex) > startIndex) {
              setTimeout(function() {collectGeometryPage(Number(page.nextIndex));}, 0);
              return;
            }
            compareCollectedGeometry();
          });
        }
        function compareCollectedGeometry() {
          const ptPerMm = 72 / 25.4;
          PreflightLogic.findOverlapsAsync(geometry.paths, {
            tolerancePt: toleranceMm * ptPerMm,
            minimumPt: minimumMm * ptPerMm,
            gridCellPt: gridMm * ptPerMm
          }).then(result => {
            lastIssues = result.issues || [];
            if(selectBtn) selectBtn.disabled = !lastIssues.length;
            if(highlightBtn) highlightBtn.disabled = !lastIssues.length;
            if(issueActions) issueActions.classList.toggle('hidden', !lastIssues.length);
            setStep(lastIssues.length ? 'failed' : 'complete'); runBtn.disabled = false;
          }).catch(err => {
            setStep('idle'); runBtn.disabled = false;
            showToast('Double-cut comparison failed: ' + (err && err.message ? err.message : err), {type: 'error', title: 'Preflight'});
          });
        }
        collectGeometryPage(0);
      });
    }
    if(runBtn) runBtn.onclick = runDoubleCutCheck;
    if(firstStep) firstStep.onclick = runDoubleCutCheck;
    if(firstStep) firstStep.onkeydown = (event) => {
      if(event.key === 'Enter' || event.key === ' ') {event.preventDefault(); runDoubleCutCheck();}
    };
    if(retryBtn) retryBtn.onclick = (event) => {event.stopPropagation(); runDoubleCutCheck();};
    if(selectBtn) selectBtn.onclick = (event) => {
      event.stopPropagation();
      const json = jsxEscapeDoubleQuoted(JSON.stringify(issuePathIndices()));
      runButtonJsxOperation('signarama_helper_preflight_selectIssues("' + json + '")', {logFn: log, toastTitle: 'Preflight'});
    };
    if(highlightBtn) highlightBtn.onclick = (event) => {
      event.stopPropagation();
      const regions = lastIssues.map(issue => ({a: issue.a, b: issue.b}));
      const json = jsxEscapeDoubleQuoted(JSON.stringify(regions));
      runButtonJsxOperation('signarama_helper_preflight_highlightIssues("' + json + '")', {logFn: log, toastTitle: 'Preflight'});
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
      enqueueEvalScript(cmd, function(res) {
        var txt = String(res || '');
        log('JSX load result: ' + txt);
        if(done) done(txt);
      });
    } catch(e) {
      log('Failed to load JSX: ' + (e && e.message ? e.message : e));
    }
  }

  const jsxRequestQueue = [];
  let jsxRequestInFlight = false;

  function drainJSXQueue() {
    if(jsxRequestInFlight || !jsxRequestQueue.length) return;
    jsxRequestInFlight = true;
    const request = jsxRequestQueue.shift();
    function finish(res) {
      try {
        if(request.callback) request.callback(res);
      } catch(e) {
        log('Panel JSX callback error: ' + (e && e.message ? e.message : e));
      } finally {
        jsxRequestInFlight = false;
        drainJSXQueue();
      }
    }
    try {
      cs.evalScript(request.script, finish);
    } catch(e) {
      finish('Error: ' + (e && e.message ? e.message : e));
    }
  }

  function enqueueEvalScript(script, callback) {
    jsxRequestQueue.push({script: script, callback: callback});
    drainJSXQueue();
  }

  function callJSX(fnCall, cb) {
    try {
      var wrapped = '(function(){try{return ' + fnCall + ' }catch(e){return "Error: " + e}})()';
      enqueueEvalScript(wrapped, cb);
    } catch(e) {
      if(cb) cb('Error: ' + (e && e.message ? e.message : e));
    }
  }
})();
