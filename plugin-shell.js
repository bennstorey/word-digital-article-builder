/**
 * Word & Web → Digital — WoodWing Studio plug-in
 *
 * Content Station SDK plug-in that converts Top Gear AN+ Word documents
 * (.docx) into digital articles (.digital). Single entry point: a button in
 * the Dossier toolbar that parses a Word doc and creates the digital article
 * directly inside the current Dossier via the workflow API (upload through
 * the Transfer Server, CreateObjects with a 'Contained' relation, C_HEADLINE
 * set from the feed headline, component set / Look and Feel / Twixl id from
 * BRAND_DEFAULTS). The standalone web version — index.html on GitHub Pages —
 * offers a .digital file download when no Dossier context is wanted.
 *
 * GENERATED FILE — do not edit directly. The conversion engine is extracted
 * from index.html by build-plugin.js; edit there and rebuild.
 *
 * Registration (Studio Server Management Console):
 *   Integrations → Studio → Plug-ins → Studio → Add new → URL of this file.
 */
(function () {
  'use strict';

  if (typeof ContentStationSdk === 'undefined') {
    console.error('[word-digital] ContentStationSdk not available — plug-in not loaded in a Studio context.');
    return;
  }

  // ─── Conversion engine (generated from index.html) ────────────────────────
  // The plug-in runs on Studio's origin, so /proxy must be absolute.
  // Set this to your deployed Fly app before shipping.
  window.PROXY_BASE = window.PROXY_BASE || 'https://topgear-web-word-digital.fly.dev';

  /*__ENGINE__*/
  // ─── End conversion engine ────────────────────────────────────────────────

  var DIGITAL_MIME = 'application/ww-digital+json';

  // Per-brand defaults applied to created digital articles, keyed by
  // Publication (Brand) id. GUIDs verified against existing articles on
  // lab-studio.woodwing.cloud (all current TG AN+ articles carry these).
  var BRAND_DEFAULTS = {
    // Top Gear
    '3': {
      componentSet: '11bd53cb-47fd-4040-8a62-486e7eb7850e',  // Default component set
      lookAndFeel: 'cc7a498a-7980-4d69-875b-06533c881d77',   // TG-custom-styles-ISSUE-APPLE 2026
      twixlCollectionId: '102069',
    },
  };

  // Word parsing dependency, loaded on demand and kept plugin-local per the
  // SDK guidance on managing external dependencies.
  var mammothPromise = null;
  function loadMammoth() {
    if (window.mammoth) return Promise.resolve(window.mammoth);
    if (!mammothPromise) {
      mammothPromise = new Promise(function (resolve, reject) {
        var s = document.createElement('script');
        s.src = 'https://cdn.jsdelivr.net/npm/mammoth@1.6.0/mammoth.browser.min.js';
        s.onload = function () { resolve(window.mammoth); };
        s.onerror = function () {
          mammothPromise = null;
          reject(new Error('Could not load the Word parser (mammoth.js). Check that cdn.jsdelivr.net is reachable.'));
        };
        document.head.appendChild(s);
      });
    }
    return mammothPromise;
  }

  // ─── Studio Server API (same origin — session comes from ticket or cookie) ─
  function getTicket() {
    try {
      var info = ContentStationSdk.getInfo();
      return (info && info.Ticket) || '';
    } catch (e) { return ''; }
  }

  function serverIndexUrl() {
    var rel = (window.csConfig && window.csConfig.serverUrl) || '../server/index.php';
    return new URL(rel, window.location.href).href;
  }

  function transferUrl() {
    try {
      var fs = ContentStationSdk.getInfo().ServerInfo.FeatureSet || [];
      for (var i = 0; i < fs.length; i++) {
        if (fs[i].Key === 'FileUploadUrl' && fs[i].Value) return fs[i].Value;
      }
    } catch (e) { /* fall through */ }
    return serverIndexUrl().replace(/index\.php.*$/, 'transferindex.php');
  }

  // Studio uses cookie-based sessions on current Studio Server versions:
  // requests authenticate via the session cookie plus the X-WoodWing-Application
  // header (CSRF guard), with Ticket set to null in the payload. On older
  // ticket-based setups getInfo().Ticket is populated and used instead.
  var WW_APP_HEADER = { 'X-WoodWing-Application': 'Content Station' };

  function callServer(method, params) {
    params.Ticket = getTicket() || null;
    return fetch(serverIndexUrl() + '?protocol=JSON', {
      method: 'POST',
      credentials: 'same-origin',
      headers: Object.assign({ 'Content-Type': 'application/json' }, WW_APP_HEADER),
      body: JSON.stringify({ method: method, id: '1', params: [params], jsonrpc: '2.0' }),
    }).then(function (r) {
      if (!r.ok) throw new Error(method + ' failed: HTTP ' + r.status);
      return r.json();
    }).then(function (j) {
      if (j.error) {
        console.error('[word-digital] ' + method + ' error response:', j.error);
        var e = j.error;
        var parts = [];
        if (e.message) parts.push(e.message);
        if (e.data && e.data.detail && e.data.detail !== e.message) parts.push(e.data.detail);
        if (e.code) parts.push('(code ' + e.code + ')');
        throw new Error(method + ' failed: ' + (parts.join(' — ') || JSON.stringify(e)));
      }
      // Some services report per-object failures in Reports with an otherwise
      // successful envelope — treat those as errors too.
      if (j.result && j.result.Reports && j.result.Reports.length &&
          (!j.result.Objects || !j.result.Objects.length)) {
        console.error('[word-digital] ' + method + ' reports:', j.result.Reports);
        var msgs = j.result.Reports.map(function (rep) {
          return (rep.Entries || []).map(function (en) { return en.Message || ''; }).join(' ') || rep.BelongsTo && rep.BelongsTo.Id || '';
        }).filter(Boolean);
        throw new Error(method + ' failed: ' + (msgs.join(' | ') || 'server returned error reports'));
      }
      return j.result;
    });
  }

  function guid() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      var r = Math.random() * 16 | 0;
      return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
    });
  }

  // Upload to the Transfer Server the way Studio itself does: client-side
  // fileguid, PUT to transferindex.php, and the PUT URL doubles as the
  // Attachment FileUrl in CreateObjects.
  function uploadToTransferServer(content) {
    var url = transferUrl() + '?fileguid=' + guid() + '&ww-app=' + encodeURIComponent('Content+Station');
    var ticket = getTicket();
    if (ticket) url += '&ticket=' + encodeURIComponent(ticket);
    url += '&format=' + encodeURIComponent(DIGITAL_MIME);
    return fetch(url, {
      method: 'PUT',
      credentials: 'same-origin',
      headers: Object.assign({ 'Content-Type': DIGITAL_MIME }, WW_APP_HEADER),
      body: content,
    }).then(function (r) {
      if (!r.ok) throw new Error('File upload to Transfer Server failed: HTTP ' + r.status);
      return url;
    });
  }

  // Object names may not contain the characters Enterprise rejects (/ \ : * ? " < > |).
  // Only the Studio object name is sanitised — headlines inside the article keep them.
  function sanitizeObjectName(name) {
    return String(name)
      .replace(/[\/\\:*?"<>|]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 60)
      .trim();
  }

  function extraMeta(property, values) {
    return { __classname__: 'ExtraMetaData', Property: property, Values: values };
  }

  // Create the digital article inside the given dossier.
  // Publication/Category are taken from the dossier; Targets are copied from
  // the dossier so the article lands on the same channel/issue. Component set,
  // Look and Feel and Twixl collection come from BRAND_DEFAULTS.
  function createArticleInDossier(digital, name, feedHeadline, dossier) {
    var digitalJson = JSON.stringify(digital);
    // Unique component identifiers in order of first use (C_CS_DE_COMPONENT_NAMES)
    var componentNames = [];
    (digital.data.content || []).forEach(function (comp) {
      if (comp.identifier && componentNames.indexOf(comp.identifier) === -1) componentNames.push(comp.identifier);
    });
    var pubId = String(dossier.PublicationId || (dossier.Publication && dossier.Publication.Id) || '');
    var catId = String(dossier.CategoryId || (dossier.Category && dossier.Category.Id) || '');
    var dossierId = String(dossier.ID || dossier.Id);

    var dossierTargets = [];
    return callServer('GetObjects', {
      IDs: [dossierId], Lock: false, Rendition: 'none',
      RequestInfo: ['Targets', 'MetaData'], HaveVersions: null, Areas: null, EditionId: null,
    }).then(function (res) {
      var obj = res.Objects && res.Objects[0];
      if (obj) {
        dossierTargets = obj.Targets || [];
        var bm = obj.MetaData && obj.MetaData.BasicMetaData;
        if (bm) {
          pubId = pubId || String((bm.Publication && bm.Publication.Id) || '');
          catId = catId || String((bm.Category && bm.Category.Id) || '');
        }
      }
      return callServer('GetStates', {
        ID: null,
        Publication: { Id: pubId, __classname__: 'Publication' },
        Issue: null,
        Section: catId ? { Id: catId, __classname__: 'Category' } : null,
        Type: 'Article',
      });
    }).then(function (res) {
      var states = (res && res.States) || [];
      if (!states.length) throw new Error('No workflow statuses available for Articles in this Brand/Category');
      var state = states[0];
      return uploadToTransferServer(digitalJson).then(function (fileUrl) {
        return callServer('CreateObjects', {
          Lock: false, Autonaming: true,
          Objects: [{
            __classname__: 'Object',
            MetaData: {
              __classname__: 'MetaData',
              BasicMetaData: {
                __classname__: 'BasicMetaData',
                ID: null, DocumentID: null,
                Name: sanitizeObjectName(name),
                Type: 'Article',
                Publication: { Id: pubId, __classname__: 'Publication' },
                Category: { Id: catId, __classname__: 'Category' },
                ContentSource: null,
              },
              RightsMetaData: null,
              SourceMetaData: null,
              ContentMetaData: {
                __classname__: 'ContentMetaData',
                Format: DIGITAL_MIME,
              },
              WorkflowMetaData: {
                __classname__: 'WorkflowMetaData',
                State: { Id: state.Id, __classname__: 'State' },
              },
              ExtraMetaData: (function () {
                var extra = [
                  extraMeta('C_CS_FILEFORMATVERSION', [digital.version || '2.4']),
                  extraMeta('C_CS_DE_COMPONENT_NAMES', [componentNames.join(',')]),
                ];
                if (feedHeadline) extra.push(extraMeta('C_HEADLINE', [feedHeadline]));
                var bd = BRAND_DEFAULTS[pubId];
                if (bd) {
                  extra.push(extraMeta('C_CS_COMPONENTSET', [bd.componentSet]));
                  extra.push(extraMeta('C_CS_STYLEID', [bd.lookAndFeel]));
                  if (bd.twixlCollectionId) extra.push(extraMeta('C_TW_COLLECTION_ID', [bd.twixlCollectionId]));
                }
                return extra;
              })(),
            },
            Relations: [{
              __classname__: 'Relation',
              Parent: dossierId, Child: null, Type: 'Contained',
              Placements: null, ParentVersion: null, ChildVersion: null,
              Geometry: null, Rating: null, Targets: null,
            }],
            Pages: null,
            Files: [{
              __classname__: 'Attachment',
              Rendition: 'native',
              Type: DIGITAL_MIME,
              Content: null, FilePath: null,
              FileUrl: fileUrl,
              EditionId: null, ContentSourceFileLink: null, ContentSourceProxyLink: null,
            }],
            Messages: null, Elements: null,
            Targets: dossierTargets,
            Renditions: null, MessageList: null, ObjectLabels: null, Operations: null,
          }],
        });
      });
    });
  }

  // ─── Shared converter UI ───────────────────────────────────────────────────
  var CSS = [
    '.wdab-scroll{max-height:calc(100vh - 140px);overflow-y:auto;-webkit-overflow-scrolling:touch}',
    '.wdab{max-width:640px;margin:0 auto;padding:24px 16px 48px;font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;color:#1e293b}',
    '.wdab-modal{max-height:min(600px,calc(100vh - 280px));overflow-y:auto;-webkit-overflow-scrolling:touch}',
    '.wdab-modal .wdab{padding:4px 2px 8px;max-width:none}',
    '.wdab h2{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;color:#64748b;margin:0 0 14px}',
    '.wdab .wdab-card{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:20px;margin-bottom:16px;box-shadow:0 1px 2px rgba(0,0,0,.04)}',
    '.wdab-modal .wdab .wdab-card{border:0;box-shadow:none;padding:8px 0;margin-bottom:4px}',
    '.wdab label{display:block;font-weight:500;color:#334155;margin:0 0 4px}',
    '.wdab select,.wdab input[type=text]{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:8px;padding:7px 10px;font:inherit;color:#1e293b;background:#fff}',
    '.wdab input[type=file]{width:100%;font:inherit}',
    '.wdab .wdab-row{margin-bottom:14px}',
    '.wdab button.wdab-btn{display:inline-block;border:0;border-radius:8px;padding:9px 16px;font:inherit;font-weight:600;cursor:pointer;background:#2563eb;color:#fff;width:100%}',
    '.wdab button.wdab-btn:disabled{opacity:.4;cursor:not-allowed}',
    '.wdab .wdab-feed input{border-color:#fcd34d;background:#fffbeb}',
    '.wdab .wdab-feed-note{display:inline-block;font-size:11px;color:#b45309;background:#fffbeb;border:1px solid #fde68a;border-radius:4px;padding:1px 6px;margin-left:8px;font-weight:400}',
    '.wdab .wdab-error{color:#dc2626;margin-top:8px;display:none;white-space:pre-wrap}',
    '.wdab .wdab-warn{background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:10px 12px;margin-bottom:14px;display:none}',
    '.wdab .wdab-warn ul{margin:6px 0 0;padding-left:18px;color:#92400e}',
    '.wdab .wdab-entries{max-height:220px;overflow-y:auto;border:1px solid #e2e8f0;border-radius:8px;padding:8px 12px;color:#334155}',
    '.wdab .wdab-entries .n{color:#94a3b8;margin-right:6px}',
    '.wdab .wdab-hidden{display:none}',
    '.wdab .wdab-note{color:#64748b;font-size:12px;margin-top:8px}'
  ].join('\n');

  var cssInjected = false;
  function injectCss() {
    if (cssInjected) return;
    var style = document.createElement('style');
    style.textContent = CSS;
    document.head.appendChild(style);
    cssInjected = true;
  }

  function esc(str) {
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function formHtml(p, actionLabel) {
    return '<div class="wdab">' +
      '  <div class="wdab-card">' +
      '    <h2>1 — Set up</h2>' +
      '    <div class="wdab-row">' +
      '      <label for="' + p + '-type">Article type</label>' +
      '      <select id="' + p + '-type">' +
      '        <option value="countdown">Type 1 — Numbered countdown (50 → 1)</option>' +
      '        <option value="ascending">Type 2 — Numbered ascending (1 → 50)</option>' +
      '        <option value="crosshead">Type 3 — Crosshead / generic article</option>' +
      '      </select>' +
      '    </div>' +
      '    <div class="wdab-row">' +
      '      <label for="' + p + '-source">Source</label>' +
      '      <select id="' + p + '-source">' +
      '        <option value="docx">Word document (.docx)</option>' +
      '        <option value="url">TopGear article URL</option>' +
      '      </select>' +
      '    </div>' +
      '    <div class="wdab-row" id="' + p + '-docx-row">' +
      '      <label for="' + p + '-file">Word document (.docx)</label>' +
      '      <input type="file" id="' + p + '-file" accept=".docx">' +
      '    </div>' +
      '    <div class="wdab-row wdab-hidden" id="' + p + '-url-row">' +
      '      <label for="' + p + '-url">TopGear article URL</label>' +
      '      <input type="url" id="' + p + '-url" placeholder="https://www.topgear.com/…">' +
      '    </div>' +
      '    <button class="wdab-btn" id="' + p + '-parse" disabled>Parse Document</button>' +
      '    <p class="wdab-error" id="' + p + '-parse-error"></p>' +
      '  </div>' +
      '  <div id="' + p + '-preview" class="wdab-hidden">' +
      '    <div class="wdab-card">' +
      '      <h2>2 — Detected metadata (editable)</h2>' +
      '      <div class="wdab-row wdab-feed wdab-hidden" id="' + p + '-feed-row">' +
      '        <label>Feed Headline<span class="wdab-feed-note" id="' + p + '-feed-note">Saved to C_HEADLINE</span></label>' +
      '        <input type="text" id="' + p + '-feed">' +
      '      </div>' +
      '      <div class="wdab-row"><label>Article title</label><input type="text" id="' + p + '-title"></div>' +
      '      <div class="wdab-row"><label>Subtitle</label><input type="text" id="' + p + '-subtitle"></div>' +
      '      <div class="wdab-row"><label>Author name</label><input type="text" id="' + p + '-author"></div>' +
      '      <div class="wdab-warn" id="' + p + '-warn"><strong>Flagged for review — kept in the article:</strong> these look like editor instructions rather than copy. Each stays in place as plain body text; delete any that shouldn\'t ship.<ul id="' + p + '-warn-list"></ul></div>' +
      '    </div>' +
      '    <div class="wdab-card">' +
      '      <h2>3 — Entries (<span id="' + p + '-count">0</span>)</h2>' +
      '      <div class="wdab-entries" id="' + p + '-entries"></div>' +
      '    </div>' +
      (actionLabel
        ? '<div class="wdab-card"><h2>4 — ' + esc(actionLabel) + '</h2>' +
          '<button class="wdab-btn" id="' + p + '-action">' + esc(actionLabel) + '</button>' +
          '<p class="wdab-error" id="' + p + '-action-error"></p>' +
          '<p class="wdab-note" id="' + p + '-note"></p></div>'
        : '') +
      '  </div>' +
      '</div>';
  }

  // Wires the converter form; returns a controller for reading the result.
  function wireForm(p, onAction) {
    var $ = function (id) { return document.getElementById(p + '-' + id); };
    var state = { parsedData: null, uploadedFilename: '' };

    $('file').addEventListener('change', function () {
      $('parse').disabled = !$('file').files.length;
      $('preview').classList.add('wdab-hidden');
      state.parsedData = null;
      $('parse-error').style.display = 'none';
    });

    $('source').addEventListener('change', function () {
      var isDocx = $('source').value === 'docx';
      $('docx-row').classList.toggle('wdab-hidden', !isDocx);
      $('url-row').classList.toggle('wdab-hidden', isDocx);
      refreshParse();
    });
    $('url').addEventListener('input', refreshParse);

    function refreshParse() {
      var isDocx = $('source').value === 'docx';
      $('parse').disabled = isDocx ? !$('file').files.length : !$('url').value.trim();
      $('parse').textContent = isDocx ? 'Parse Document' : 'Fetch & Parse Article';
    }

    $('parse').addEventListener('click', function () {
      var source = $('source').value;
      var file = $('file').files[0];
      if (source === 'docx' && !file) return;
      if (source === 'url' && !$('url').value.trim()) return;
      var type = $('type').value;
      $('parse').disabled = true;
      $('parse').textContent = source === 'docx' ? 'Parsing…' : 'Fetching…';
      $('parse-error').style.display = 'none';

      var pipeline;
      if (source === 'url') {
        var articleUrl = $('url').value.trim();
        pipeline = parseFromUrl(articleUrl, type).then(function (r) {
          state.imageUrls = r.imageUrls || [];
          state.uploadedFilename = slugFromUrl(articleUrl);
          return { meta: r.meta, entries: r.entries };
        });
      } else {
        pipeline = loadMammoth()
          .then(function (mammoth) { return file.arrayBuffer().then(function (buf) { return mammoth.convertToHtml({ arrayBuffer: buf }); }); })
          .then(function (result) {
            state.imageUrls = [];
            state.uploadedFilename = file.name.replace(/\.docx$/i, '');
            return type === 'crosshead' ? parseCrosshead(result.value) : parseNumbered(result.value, type);
          });
      }

      pipeline
        .then(function (parsed) {
          state.parsedData = { meta: parsed.meta, entries: parsed.entries, type: type };

          $('feed').value = parsed.meta.feedHeadline;
          $('feed-row').classList.toggle('wdab-hidden', !parsed.meta.feedHeadline);
          $('title').value = parsed.meta.title;
          $('subtitle').value = parsed.meta.subtitle;
          $('author').value = parsed.meta.author;

          if (parsed.meta.flagged && parsed.meta.flagged.length) {
            $('warn-list').innerHTML = parsed.meta.flagged.map(function (t) { return '<li>' + esc(t) + '</li>'; }).join('');
            $('warn').style.display = 'block';
          } else {
            $('warn').style.display = 'none';
          }

          $('count').textContent = parsed.entries.length;
          $('entries').innerHTML = parsed.entries.slice(0, 60).map(function (e, i) {
            return type === 'crosshead'
              ? '<div><span class="n">' + (i + 1) + '.</span>' + esc(e.crosshead || '(no crosshead)') + '</div>'
              : '<div><span class="n">[' + e.number + ']</span>' + esc(e.name) + '</div>';
          }).join('') + (parsed.entries.length > 60 ? '<div class="n">… and ' + (parsed.entries.length - 60) + ' more</div>' : '');

          $('preview').classList.remove('wdab-hidden');
        })
        .catch(function (err) {
          $('parse-error').textContent = 'Error parsing document: ' + err.message;
          $('parse-error').style.display = 'block';
        })
        .then(function () {
          $('parse').disabled = false;
          refreshParse();
        });
    });

    if (onAction && $('action')) {
      $('action').addEventListener('click', function () { onAction(controller); });
    }

    var controller = {
      $: $,
      getResult: function () {
        if (!state.parsedData) return null;
        var pm = state.parsedData.meta;
        var pd = pm.deltas || {};
        var meta = {
          feedHeadline: $('feed').value,
          title: $('title').value,
          subtitle: $('subtitle').value,
          author: $('author').value,
          score: pm.score,
          intro: pm.intro,
          introParts: pm.introParts,
          leadParts: pm.leadParts,
          titleDeltas: $('title').value === pm.title ? pd.title : null,
          subtitleDeltas: $('subtitle').value === pm.subtitle ? pd.subtitle : null,
        };
        var template = deepClone(TEMPLATES[state.parsedData.type]);
        var digital = state.parsedData.type === 'crosshead'
          ? buildCrosshead(template, meta, state.parsedData.entries)
          : buildNumbered(template, meta, state.parsedData.entries, state.parsedData.type);
        return { digital: digital, meta: meta, filename: state.uploadedFilename };
      },
    };
    return controller;
  }

  // ─── Dossier toolbar: convert and create directly in this Dossier ─────────
  // (The Apps-menu custom app was removed on request — the Dossier button is
  // the only trigger. The standalone web version on GitHub Pages still offers
  // a .digital download when one is needed.)
  ContentStationSdk.addDossierToolbarButton({
    label: 'Word & Web → Digital',
    onAction: function (config, selection, dossier) {
      injectCss();
      var dialogId = null;
      var busy = false;

      var content = '<div class="wdab-modal">' + formHtml('wdabm', 'Create Digital Article in this Dossier') + '</div>';

      dialogId = ContentStationSdk.openModalDialog({
        title: 'Word & Web → Digital',
        subtitle: 'Creates the digital article in Dossier “' + esc(dossier.Name || '') + '”',
        content: content,
        width: 640,
        buttons: [{ label: 'Close', class: 'pale' }],
      });

      wireForm('wdabm', function (ctl) {
        if (busy) return;
        var result = ctl.getResult();
        if (!result) return;
        busy = true;
        var btn = ctl.$('action');
        var errEl = ctl.$('action-error');
        var noteEl = ctl.$('note');
        btn.disabled = true;
        btn.textContent = 'Creating…';
        errEl.style.display = 'none';

        var name = result.meta.title || result.filename;
        createArticleInDossier(result.digital, name, result.meta.feedHeadline, dossier)
          .then(function (res) {
            var created = res && res.Objects && res.Objects[0];
            var newName = created && created.MetaData && created.MetaData.BasicMetaData
              ? created.MetaData.BasicMetaData.Name : name;
            ContentStationSdk.showNotification({
              content: 'Digital article “' + esc(newName) + '” created in Dossier “' + esc(dossier.Name || '') + '”.',
              icon: 'check',
            });
            try { ContentStationSdk.refreshCurrentSearch(); } catch (e) { /* non-fatal */ }
            if (dialogId !== null) ContentStationSdk.closeModalDialog(dialogId);
          })
          .catch(function (err) {
            errEl.textContent = err.message + '\nNothing was created. You can fix the issue and try again.';
            errEl.style.display = 'block';
          })
          .then(function () {
            busy = false;
            btn.disabled = false;
            btn.textContent = 'Create Digital Article in this Dossier';
          });

        noteEl.textContent = result.meta.feedHeadline ? 'Feed headline will be saved to C_HEADLINE.' : '';
      });
    },
  });
})();
