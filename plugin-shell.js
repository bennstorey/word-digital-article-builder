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
      // Article templates read from Studio when an article is built, so an
      // edit to the template in Studio reaches the next article. All four in
      // dossier "2024-tg-templates" (29570). The built-in copies in TEMPLATES
      // are used only if Studio can't be read.
      templates: {
        countdown: '66075',  // Ad hoc Apple-features-long numbered list 50-1
        ascending: '67380',  // Ad hoc Apple-features-long numbered list 1-50
        snippet: '67380',    // Type 4 reuses the ascending layout
        crosshead: '66074',  // Ad hoc Apple-features-generic-crossheads
      },
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
  function uploadToTransferServer(content, mime) {
    var type = mime || DIGITAL_MIME;
    var url = transferUrl() + '?fileguid=' + guid() + '&ww-app=' + encodeURIComponent('Content+Station');
    var ticket = getTicket();
    if (ticket) url += '&ticket=' + encodeURIComponent(ticket);
    url += '&format=' + encodeURIComponent(type);
    return fetch(url, {
      method: 'PUT',
      credentials: 'same-origin',
      headers: Object.assign({ 'Content-Type': type }, WW_APP_HEADER),
      body: content,
    }).then(function (r) {
      if (!r.ok) throw new Error('File upload to Transfer Server failed: HTTP ' + r.status);
      return url;
    });
  }

  // Studio object names are made web-safe: letters, digits, spaces, "-", "_"
  // and "." only. Studio rejects more than the documented / \ : * ? " < > | —
  // "These are the UK's top…" failed with S1026 at every length (2026-10-06),
  // so this is an allow-list, not a block-list. Only the object's name is
  // changed; the headline inside the article keeps its punctuation.
  //   UK's → UKs · café → cafe · £5,000 → 5000 · R&D → R and D · 50–1 → 50-1
  // The server's name-length limit varies by install (60 has been seen to
  // fail), so the name is cut on a word boundary and retried shorter.
  function webSafeName(name) {
    var out = String(name == null ? '' : name).replace(/\u2026|\.{2,}/g, ' '); // … and ... → a space
    if (out.normalize) out = out.normalize('NFKD').replace(/[\u0300-\u036f]/g, ''); // é → e
    return out
      .replace(/[\u2018\u2019\u02bc\u0060\u00b4']/g, '')    // apostrophes: UK's → UKs
      .replace(/[\u201c\u201d"]/g, '')                        // quote marks
      .replace(/[\u2013\u2014\u2212]/g, '-')                  // en/em dash, minus
      .replace(/&/g, ' and ')
      .replace(/(\d),(?=\d{3}\b)/g, '$1')                     // 5,000 → 5000
      .replace(/[^A-Za-z0-9 ._-]+/g, ' ')                      // everything else
      .replace(/\s+/g, ' ')
      .replace(/^[ ._-]+|[ ._-]+$/g, '');                      // no leading/trailing punctuation
  }

  function sanitizeObjectName(name, maxLen) {
    var out = webSafeName(name);
    var limit = maxLen || NAME_LIMITS[0];
    if (out.length > limit) {
      out = out.slice(0, limit);
      var cut = out.lastIndexOf(' ');
      if (cut > limit * 0.5) out = out.slice(0, cut);
    }
    return out.replace(/[ ._-]+$/, '') || 'Untitled article';
  }

  // Tried in order when the server rejects a name as too long (S1026).
  var NAME_LIMITS = [60, 40, 25];

  function extraMeta(property, values) {
    return { __classname__: 'ExtraMetaData', Property: property, Values: values };
  }

  // Brand/Category/Targets for a dossier — shared by the article and image
  // creation paths so both land in the same place with the same targets.
  function resolveDossierContext(dossier) {
    var pubId = String(dossier.PublicationId || (dossier.Publication && dossier.Publication.Id) || '');
    var catId = String(dossier.CategoryId || (dossier.Category && dossier.Category.Id) || '');
    var dossierId = String(dossier.ID || dossier.Id);
    return callServer('GetObjects', {
      IDs: [dossierId], Lock: false, Rendition: 'none',
      RequestInfo: ['Targets', 'MetaData'], HaveVersions: null, Areas: null, EditionId: null,
    }).then(function (res) {
      var obj = res.Objects && res.Objects[0];
      var targets = [];
      var pubName = (dossier.Publication && dossier.Publication.Name) || '';
      if (obj) {
        targets = obj.Targets || [];
        var bm = obj.MetaData && obj.MetaData.BasicMetaData;
        if (bm) {
          pubId = pubId || String((bm.Publication && bm.Publication.Id) || '');
          catId = catId || String((bm.Category && bm.Category.Id) || '');
          pubName = pubName || (bm.Publication && bm.Publication.Name) || '';
        }
      }
      return { pubId: pubId, pubName: pubName, catId: catId, dossierId: dossierId, targets: targets };
    });
  }

  // ─── Article templates from Studio ────────────────────────────────────────
  // The template's native file, read under the editor's session like any
  // other download. Studio's own template metadata (name, version, component
  // set, Look and Feel, Twixl collection) goes onto the created article.
  // Never fatal: if Studio can't be read the built-in copy is used and the
  // dialog says so.
  function fetchTransferFile(fileUrl) {
    var url = fileUrl;
    var ticket = getTicket();
    if (ticket && !/[?&]ticket=/.test(url)) url += (url.indexOf('?') === -1 ? '?' : '&') + 'ticket=' + encodeURIComponent(ticket);
    return fetch(url, { credentials: 'same-origin', headers: WW_APP_HEADER }).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.text();
    });
  }

  // A template the builders can use: the components they fill must be there
  function checkTemplate(t, type) {
    var content = t && t.data && t.data.content;
    if (!Array.isArray(content) || !content.length) throw new Error('not a digital article template');
    var has = function (id) { return content.some(function (c) { return c.identifier === id; }); };
    if (!has('header-image') || !has('title')) throw new Error('it has no header picture or title');
    if (type === 'crosshead' ? !has('crosshead') : content.filter(function (c) { return c.identifier === 'image'; }).length < 2) {
      throw new Error('it does not have the ' + (type === 'crosshead' ? 'crossheads' : 'numbered entries') + ' this article template needs');
    }
  }

  function loadTemplate(type, dossier) {
    var builtIn = function (info) { return { template: TEMPLATES[type], info: Object.assign({ source: 'built-in' }, info) }; };
    return resolveDossierContext(dossier).then(function (ctx) {
      var bd = BRAND_DEFAULTS[ctx.pubId];
      var id = bd && bd.templates && bd.templates[type];
      if (!id) return builtIn({ error: 'no Studio template set up for this brand' });
      return callServer('GetObjects', {
        IDs: [id], Lock: false, Rendition: 'native',
        RequestInfo: ['MetaData'], HaveVersions: null, Areas: null, EditionId: null,
      }).then(function (res) {
        var obj = res.Objects && res.Objects[0];
        var file = obj && (obj.Files || []).filter(function (f) { return f.FileUrl; })[0];
        if (!file) throw new Error('Studio sent no file');
        return fetchTransferFile(file.FileUrl).then(function (text) {
          var t = JSON.parse(text);
          checkTemplate(t, type);
          var md = obj.MetaData;
          var ex = {};
          (md.ExtraMetaData || []).forEach(function (e) { ex[e.Property] = (e.Values || [])[0] || ''; });
          return {
            template: useSharedFurniture(t),
            info: {
              source: 'studio', id: String(id),
              name: md.BasicMetaData.Name,
              version: md.WorkflowMetaData && md.WorkflowMetaData.Version,
              templateName: ex.C_ARTICLE_TEMPLATE_NAME || md.BasicMetaData.Name,
              componentSet: ex.C_CS_COMPONENTSET, lookAndFeel: ex.C_CS_STYLEID, twixlCollectionId: ex.C_TW_COLLECTION_ID,
            },
          };
        });
      }).catch(function (e) {
        console.warn('[word-digital] Studio template ' + id + ' not used:', e.message);
        return builtIn({ id: String(id), error: e.message });
      });
    });
  }

  // Create the digital article inside the given dossier.
  // Publication/Category are taken from the dossier; Targets are copied from
  // the dossier so the article lands on the same channel/issue. Component set,
  // Look and Feel and Twixl collection come from the Studio template it was
  // built from (tplInfo, from loadTemplate), else from BRAND_DEFAULTS; the
  // article is linked to that template as Studio's own "new from template" does.
  function createArticleInDossier(digital, name, feedHeadline, dossier, nameLimit, tplInfo) {
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
    return resolveDossierContext(dossier).then(function (ctx) {
      dossierTargets = ctx.targets;
      pubId = ctx.pubId; catId = ctx.catId;
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
                Name: sanitizeObjectName(name, nameLimit),
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
                var bd = BRAND_DEFAULTS[pubId] || {};
                var st = tplInfo && tplInfo.source === 'studio' ? tplInfo : {};
                var componentSet = st.componentSet || bd.componentSet;
                var lookAndFeel = st.lookAndFeel || bd.lookAndFeel;
                var twixl = st.twixlCollectionId || bd.twixlCollectionId;
                if (componentSet) extra.push(extraMeta('C_CS_COMPONENTSET', [componentSet]));
                if (lookAndFeel) extra.push(extraMeta('C_CS_STYLEID', [lookAndFeel]));
                if (twixl) extra.push(extraMeta('C_TW_COLLECTION_ID', [twixl]));
                if (st.id) {
                  extra.push(extraMeta('C_CS_ARTICLE_TEMPLATE_ID', [st.id]));
                  extra.push(extraMeta('C_ARTICLE_TEMPLATE_NAME', [st.templateName]));
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

  // ─── Images into the dossier ──────────────────────────────────────────────
  // Article images are fetched through the proxy (topgear.com sends no CORS
  // headers), uploaded to the Transfer Server exactly like the .digital file,
  // then created as Image objects contained in the same dossier.

  var IMAGE_MIME_BY_EXT = {
    jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png',
    gif: 'image/gif', webp: 'image/webp', avif: 'image/avif',
  };

  function imageNameFromUrl(u) {
    var base = '';
    try { base = decodeURIComponent(new URL(u).pathname.split('/').pop() || ''); }
    catch (e) { base = String(u).split('/').pop() || ''; }
    // Drupal sometimes bakes a path parameter into the stored filename, e.g.
    // "_V2A0009V2.jpg;jsessionid=null_1.jpg". Studio rejects the ';' / '=' in an
    // object name (S1026), so keep only what precedes the first ';'.
    base = base.split(';')[0];
    return base || 'image';
  }

  // Last-resort name when Studio still rejects the cleaned one: letters, digits,
  // spaces, '-' and '_' only, falling back to a positional name.
  function safeImageName(u, index) {
    var n = imageNameFromUrl(u).replace(/\.[a-z0-9]+$/i, '')
      .replace(/[^A-Za-z0-9 _-]+/g, '_').replace(/_+/g, '_').replace(/^[_ ]+|[_ ]+$/g, '')
      .slice(0, 40);
    return n || ('topgear-image-' + (index + 1));
  }

  function imageMimeFromUrl(u, blobType) {
    if (blobType && blobType.indexOf('image/') === 0) return blobType;
    var ext = (imageNameFromUrl(u).split('.').pop() || '').toLowerCase();
    return IMAGE_MIME_BY_EXT[ext] || 'image/jpeg';
  }

  // Creates one Image object. Returns the created object's Id.
  function createImageObject(blob, url, ctx, state, nameOverride) {
    var mime = imageMimeFromUrl(url, blob.type);
    var name = nameOverride || imageNameFromUrl(url).replace(/\.[a-z0-9]+$/i, '');
    return uploadToTransferServer(blob, mime).then(function (fileUrl) {
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
              Type: 'Image',
              Publication: { Id: ctx.pubId, __classname__: 'Publication' },
              Category: { Id: ctx.catId, __classname__: 'Category' },
              ContentSource: null,
            },
            RightsMetaData: null, SourceMetaData: null,
            ContentMetaData: { __classname__: 'ContentMetaData', Format: mime },
            WorkflowMetaData: {
              __classname__: 'WorkflowMetaData',
              State: { Id: state.Id, __classname__: 'State' },
              Comment: null, Version: null, Modifier: null, Modified: null,
              Creator: null, Created: null, Deletor: null, Deleted: null,
              Routing: null, LockedBy: null,
            },
            ExtraMetaData: [],
          },
          Relations: [{
            __classname__: 'Relation',
            Parent: ctx.dossierId, Child: null, Type: 'Contained',
            Placements: null, ParentVersion: null, ChildVersion: null,
            Geometry: null, Rating: null, Targets: null,
          }],
          Pages: null,
          Files: [{
            __classname__: 'Attachment',
            Rendition: 'native', Type: mime,
            Content: null, FilePath: null, FileUrl: fileUrl,
            EditionId: null, ContentSourceFileLink: null, ContentSourceProxyLink: null,
          }],
          Messages: null, Elements: null,
          Targets: ctx.targets,
          Renditions: null, MessageList: null, ObjectLabels: null, Operations: null,
        }],
      });
    }).then(function (res) {
      var created = res && res.Objects && res.Objects[0];
      var id = created && created.MetaData && created.MetaData.BasicMetaData
        ? created.MetaData.BasicMetaData.ID : null;
      return { id: id, name: name, url: url };
    });
  }

  function fetchViaProxy(u) {
    return fetch(proxyUrl(u), { credentials: 'omit' }).then(function (r) {
      if (!r.ok) throw new Error('fetch failed: HTTP ' + r.status);
      return r.blob();
    });
  }

  // Uploads sequentially so a long gallery can't swamp the server, and so a
  // single failure is reported against its own image rather than aborting all.
  // fetchImage(url) → Promise<Blob>; defaults to the topgear.com proxy, the
  // WhatsApp source passes one that reads from the receiver.
  function createImagesInDossier(urls, dossier, onProgress, fetchImage) {
    fetchImage = fetchImage || fetchViaProxy;
    if (!urls || !urls.length) return Promise.resolve({ created: [], failed: [] });
    return resolveDossierContext(dossier).then(function (ctx) {
      return callServer('GetStates', {
        ID: null,
        Publication: { Id: ctx.pubId, __classname__: 'Publication' },
        Issue: null,
        Section: ctx.catId ? { Id: ctx.catId, __classname__: 'Category' } : null,
        Type: 'Image',
      }).then(function (res) {
        var states = (res && res.States) || [];
        if (!states.length) {
          throw new Error('No workflow statuses available for Images in this Brand/Category — ' +
                          'the article was created, but images could not be added.');
        }
        var state = states[0];
        var created = [], failed = [];
        var chain = Promise.resolve();
        urls.forEach(function (u, i) {
          chain = chain.then(function () {
            if (onProgress) onProgress(i, urls.length);
            return fetchImage(u)
              .then(function (blob) {
                return createImageObject(blob, u, ctx, state).catch(function (e) {
                  if (!/S1026|invalid characters|too long/i.test(e.message || '')) throw e;
                  return createImageObject(blob, u, ctx, state, safeImageName(u, i));
                });
              })
              .then(function (info) { created.push(info); })
              .catch(function (e) { failed.push({ url: u, error: e.message }); });
          });
        });
        return chain.then(function () {
          if (onProgress) onProgress(urls.length, urls.length);
          return { created: created, failed: failed };
        });
      });
    });
  }

  // How a Digital Editor image component references a Studio Image object.
  // Mirrors the shape the templates already use for the apple-news-follow
  // component's image. If Studio expects something different, this is the only
  // place that needs changing.
  function imageRef(objectId) {
    return { id: String(objectId), focuspoint: { x: 0.5, y: 0.5 }, cropper: false };
  }

  // Fills the article's image slots, in document order, from the created Image
  // objects. imageUrls order is hero first then the gallery, and the template's
  // slots run header-image then one per entry, so index order lines up.
  // apple-news-follow is untouched — it carries its own branded image.
  function applyImageIds(digital, ids) {
    var slots = (digital.data.content || []).filter(function (c) {
      return c.identifier === 'image' || c.identifier === 'header-image';
    });
    if (!ids || !ids.length) return { digital: digital, filled: 0, slots: slots.length };
    var filled = 0;
    for (var i = 0; i < slots.length && i < ids.length; i++) {
      if (!ids[i]) continue;
      slots[i].content = Object.assign({}, slots[i].content || {}, { image: imageRef(ids[i]) });
      filled++;
    }
    return { digital: digital, filled: filled, slots: slots.length };
  }

  // Studio ingests an uploaded image asynchronously — the object exists as soon
  // as CreateObjects returns, but its renditions are generated afterwards. If the
  // article is created and opened before that finishes, the placements look
  // empty. Wait until every new image reports a rendition before continuing.
  function waitForImagesReady(ids, onProgress, timeoutMs) {
    if (!ids || !ids.length) return Promise.resolve({ ready: [], pending: [] });
    var deadline = Date.now() + (timeoutMs || 60000);
    var remaining = ids.slice();
    var ready = [];

    function poll() {
      return callServer('GetObjects', {
        IDs: remaining, Lock: false, Rendition: 'thumb',
        RequestInfo: ['MetaData'], HaveVersions: null, Areas: null, EditionId: null,
      }).then(function (res) {
        var objs = (res && res.Objects) || [];
        var stillPending = [];
        objs.forEach(function (o) {
          var bm = (o.MetaData && o.MetaData.BasicMetaData) || {};
          var hasFile = !!(o.Files && o.Files.length);
          if (hasFile) ready.push(String(bm.ID));
          else stillPending.push(String(bm.ID));
        });
        // Ids the server did not return at all are still settling
        var returned = objs.map(function (o) {
          return String(((o.MetaData || {}).BasicMetaData || {}).ID);
        });
        remaining.forEach(function (id) {
          if (returned.indexOf(String(id)) === -1) stillPending.push(String(id));
        });
        remaining = stillPending;

        if (onProgress) onProgress(ready.length, ids.length);
        if (!remaining.length) return { ready: ready, pending: [] };
        if (Date.now() > deadline) return { ready: ready, pending: remaining };
        return new Promise(function (r) { setTimeout(r, 1500); }).then(poll);
      }).catch(function () {
        // A failed poll should never block article creation
        return { ready: ready, pending: remaining };
      });
    }
    return poll();
  }

  // ─── Shared furniture images ──────────────────────────────────────────────
  // The Follow and Newsletter blocks use shared assets that live in the
  // "general-furniture" dossier (52254), not per-article copies. Studio makes
  // them show up in an article's dossier by adding a Contained relation from
  // that dossier to the same object — the object ID never changes. So we
  // reference the shared IDs in the template and link them into the dossier.
  //
  //   91357 tg-logo-tech-blue4x-88                    Follow, light
  //   91358 tg-logo-white4x-100                       Follow, dark
  //   91356 tg-follow-newsletter-signup-light-wide-88 Newsletter, light
  //   91355 tg-follow-newsletter-signup-dark-wide-88  Newsletter, dark
  var FURNITURE_IMAGE_IDS = ['91357', '91358', '91356', '91355'];

  function linkFurnitureToDossier(dossierId) {
    var relations = FURNITURE_IMAGE_IDS.map(function (id) {
      return {
        __classname__: 'Relation',
        Parent: String(dossierId), Child: String(id), Type: 'Contained',
        Placements: null, ParentVersion: null, ChildVersion: null,
        Geometry: null, Rating: null, Targets: null,
      };
    });
    return callServer('CreateObjectRelations', { Relations: relations })
      .then(function () { return { linked: FURNITURE_IMAGE_IDS.length, error: null }; })
      .catch(function (e) {
        // Already-linked furniture is not an error worth failing the run for
        return { linked: 0, error: e.message };
      });
  }

  // ─── WhatsApp receiver (same Fly host as the proxy) ───────────────────────
  // The receiver turns WhatsApp chat exports into bundles: the Word doc, its
  // Dropbox pictures, and the AI's reading of picture instructions. The key is
  // kept in this browser only (demo; see the PROXY_SECRET note in server.js).
  var RECEIVER_KEY_STORE = 'wdab-receiver-key';

  function receiverKey() {
    try { return localStorage.getItem(RECEIVER_KEY_STORE) || ''; } catch (e) { return ''; }
  }
  function setReceiverKey(k) {
    try { localStorage.setItem(RECEIVER_KEY_STORE, k); } catch (e) { /* private mode */ }
  }

  function receiverUrl(path) {
    return String(window.PROXY_BASE || '').replace(/\/$/, '') + path;
  }

  function receiverFetch(urlOrPath, opts) {
    opts = opts || {};
    var u = /^https?:/.test(urlOrPath) ? urlOrPath : receiverUrl(urlOrPath);
    return fetch(u, {
      method: opts.method || 'GET',
      body: opts.body,
      credentials: 'omit',
      headers: Object.assign({ 'X-Proxy-Key': receiverKey() }, opts.headers || {}),
    }).then(function (r) {
      if (r.status === 401) throw new Error('The receiver refused the key — check it and try again.');
      if (!r.ok) {
        // The receiver explains itself in { error }; show that, not a bare status.
        return r.json().then(function (j) { return j && j.error; }, function () { return null; }).then(function (why) {
          throw new Error(why || ('The receiver answered HTTP ' + r.status + '.'));
        });
      }
      return r;
    });
  }

  function fetchFromReceiver(u) {
    return receiverFetch(u).then(function (r) { return r.blob(); });
  }

  function bundleFileUrl(bundle, rel) {
    return receiverUrl('/bundles/' + bundle.key + '/files/' + rel.split('/').map(encodeURIComponent).join('/'));
  }

  // Upload order decides placement (applyImageIds fills slots in order):
  // the AI's hero and ranking when it made one, else a hero named in the chat
  // or doc, else the Dropbox filename order.
  function orderedBundleImages(bundle) {
    var images = (bundle.images || []).slice();
    var byName = function (n) { return images.filter(function (im) { return im.name === n; })[0]; };
    var front = [];
    if (bundle.selection) {
      [bundle.selection.hero].concat((bundle.selection.ranked || []).map(function (r) { return r.file; }))
        .forEach(function (n) { var im = byName(n); if (im && front.indexOf(im) === -1) front.push(im); });
    } else if (bundle.instructions && bundle.instructions.hero) {
      var h = String(bundle.instructions.hero).toLowerCase();
      var hit = images.filter(function (im) { return im.name.toLowerCase().indexOf(h) !== -1; })[0];
      if (hit) front.push(hit);
    }
    return front.concat(images.filter(function (im) { return front.indexOf(im) === -1; }));
  }

  var PICTURE_STATUS = {
    ready: 'Pictures downloaded',
    fetchable: 'Dropbox link found, pictures not downloaded yet',
    'flagged-link': 'Picture link needs a person (WeTransfer / press site)',
    waiting: 'No pictures yet',
    'built-by-hand': 'Already drafted — apple.news link posted in the chat',
    'doc-missing': 'Word doc missing from the export',
  };

  // The article summary under the picker: one status line, a few one-line
  // bullets for what the editor must know, and the AI's notes folded away.
  // (Benn, 2026-10-12: "too much text, needs to be collapsed bullet points".)
  function bundleInfoHtml(b) {
    var when = function (ts) { return esc(String(ts || '').replace('T', ' ').slice(5, 16)); };
    var n = (b.images || []).length;
    var head = '<strong>' + esc(PICTURE_STATUS[b.pictureStatus] || b.pictureStatus) + '</strong>' +
      (n ? ' · ' + n + ' picture' + (n === 1 ? '' : 's') : '') +
      (b.source === 'upload' ? ' · your upload' : ' · ' + esc(b.sender) + ', ' + when(b.ts));
    var link = function (href, text) { return '<a href="' + esc(href) + '" target="_blank" rel="noopener">' + text + '</a>'; };

    var must = [];
    if (b.statusNote) must.push(esc(brief(b.statusNote, 110)));
    if (b.picturesArrivedAt && !n) must.push('Pictures arrived ' + when(b.picturesArrivedAt) + '. Re-import to fetch them.');
    if (b.imported) must.push('In Studio since ' + esc(String(b.imported.at || '').slice(0, 10)) + '. Re-import fetches the pictures and runs the AI again.');
    else if (b.builtByHand) {
      must.push('Drafted by hand, ' + when(b.builtByHand.ts) + ' · ' + link(b.builtByHand.href, 'open draft') +
        (b.builtByHand.rule === 'order' ? ' · matched by timing, check it' : ''));
    }
    (b.pictures || []).forEach(function (p) {
      if (p.type !== 'dropbox-folder' && p.type !== 'dropbox-file') must.push('Fetch by hand: ' + link(p.href, esc(p.type) + ' link'));
    });
    (b.suggestedPictures || []).forEach(function (p) {
      must.push('Possible pictures (' + Math.round(p.confidence * 100) + '%): ' + link(p.href, 'link'));
    });
    (b.errors || []).forEach(function (e) { must.push('⚠ ' + esc(e.step) + ': ' + esc(brief(e.message, 90))); });

    var ai = [];
    if (b.selection) {
      ai.push(b.selection.hero
        ? 'Opener: ' + esc(b.selection.hero) + (b.selection.heroReason ? ' — ' + esc(brief(b.selection.heroReason, 80)) : '')
        : 'No picture works as the opener');
      var matched = (b.selection.placements || []).filter(function (x) { return x.section && x.confidence >= SECTION_CONFIDENCE; }).length;
      if (matched) ai.push(matched + ' picture' + (matched === 1 ? '' : 's') + ' matched to sections');
    }
    var ins = b.instructions;
    if (ins) {
      (ins.missing || []).forEach(function (m) { ai.push('Missing picture ' + esc(m.entry) + ': ' + esc(brief(m.note, 60))); });
      (ins.embeds || []).forEach(function (e) { ai.push('Embed: ' + esc(e.url)); });
      if (ins.embargo) ai.push('Embargo: ' + esc(brief(ins.embargo, 80)));
      (ins.otherNotes || []).forEach(function (t) { if (!ALREADY_HANDLED.test(t)) ai.push(esc(brief(t, 110))); });
    }
    var q = ((b.review && b.review.queries) || []).length;
    if (q) ai.push(q + ' copy quer' + (q === 1 ? 'y' : 'ies') + ' to check');
    var list = function (items) { return '<ul>' + items.map(function (l) { return '<li>' + l + '</li>'; }).join('') + '</ul>'; };
    return '<div>' + head + '</div>' + (must.length ? list(must) : '') +
      (ai.length ? '<details><summary>AI notes (' + ai.length + ')</summary>' + list(ai) + '</details>' : '');
  }
  // What the receiver is doing with the latest export, in plain words — so an
  // empty list is never mistaken for "it failed" while an export is still being
  // prepared (pictures + AI take about 1–1½ minutes per new article).
  function clockTime(iso) {
    var d = new Date(iso);
    if (isNaN(d)) return '';
    var t = d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    var today = new Date().toDateString() === d.toDateString();
    return today ? t : d.toLocaleDateString([], { day: 'numeric', month: 'short' }) + ' ' + t;
  }
  function exportStatusText(ex, hasArticles) {
    var pr = ex && ex.processing;
    if (pr) {
      var head = 'Export from ' + clockTime(pr.startedAt) + ': ';
      if (pr.stage !== 'preparing' || pr.toPrepare == null) return { busy: true, text: head + 'reading the chat…' };
      if (!pr.toPrepare) return { busy: true, text: head + 'nothing new so far, finishing up.' };
      var mins = Math.max(1, Math.round((pr.expectedSeconds || 100) / 60));
      return { busy: true, text: head + 'preparing ' + (pr.article ? '“' + pr.article + '”' : 'new articles') +
        (pr.toPrepare > 1 ? ' (' + (pr.preparing || 1) + ' of ' + pr.toPrepare + ')' : '') +
        ', about ' + mins + ' min. Updates by itself.' };
    }
    var last = ex && ex.last;
    if (!last) {
      return { text: hasArticles
        ? 'Watching for exports. New ones appear here, usually within 2–3 minutes.'
        : 'No export yet. Export the chat to the Dropbox folder “TG WhatsApp exports”.' };
    }
    if (last.error) return { failed: true, text: 'Export from ' + clockTime(last.startedAt) + ' failed (' + brief(last.error, 80) + '). Export again.' };
    var n = (last.newArticles || []).length;
    var pics = last.picturesArrived || [];
    return { text: 'Last export ' + clockTime(last.finishedAt) + ': ' +
      (n ? n + ' new (' + last.newArticles.join(', ') + ')' : pics.length ? 'no new articles' : 'nothing new') +
      (pics.length ? '; pictures arrived for ' + pics.join(', ') + ' — re-import to fetch them' : '') + '.' };
  }

  // ─── Comments in the created article ──────────────────────────────────────
  // Notes become Digital editor comments (see addComments in the engine). They
  // carry the creating editor's user id; the prefix says who raised them.
  // Comments are read at a glance (Benn, 2026-10-12): the problem and the
  // fix, or one short fact. No paragraphs, no long justifications.
  var AI_PREFIX = 'AI: ';
  var TOOL_PREFIX = '';
  // The first sentence of an AI answer, cut at a word, at most `max` characters
  function brief(text, max) {
    var t = String(text || '').replace(/\s+/g, ' ').trim();
    var first = t.match(/^.*?[.!?](?=\s|$)/);
    if (first && first[0].length >= 20) t = first[0];
    t = t.replace(/[.\s]+$/, '');
    if (t.length <= max) return t;
    return t.slice(0, max).replace(/\s+\S*$/, '') + '…';
  }
  var ALREADY_HANDLED = /feed headline|article headline|web gallery|^pics?:/i;

  function currentUserId() {
    try {
      var info = ContentStationSdk.getInfo() || {};
      var u = info.CurrentUser || info.User || {};
      return u.UserID || u.ShortName || u.Id || info.UserID || '';
    } catch (e) { return ''; }
  }

  // Notes from the WhatsApp bundle: AI copy queries, picture decisions, links
  // someone must fetch, instructions. Empty frames are added after placement.
  function bundleNotes(bundle) {
    var notes = [];
    var review = (bundle.review && bundle.review.queries) || [];
    review.forEach(function (q) {
      notes.push({
        anchor: { quote: q.quote, entry: q.entry },
        text: AI_PREFIX + brief(q.comment, 160),
      });
    });
    var head = [];
    // (The opener note is written after placement — see openerNote.)
    (bundle.pictures || []).forEach(function (p) {
      if (p.type !== 'dropbox-folder' && p.type !== 'dropbox-file') head.push('Fetch pictures by hand: ' + p.href);
    });
    (bundle.suggestedPictures || []).forEach(function (p) {
      head.push('Possible pictures (' + Math.round(p.confidence * 100) + '%): ' + p.href);
    });
    var ins = bundle.instructions;
    if (ins) {
      (ins.embeds || []).forEach(function (e) { head.push('Embed: ' + e.url + (e.where ? ' (' + brief(e.where, 40) + ')' : '')); });
      if (ins.embargo) head.push('Embargo: ' + brief(ins.embargo, 80));
      (ins.otherNotes || []).forEach(function (n) { if (!ALREADY_HANDLED.test(n)) head.push(brief(n, 110)); });
    }
    head.forEach(function (h) { notes.push({ anchor: {}, text: h }); });
    return notes;
  }

  // Which numbered entry a bundle image shows: the AI's match first, then a
  // filename like "15-F90.jpg" or "15.jpg". Camera names such as
  // "03.02.2026-Geely…" are not read as entry 3 (digit after the separator).
  function entryOfImage(im, bundle) {
    // The number the desk put on the file name comes first: "10 - Renault
    // 4.jpg" was chosen for entry 10 even though a Renault 4 is also entry 51.
    // The AI's reading of the picture is used only when the name has no number.
    var fromName = entryFromFileName(im.name);
    if (fromName != null) return fromName;
    var r = rankedFor(im.name, bundle);
    var fromAi = r && r.entry != null && String(r.entry).match(/\d+/);
    return fromAi ? Number(fromAi[0]) : null;
  }
  function entryFromFileName(name) {
    var m = String(name || '').match(/^0*(\d{1,3})(?:[\s._-]+(?!\d)|\.[a-z]+$)/i);
    return m ? Number(m[1]) : null;
  }
  function rankedFor(name, bundle) {
    return bundle.selection && (bundle.selection.ranked || []).filter(function (x) { return x.file === name; })[0];
  }
  // Pictures where the file name says one entry and the AI saw another. The
  // name wins; the editor is told, on that entry, what the AI thought.
  function numberConflictNotes(bundle) {
    if (!bundle || !bundle.selection) return [];
    return (bundle.selection.ranked || []).map(function (r) {
      var n = entryFromFileName(r.file);
      var aiRaw = r.aiEntry != null ? r.aiEntry : r.entry;
      var ai = aiRaw != null && String(aiRaw).match(/\d+/);
      if (n == null || !ai || Number(ai[0]) === n) return null;
      return { anchor: { entry: n }, text: 'File name says entry ' + n + '; AI thinks entry ' + Number(ai[0]) + '. Check the picture.' };
    }).filter(Boolean);
  }

  // Crosshead articles (first drives, Q&As): a picture frame above every
  // crosshead from the second one down, the way the list layouts have one
  // above every item (Benn, 2026-10-12). The first crosshead has none: it sits
  // just under the opener.
  //   - the first two pictures in upload order are the opener and the closing
  //     picture, and are not used again;
  //   - a frame is filled with the picture the AI matched most surely to the
  //     section below it (at or above SECTION_CONFIDENCE), and gets a comment
  //     with the AI's reason;
  //   - otherwise the frame is left empty for the editor, with the usual
  //     "no picture in this frame" comment. Empty frames must be deleted
  //     before publishing;
  //   - no gallery is built: where one belongs is the editor's call. Pictures
  //     not placed stay in the Dossier.
  // Numbers on these file names are export numbers and are ignored.
  // 0.5, not the 0.7 used for chat links: the AI is told "a guess is below
  // 0.5", and on the first real article (12 Oct) its sound matches — exhausts
  // to the section on the engine note, cornering to handling — came in at
  // 0.5–0.6.
  var SECTION_CONFIDENCE = 0.5;
  function sectionPictures(entries, bundle) {
    var sel = (bundle && bundle.selection) || {};
    var images = orderedBundleImages(bundle || {});
    var heads = entries.filter(function (e) { return e.crosshead; });
    if (heads.length < 2) return null;
    var norm = function (t) { return String(t || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); };
    var urlOf = function (name) {
      var im = images.filter(function (x) { return x.name === name; })[0];
      return im ? bundleFileUrl(bundle, im.file) : null;
    };
    var taken = {};
    images.slice(0, 2).forEach(function (im) { taken[im.name] = true; });

    var out = entries.map(function (e) { return Object.assign({}, e, { bodyParts: (e.bodyParts || []).slice() }); });
    var notes = [];
    var filled = 0, frames = 0;
    // Best matches first, so a picture goes to the section it fits most surely
    var claims = (sel.placements || []).filter(function (p) {
      return p.section && p.confidence >= SECTION_CONFIDENCE && urlOf(p.file);
    }).sort(function (a, b) { return b.confidence - a.confidence; });
    var chosen = {};
    claims.forEach(function (p) {
      var k = norm(p.section);
      if (taken[p.file] || chosen[k]) return;
      chosen[k] = p; taken[p.file] = true;
    });
    out.forEach(function (e, i) {
      if (!e.crosshead || i === 0 || entries[i] === heads[0]) return;
      // The frame above crosshead i is the last thing in the section before it
      var best = chosen[norm(e.crosshead)];
      out[i - 1].bodyParts.push({ picture: best ? urlOf(best.file) : PICTURE_EMPTY });
      frames++;
      if (!best) return;
      filled++;
      notes.push({ anchor: { crosshead: e.crosshead }, text: AI_PREFIX + 'picture above chosen for this section' + (best.reason ? ' — ' + brief(best.reason, 80) : '') });
    });
    // A picture matched to the first section has no frame; un-take it
    var first = chosen[norm(heads[0].crosshead)];
    if (first) delete taken[first.file];

    var left = images.filter(function (im) { return !taken[im.name]; }).length;
    if (frames) {
      notes.push({ anchor: {}, text: 'Frames: ' + filled + ' of ' + frames + ' filled.' +
        (frames - filled ? ' Fill or delete the empty ones.' : '') +
        (left ? ' ' + left + ' spare picture' + (left === 1 ? '' : 's') + ' in the Dossier.' : '') });
    }
    return { entries: out, notes: notes, frames: frames, filled: filled };
  }

  // Image object ids per frame, in frame order (null = leave empty).
  // created: [{ id, url }] in upload order; placement: { byUrl: {url: entry}, heroUrl }.
  // A numbered list whose pictures mostly know their entry is placed by entry,
  // so a missing picture leaves its own frame empty instead of shifting the
  // rest; otherwise frames fill in upload order, as before.
  function assignSlots(digital, created, placement) {
    var content = digital.data.content || [];
    var slots = [];
    content.forEach(function (c, i) {
      if (c.identifier !== 'image' && c.identifier !== 'header-image') return;
      var title = content.slice(i + 1).filter(function (x) { return x.identifier === 'title'; })[0];
      var num = c.identifier === 'image' && title && (opsText(title.content.text || []).match(/^\s*0*(\d+)/) || [])[1];
      slots.push({ header: c.identifier === 'header-image', entry: num ? Number(num) : null, isEntry: c.identifier === 'image' });
    });
    // Type 4 snippet lists: titles carry no number, so entry N is the Nth
    // entry picture frame — filename numbers and the AI's matches still apply.
    if (!slots.some(function (s) { return s.entry != null; })) {
      var n = 0;
      slots.forEach(function (s) { if (s.isEntry) s.entry = ++n; });
      if (n < 3) slots.forEach(function (s) { s.entry = null; });
    }
    var ids = created.map(function (c) { return c.id; });
    if (!placement) return ids;
    var known = created.filter(function (c) { return placement.byUrl[c.url] != null; });
    var numbered = slots.some(function (s) { return s.entry != null; });
    if (!numbered || known.length < created.length / 2) return ids;

    var out = slots.map(function () { return null; });
    var used = {};
    known.forEach(function (c) {
      var k = slots.findIndex(function (s, i) { return s.entry === placement.byUrl[c.url] && !out[i]; });
      if (k >= 0) { out[k] = c.id; used[c.url] = true; }
    });
    // Opener: the AI's pick, which may reuse any entry's picture (the object is
    // simply placed twice) — but never the picture in the frame directly under
    // the header, or the same image would sit on top of itself. A picture with
    // no entry number (e.g. "opener.jpg") is the fallback. Otherwise the header
    // stays empty and gets a comment.
    var headerIdx = slots.findIndex(function (s) { return s.header; });
    if (headerIdx >= 0) {
      var firstEntryId = out.slice(headerIdx + 1).filter(function (id, i) { return slots[headerIdx + 1 + i].entry != null; })[0];
      var firstEntrySlot = slots.slice(headerIdx + 1).filter(function (s) { return s.entry != null; })[0];
      var directlyBelow = function (c) {
        return (firstEntryId && c.id === firstEntryId) ||
               (firstEntrySlot && placement.byUrl[c.url] === firstEntrySlot.entry);
      };
      var hero = created.filter(function (c) { return c.url === placement.heroUrl && !directlyBelow(c); })[0] ||
                 created.filter(function (c) { return placement.byUrl[c.url] == null && !used[c.url]; })[0];
      if (hero) out[headerIdx] = hero.id;
    }
    return out; // pictures without a matching frame stay in the Dossier only
  }

  // What actually went in the header, said after placement so the comment can
  // never contradict the article: the AI's pick and why, or why its pick was
  // not used. An empty header gets its own note from pictureNotes.
  function openerNote(digital, created, bundle) {
    var sel = bundle && bundle.selection;
    var header = (digital.data.content || []).filter(function (c) { return c.identifier === 'header-image'; })[0];
    var placedId = header && header.content && header.content.image && header.content.image.id;
    var nameOf = function (url) { return decodeURIComponent(String(url).split('/').pop()); };
    var total = bundle ? (bundle.images || []).length : created.length;
    if (placedId) {
      var c = created.filter(function (x) { return String(x.id) === String(placedId); })[0];
      var name = c ? nameOf(c.url) : String(placedId);
      var why = sel && sel.hero === name && sel.heroReason ? ' — ' + brief(sel.heroReason, 80) : '';
      return { anchor: {}, text: 'Opener' + (why ? ' (AI)' : '') + ': ' + name + why };
    }
    if (sel && sel.hero) {
      return { anchor: {}, text: 'No opener: the AI\'s pick (' + sel.hero + ') is the picture directly below. Choose another.' };
    }
    return null;
  }

  // Frames still empty after placement, with the chat's reason where known;
  // and frames the chat says were missing but that got a picture anyway.
  function pictureNotes(digital, bundle) {
    var missing = (bundle && bundle.instructions && bundle.instructions.missing) || [];
    var notes = emptySlotNotes(digital, missing, TOOL_PREFIX);
    var frames = (digital.data.content || []).filter(function (c) {
      return c.identifier === 'image' || c.identifier === 'header-image';
    }).length;
    // Nothing placed at all: one note on the headline, not one per frame.
    if (frames > 2 && notes.length === frames) {
      var why = bundle && bundle.pictureStatus === 'waiting'
        ? ' Doc says: ' + brief((bundle.doc && bundle.doc.picLineText) || 'no picture link', 60) + '.'
        : '';
      notes = [{ anchor: {}, text: 'No pictures placed: all ' + frames + ' frames are empty.' + why }];
    }
    var emptyEntries = notes.map(function (n) { return n.text; }).join('\n');
    missing.forEach(function (m) {
      if (emptyEntries.indexOf(m.note) !== -1) return; // already said on the empty frame
      notes.push({ anchor: { entry: m.entry }, text: 'Chat says this picture was missing (' + brief(m.note, 60) + '). Check it.' });
    });
    return notes.concat(numberConflictNotes(bundle));
  }

  // ─── Shared converter UI ───────────────────────────────────────────────────
  // Look taken from the content reuse plug-in (claude-content-re-use-assets-
  // studio/plugin/src/15-styles.js): slate + amber + teal, numbered stepper,
  // staged progress list with a spinning mark. Class names are the contract
  // with the markup below; restyle freely.
  var CSS = [
    '.wdab{--w-page:#f8fafc;--w-surface:#fff;--w-line:#e2e8f0;--w-field:#d1d5db;--w-ink:#0f172a;--w-ink2:#374151;--w-ink3:#4b5563;--w-ink4:#94a3b8;' +
      '--w-accent:#f59e0b;--w-accent-soft:#fff8e8;--w-mute:#f1f5f9;--w-teal:#0f766e;--w-ok:#0d9488;--w-bad:#8c1d18;--w-bad-soft:#fdecea;--w-warn:#7a5300;--w-warn-soft:#fff6e5;' +
      'max-width:640px;margin:0 auto;padding:0 2px 8px;font:14px/1.45 Geist,Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif;color:var(--w-ink)}',
    '.wdab-scroll{max-height:calc(100vh - 140px);overflow-y:auto;-webkit-overflow-scrolling:touch}',
    '.wdab-modal{max-height:min(620px,calc(100vh - 280px));overflow-y:auto;-webkit-overflow-scrolling:touch}',
    '.wdab button,.wdab input,.wdab select{font-family:inherit}',
    '.wdab .wdab-hidden{display:none}',
    // Stepper
    '.wdab-steps{display:flex;align-items:center;gap:14px;list-style:none;margin:0 0 14px;padding:10px 0;border-bottom:1px solid var(--w-line);position:sticky;top:0;background:var(--w-surface);z-index:2}',
    '.wdab-step{display:flex;align-items:center;gap:6px;font-size:13px;font-weight:500;color:var(--w-ink4);white-space:nowrap}',
    '.wdab-step i{font-style:normal;width:22px;height:22px;border-radius:50%;display:grid;place-items:center;background:var(--w-mute);border:1px solid var(--w-line);color:var(--w-ink4);font-size:11px;font-weight:600}',
    '.wdab-step.is-on{color:var(--w-ink);font-weight:600}',
    '.wdab-step.is-on i,.wdab-step.is-done i{background:var(--w-teal);border-color:var(--w-teal);color:#fff}',
    '.wdab-step.is-done{color:var(--w-ink3)}',
    '.wdab-step+.wdab-step::before{content:"";width:18px;height:1px;background:var(--w-line);margin-right:8px}',
    // Fields
    '.wdab .wdab-row{margin-bottom:12px}',
    '.wdab label{display:block;font-size:13px;font-weight:500;color:var(--w-ink3);margin:0 0 4px}',
    '.wdab select,.wdab input[type=text],.wdab input[type=password],.wdab input[type=url]{width:100%;box-sizing:border-box;border:1px solid var(--w-field);border-radius:6px;padding:9px 12px;font-size:14px;color:var(--w-ink);background:var(--w-surface)}',
    // Own arrow: the browser's sits against the edge and is too faint to read as a menu
    '.wdab select{appearance:none;-webkit-appearance:none;padding-right:40px;cursor:pointer;background-repeat:no-repeat;background-position:right 14px center;background-size:14px 14px;' +
      'background-image:url("data:image/svg+xml,%3Csvg xmlns=%27http://www.w3.org/2000/svg%27 viewBox=%270 0 16 16%27 fill=%27none%27 stroke=%27%230f172a%27 stroke-width=%272.2%27 stroke-linecap=%27round%27 stroke-linejoin=%27round%27%3E%3Cpath d=%27M3 6l5 5 5-5%27/%3E%3C/svg%3E")}',
    '.wdab input[type=file]{width:100%;box-sizing:border-box;font-size:13px;color:var(--w-ink3);border:1px dashed var(--w-field);border-radius:6px;padding:6px;background:var(--w-surface)}',
    '.wdab input[type=file]::file-selector-button{margin-right:12px;padding:7px 14px;border:1px solid var(--w-field);border-radius:6px;background:var(--w-surface);color:var(--w-ink);font-family:inherit;font-size:13px;font-weight:600;cursor:pointer}',
    '.wdab input[type=file]::file-selector-button:hover{background:var(--w-mute)}',
    '.wdab input[type=checkbox]{width:auto;margin:0 6px 0 0;vertical-align:middle;accent-color:var(--w-teal)}',
    '.wdab-row label input[type=checkbox]+span{font-weight:400;color:var(--w-ink2)}',
    '.wdab .wdab-feed input{border-color:var(--w-accent);background:var(--w-accent-soft)}',
    '.wdab .wdab-feed-note{display:inline-block;font-size:11px;color:var(--w-warn);background:var(--w-warn-soft);border-radius:4px;padding:1px 6px;margin-left:8px;font-weight:500}',
    // Buttons
    // Pinned to the bottom of the dialog, so the next action never scrolls away
    '.wdab-foot{display:flex;gap:10px;justify-content:flex-end;align-items:center;margin-top:14px;padding:12px 0 4px;border-top:1px solid var(--w-line);position:sticky;bottom:0;background:var(--w-surface);z-index:2}',
    '.wdab button.wdab-btn{border:none;border-radius:6px;padding:10px 18px;background:var(--w-accent);color:var(--w-ink);font-size:14px;font-weight:700;cursor:pointer}',
    '.wdab button.wdab-btn:disabled{opacity:.45;cursor:default}',
    '.wdab button.wdab-back{padding:10px 16px;border:1px solid var(--w-field);border-radius:6px;background:var(--w-surface);color:var(--w-ink);font-size:14px;font-weight:600;cursor:pointer;margin-right:auto}',
    '.wdab button.wdab-linkbtn{border:0;background:none;padding:0;font-size:12px;font-weight:600;color:var(--w-ink3);cursor:pointer;text-decoration:underline}',
    // Boxes, notes, alerts
    '.wdab .wdab-box{background:var(--w-surface);border:1px solid var(--w-line);border-radius:8px;padding:12px 14px;margin-top:10px}',
    '.wdab .wdab-wa-info{background:var(--w-page);border:1px solid var(--w-line);border-radius:8px;padding:10px 12px;margin-top:8px;font-size:12px;color:var(--w-ink2)}',
    '.wdab .wdab-wa-info ul{margin:4px 0 0;padding-left:16px}',
    '.wdab .wdab-wa-info li{margin:2px 0}',
    '.wdab .wdab-wa-info details{margin-top:6px}',
    '.wdab .wdab-wa-info summary,.wdab details.wdab-more>summary{cursor:pointer;font-weight:600;color:var(--w-ink3);font-size:12px}',
    '.wdab .wdab-wa-status{display:flex;gap:10px;align-items:center;margin-top:8px;font-size:12px;color:var(--w-ink3);background:var(--w-mute);border-radius:6px;padding:7px 10px}',
    '.wdab .wdab-wa-status>span{flex:1 1 auto;min-width:0}',
    '.wdab .wdab-wa-status.wdab-busy{background:var(--w-accent-soft);color:var(--w-warn)}',
    '.wdab .wdab-wa-status.wdab-failed{background:var(--w-bad-soft);color:var(--w-bad)}',
    '.wdab .wdab-error{display:none;white-space:pre-wrap;margin-top:10px;padding:10px 12px;border-radius:6px;background:var(--w-bad-soft);color:var(--w-bad);font-size:13px}',
    '.wdab .wdab-warn{display:none;margin-bottom:12px;padding:10px 12px;border-radius:6px;background:var(--w-warn-soft);border:1px solid #f0dcb4;color:var(--w-warn);font-size:13px}',
    '.wdab .wdab-warn ul{margin:6px 0 0;padding-left:18px}',
    '.wdab .wdab-note{color:var(--w-ink3);font-size:12px;margin:6px 0 0}',
    '.wdab .wdab-chip{display:inline-block;background:var(--w-mute);border-radius:4px;padding:3px 8px;font-size:12px;font-weight:600;color:var(--w-ink3);margin-bottom:12px}',
    '.wdab .wdab-entries{max-height:200px;overflow-y:auto;border:1px solid var(--w-line);border-radius:6px;padding:8px 12px;margin-top:6px;color:var(--w-ink2);font-size:13px}',
    '.wdab .wdab-entries .n{color:var(--w-ink4);margin-right:6px}',
    '.wdab h3.wdab-h{margin:4px 0 10px;font-size:15px;font-weight:600}',
    // Staged progress
    '.wdab-stages{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:2px}',
    '.wdab-stage{display:grid;grid-template-columns:20px minmax(0,1fr) auto;gap:4px 10px;align-items:center;padding:6px 0;font-size:13px;color:var(--w-ink4)}',
    '.wdab-stage.is-active,.wdab-stage.is-done{color:var(--w-ink)}',
    '.wdab-stage.is-active{font-weight:600}',
    '.wdab-stage.is-failed,.wdab-stage.is-failed .wdab-stage-d{color:var(--w-bad)}',
    '.wdab-stage-d{grid-column:2 / 4;font-size:12px;font-weight:400;color:var(--w-ink3);margin-top:-2px;overflow-wrap:anywhere}',
    '.wdab-stage-t{font-size:12px;font-weight:400;color:var(--w-ink4);font-variant-numeric:tabular-nums}',
    '.wdab-mark{width:18px;height:18px;border-radius:50%;display:grid;place-items:center;font-size:11px;font-weight:700;box-sizing:border-box}',
    '.wdab-mark.pending{border:1.5px solid var(--w-field)}',
    '.wdab-mark.done{background:var(--w-ok);color:#fff}',
    '.wdab-mark.failed{background:var(--w-bad);color:#fff}',
    '.wdab-mark.active{border:2px solid var(--w-mute);border-top-color:var(--w-accent);animation:wdab-spin .8s linear infinite}',
    '@keyframes wdab-spin{to{transform:rotate(360deg)}}',
    '.wdab-sbar{grid-column:2 / 4;height:4px;border-radius:2px;background:var(--w-mute);overflow:hidden}',
    '.wdab-sbar>span{display:block;height:100%;background:var(--w-accent);transition:width .3s ease}',
    '@media (prefers-reduced-motion:reduce){.wdab-mark.active{animation-duration:2.4s}}'
  ].join('\n');

  var cssInjected = false;
  // Build id, replaced by build-plugin.js. Check it in Studio's console with
  // __wdVersion to confirm which build the browser actually loaded.
  var PLUGIN_BUILD = '__BUILD_ID__';
  try {
    window.__wdVersion = PLUGIN_BUILD;
    console.info('[word-digital] plug-in build ' + PLUGIN_BUILD);
  } catch (e) { /* non-fatal */ }

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

  // A staged progress list: one row per step, with a mark (pending, spinning,
  // done, failed), the time it took and an optional bar. start(id) finishes
  // whatever was running before it, so the callers only say what begins.
  var STAGE_MARK = { done: '✓', failed: '!' };
  function stageList(el, defs) {
    var st = {};
    defs.forEach(function (d) { st[d.id] = { state: 'pending', detail: d.detail || '', started: 0, took: null, bar: null }; });
    var clock = function (ms) { var x = Math.round(ms / 1000); return Math.floor(x / 60) + ':' + String(x % 60).padStart(2, '0'); };
    function render() {
      if (!el) return;
      el.innerHTML = defs.map(function (d) {
        var x = st[d.id];
        var t = x.state === 'active' ? clock(Date.now() - x.started) : x.took != null ? clock(x.took) : '';
        return '<li class="wdab-stage is-' + x.state + '" data-id="' + d.id + '">' +
          '<span class="wdab-mark ' + x.state + '" aria-hidden="true">' + (STAGE_MARK[x.state] || '') + '</span>' +
          '<span>' + esc(d.label) + '</span><span class="wdab-stage-t">' + t + '</span>' +
          (x.bar != null && x.state === 'active' ? '<div class="wdab-sbar"><span style="width:' + x.bar + '%"></span></div>' : '') +
          (x.detail ? '<div class="wdab-stage-d">' + esc(x.detail) + '</div>' : '') + '</li>';
      }).join('');
    }
    // The clock ticks without re-rendering, so the spinner never restarts
    var tick = setInterval(function () {
      if (!el || !el.isConnected) return clearInterval(tick);
      defs.forEach(function (d) {
        if (st[d.id].state !== 'active') return;
        var t = el.querySelector('[data-id="' + d.id + '"] .wdab-stage-t');
        if (t) t.textContent = clock(Date.now() - st[d.id].started);
      });
    }, 1000);
    var settle = function (state, detail) {
      defs.forEach(function (d) {
        var x = st[d.id];
        if (x.state !== 'active') return;
        x.state = state; x.took = Date.now() - x.started; x.bar = null;
        if (detail != null) x.detail = detail;
      });
    };
    var api = {
      start: function (id, detail) {
        if (!st[id]) return api;
        settle('done');
        st[id].state = 'active'; st[id].started = Date.now();
        if (detail != null) st[id].detail = detail;
        render(); return api;
      },
      // Bar and detail change in place — again so the spinner keeps turning
      progress: function (id, done, total, detail) {
        var x = st[id]; if (!x) return api;
        var first = x.bar == null;
        x.bar = total ? Math.round(100 * done / total) : 0;
        if (detail != null) x.detail = detail;
        var row = el && el.querySelector('[data-id="' + id + '"]');
        var bar = row && row.querySelector('.wdab-sbar > span'), det = row && row.querySelector('.wdab-stage-d');
        if (first || !bar || (detail != null && !det)) render();
        else { bar.style.width = x.bar + '%'; if (det && detail != null) det.textContent = detail; }
        return api;
      },
      done: function (id, detail) {
        var x = st[id]; if (!x) return api;
        if (x.state !== 'active') { x.started = Date.now(); }
        x.state = 'done'; x.took = x.started ? Date.now() - x.started : null; x.bar = null;
        if (detail != null) x.detail = detail;
        render(); return api;
      },
      // One step went wrong but the run carries on (pictures failed, article still made)
      failed: function (id, detail) {
        var x = st[id]; if (!x) return api;
        x.state = 'failed'; x.took = x.started ? Date.now() - x.started : null; x.bar = null;
        if (detail != null) x.detail = detail;
        render(); return api;
      },
      finish: function () { settle('done'); render(); clearInterval(tick); return api; },
      fail: function (detail) { settle('failed', detail); render(); clearInterval(tick); return api; },
    };
    render();
    return api;
  }

  function formHtml(p, actionLabel) {
    var step = function (n, label) {
      return '<li class="wdab-step' + (n === 1 ? ' is-on' : '') + '" data-step="' + n + '"><i>' + n + '</i><span>' + label + '</span></li>';
    };
    return '<div class="wdab">' +
      '<ol class="wdab-steps" id="' + p + '-steps">' + step(1, 'Source') + step(2, 'Check') + step(3, 'Create') + '</ol>' +
      '<div class="wdab-warn" id="' + p + '-brand-warn"></div>' +

      // ── Stage 1: where the article comes from ──
      '<div id="' + p + '-stage1">' +
      '  <div class="wdab-row">' +
      '    <label for="' + p + '-source">Source</label>' +
      '    <select id="' + p + '-source">' +
      '      <option value="docx">Word document (.docx)</option>' +
      '      <option value="url">TopGear article URL</option>' +
      '      <option value="whatsapp">From WhatsApp</option>' +
      '    </select>' +
      '  </div>' +
      '  <div class="wdab-row wdab-hidden" id="' + p + '-wa-row">' +
      '    <label for="' + p + '-wa-bundle">Article</label>' +
      '    <select id="' + p + '-wa-bundle"><option value="">—</option></select>' +
      '    <div class="wdab-wa-status" id="' + p + '-wa-status" role="status" aria-live="polite">' +
      '      <span id="' + p + '-wa-status-text">Checking for exports…</span>' +
      '      <button type="button" class="wdab-linkbtn" id="' + p + '-wa-refresh">Refresh</button>' +
      '    </div>' +
      '    <div class="wdab-wa-info wdab-hidden" id="' + p + '-wa-info"></div>' +
      '    <details class="wdab-more" style="margin-top:8px" id="' + p + '-wa-key-box"><summary>Receiver key</summary>' +
      '      <input type="password" id="' + p + '-wa-key" autocomplete="off" style="margin-top:6px">' +
      '    </details>' +
      '  </div>' +
      '  <div class="wdab-row" id="' + p + '-docx-row">' +
      '    <label for="' + p + '-file">Word document (.docx)</label>' +
      '    <input type="file" id="' + p + '-file" accept=".docx">' +
      '    <label style="margin-top:10px"><input type="checkbox" id="' + p + '-docx-fetch"> <span>Fetch the linked pictures and run the AI checks (1–2 minutes)</span></label>' +
      '    <details class="wdab-more wdab-hidden" style="margin-top:8px" id="' + p + '-docx-key-row"><summary>Receiver key</summary>' +
      '      <input type="password" id="' + p + '-docx-key" autocomplete="off" style="margin-top:6px">' +
      '    </details>' +
      '    <div class="wdab-wa-info wdab-hidden" id="' + p + '-docx-info"></div>' +
      '  </div>' +
      '  <div class="wdab-row wdab-hidden" id="' + p + '-url-row">' +
      '    <label for="' + p + '-url">TopGear article URL</label>' +
      '    <input type="url" id="' + p + '-url" placeholder="https://www.topgear.com/…">' +
      '  </div>' +
      '  <div class="wdab-row">' +
      '    <label for="' + p + '-type">Article template</label>' +
      '    <select id="' + p + '-type">' +
      '      <option value="auto" selected>Auto-detect from the document</option>' +
      '      <option value="countdown">Type 1 — Numbered countdown (50 → 1)</option>' +
      '      <option value="ascending">Type 2 — Numbered ascending (1 → 50)</option>' +
      '      <option value="crosshead">Type 3 — Crosshead / generic article</option>' +
      '      <option value="snippet">Type 4 — Snippet list (unnumbered items)</option>' +
      '    </select>' +
      '  </div>' +
      '  <div class="wdab-box wdab-hidden" id="' + p + '-load"><ol class="wdab-stages" id="' + p + '-load-list"></ol></div>' +
      '  <p class="wdab-error" id="' + p + '-parse-error"></p>' +
      '  <div class="wdab-foot"><button class="wdab-btn" id="' + p + '-parse" disabled>Load</button></div>' +
      '</div>' +

      // ── Stage 2: what was found, before anything is created ──
      '<div id="' + p + '-preview" class="wdab-hidden">' +
      '  <span class="wdab-chip wdab-hidden" id="' + p + '-type-note"></span>' +
      '  <div class="wdab-row wdab-feed wdab-hidden" id="' + p + '-feed-row">' +
      '    <label>Feed headline<span class="wdab-feed-note" id="' + p + '-feed-note">Saved to C_HEADLINE</span></label>' +
      '    <input type="text" id="' + p + '-feed">' +
      '  </div>' +
      '  <div class="wdab-row"><label>Article title</label><input type="text" id="' + p + '-title"></div>' +
      '  <div class="wdab-row"><label>Subtitle</label><input type="text" id="' + p + '-subtitle"></div>' +
      '  <div class="wdab-row"><label>Author</label><input type="text" id="' + p + '-author"></div>' +
      '  <div class="wdab-row wdab-hidden" id="' + p + '-images-row">' +
      '    <label><input type="checkbox" id="' + p + '-images-add" checked> <span id="' + p + '-images-label"></span></label>' +
      '    <p class="wdab-note" id="' + p + '-images-progress"></p>' +
      '  </div>' +
      '  <div class="wdab-row">' +
      '    <label><input type="checkbox" id="' + p + '-comments-add" checked> <span>Add comments for missing pictures and things to check</span></label>' +
      '    <p class="wdab-note" id="' + p + '-comments-note"></p>' +
      '  </div>' +
      '  <div class="wdab-warn" id="' + p + '-warn"><strong>Editor instructions left in the copy.</strong> Kept in place; delete any that shouldn\'t ship.<ul id="' + p + '-warn-list"></ul></div>' +
      '  <details class="wdab-more"><summary>Entries (<span id="' + p + '-count">0</span>)</summary>' +
      '    <div class="wdab-entries" id="' + p + '-entries"></div>' +
      '  </details>' +
      (actionLabel
        ? '<p class="wdab-note" id="' + p + '-note"></p>' +
          '<div class="wdab-foot"><button type="button" class="wdab-back" id="' + p + '-back">Back</button>' +
          '<button class="wdab-btn" id="' + p + '-action">' + esc(actionLabel) + '</button></div>'
        : '') +
      '</div>' +

      // ── Stage 3: creating, step by step ──
      '<div id="' + p + '-stage3" class="wdab-hidden">' +
      '  <h3 class="wdab-h" id="' + p + '-run-title">Creating the article</h3>' +
      '  <div class="wdab-box" style="margin-top:0"><ol class="wdab-stages" id="' + p + '-run-list"></ol></div>' +
      '  <p class="wdab-error" id="' + p + '-action-error"></p>' +
      '  <div class="wdab-foot wdab-hidden" id="' + p + '-run-foot"><button type="button" class="wdab-back" id="' + p + '-run-back">Back</button></div>' +
      '</div>' +
      '</div>';
  }

  // Wires the converter form; returns a controller for reading the result.
  function wireForm(p, onAction) {
    var $ = function (id) { return document.getElementById(p + '-' + id); };
    var state = { parsedData: null, uploadedFilename: '' };

    // 1 Source → 2 Check → 3 Create. One stage is on screen at a time.
    function setStage(n) {
      $('stage1').classList.toggle('wdab-hidden', n !== 1);
      $('preview').classList.toggle('wdab-hidden', n !== 2);
      $('stage3').classList.toggle('wdab-hidden', n !== 3);
      Array.prototype.forEach.call($('steps').children, function (li) {
        var k = Number(li.getAttribute('data-step'));
        li.classList.toggle('is-on', k === n);
        li.classList.toggle('is-done', k < n);
        li.firstChild.textContent = k < n ? '✓' : String(k);
      });
      var box = $('steps').closest('.wdab-modal') || $('steps').closest('.wdab-scroll');
      if (box) box.scrollTop = 0;
    }
    if ($('back')) $('back').addEventListener('click', function () { setStage(1); });
    $('run-back').addEventListener('click', function () { setStage(2); });

    $('file').addEventListener('change', function () {
      $('parse').disabled = !$('file').files.length;
      $('preview').classList.add('wdab-hidden');
      state.parsedData = null;
      state.bundle = null;
      $('docx-info').classList.add('wdab-hidden');
      $('parse-error').style.display = 'none';
    });

    // Word source, optional: send the doc to the receiver for its pictures and
    // the AI checks. On by default only when a receiver key is already stored,
    // so a Word upload without a key behaves exactly as before.
    $('docx-key').value = receiverKey();
    $('docx-fetch').checked = !!receiverKey();
    // Same as on the WhatsApp source: folded away once a key is stored
    $('docx-key-row').open = !receiverKey();
    function syncDocxFetch() { $('docx-key-row').classList.toggle('wdab-hidden', !$('docx-fetch').checked); }
    $('docx-fetch').addEventListener('change', syncDocxFetch);
    $('docx-key').addEventListener('change', function () {
      setReceiverKey($('docx-key').value.trim());
      $('wa-key').value = receiverKey();
    });
    syncDocxFetch();

    $('source').addEventListener('change', function () {
      var src = $('source').value;
      $('docx-row').classList.toggle('wdab-hidden', src !== 'docx');
      $('url-row').classList.toggle('wdab-hidden', src !== 'url');
      $('wa-row').classList.toggle('wdab-hidden', src !== 'whatsapp');
      if (src === 'whatsapp') loadBundles();
      refreshParse();
    });
    $('url').addEventListener('input', refreshParse);

    // ── WhatsApp source ──
    $('wa-key').value = receiverKey();
    // Folded away once a key is stored; open when one is still needed
    $('wa-key-box').open = !receiverKey();
    $('wa-key').addEventListener('change', function () {
      setReceiverKey($('wa-key').value.trim());
      $('docx-key').value = receiverKey();
      loadBundles();
    });

    function showParseError(msg) {
      $('parse-error').textContent = msg;
      $('parse-error').style.display = msg ? 'block' : 'none';
    }

    // Status line + self-refresh. While an export is being processed the list is
    // re-read every few seconds; once it finishes the list reloads by itself,
    // unless an article is already selected (then it just says so).
    var pollTimer = null;
    function showExportStatus(ex, hasArticles) {
      if (!$('wa-status')) return false; // dialog closed
      var st = exportStatusText(ex, hasArticles);
      $('wa-status-text').textContent = st.text;
      $('wa-status').classList.toggle('wdab-busy', !!st.busy);
      $('wa-status').classList.toggle('wdab-failed', !!st.failed);
      // While an export is being prepared, an empty list must not say "nothing new".
      var first = $('wa-bundle').options[0];
      if (first && !first.value && /^Nothing new/.test(first.textContent) && st.busy) first.textContent = 'Still preparing — see above';
      return !!st.busy;
    }
    // Keeps checking while the WhatsApp source is showing — the dialog is often
    // open before the export arrives. Quicker while an export is in progress.
    var seenExport = null; // finishedAt of the last export the list reflects
    function schedulePoll(busy) {
      clearTimeout(pollTimer);
      pollTimer = setTimeout(function () {
        if (!$('wa-status') || $('source').value !== 'whatsapp') return;
        receiverFetch('/bundles?all=1').then(function (r) { return r.json(); }).then(function (j) {
          var nowBusy = showExportStatus(j.export, (j.bundles || []).length > 0);
          var finished = j.export && j.export.last ? j.export.last.finishedAt : null;
          if (nowBusy || finished === seenExport) return schedulePoll(nowBusy);
          // A new export has finished. Reload the list, but never under an
          // article that is being worked on.
          if (!$('wa-bundle').value) return loadBundles();
          $('wa-status-text').textContent += ' Press Refresh to see new articles.';
          seenExport = finished;
          schedulePoll(false);
        }).catch(function () { schedulePoll(busy); });
      }, busy ? 5000 : 12000);
    }
    $('wa-refresh').addEventListener('click', function () { loadBundles(); });

    function loadBundles() {
      clearTimeout(pollTimer);
      state.bundle = null;
      $('wa-info').classList.add('wdab-hidden');
      $('wa-bundle').innerHTML = '<option value="">Loading…</option>';
      refreshParse();
      receiverFetch('/bundles?all=1')
        .then(function (r) { return r.json(); })
        .then(function (j) {
          showParseError('');
          var busy = showExportStatus(j.export, (j.bundles || []).length > 0);
          seenExport = j.export && j.export.last ? j.export.last.finishedAt : null;
          schedulePoll(busy);
          var list = j.bundles || [];
          // New first; then docs the chat shows were already built by hand
          // (BRS replied with an apple.news draft); then ones imported here.
          var waiting = list.filter(function (b) { return !b.imported && !b.builtByHand; });
          var byHand = list.filter(function (b) { return !b.imported && b.builtByHand; });
          var done = list.filter(function (b) { return b.imported; });
          var day = function (ts) { return String(ts || '').slice(0, 10); };
          var opt = function (b) {
            // Pictures that came in after the article was made (or drafted by hand)
            var newPics = b.picturesArrivedAt && (!b.imported || b.picturesArrivedAt > b.imported.at) && !b.imageCount;
            var label = newPics ? 'pictures arrived, re-import'
              : b.imported ? 'imported ' + day(b.imported.at)
              : b.builtByHand ? 'draft posted ' + day(b.builtByHand.ts)
              : (PICTURE_STATUS[b.pictureStatus] || b.pictureStatus);
            return '<option value="' + esc(b.key) + '">' + esc(b.name) + ' — ' + esc(label) + '</option>';
          };
          var group = function (label, items) {
            return items.length ? '<optgroup label="' + esc(label) + '">' + items.map(opt).join('') + '</optgroup>' : '';
          };
          $('wa-bundle').innerHTML =
            '<option value="">' + (waiting.length ? 'Choose an article…' : busy ? 'Still preparing — see above' : 'Nothing new from WhatsApp') + '</option>' +
            group('Waiting', waiting) +
            group('Already drafted — apple.news link posted in the chat', byHand) +
            group('Already in Studio — choose to re-import', done);
        })
        .catch(function (e) {
          $('wa-bundle').innerHTML = '<option value="">—</option>';
          if ($('wa-status-text')) $('wa-status-text').textContent = 'Could not reach the WhatsApp receiver.';
          showParseError(e.message);
        });
    }

    $('wa-bundle').addEventListener('change', function () {
      var key = $('wa-bundle').value;
      state.bundle = null;
      $('preview').classList.add('wdab-hidden');
      $('wa-info').classList.add('wdab-hidden');
      refreshParse();
      if (!key) return;
      receiverFetch('/bundles/' + key)
        .then(function (r) { return r.json(); })
        .then(function (b) {
          if ($('wa-bundle').value !== key) return; // changed while loading
          state.bundle = b;
          $('wa-info').innerHTML = bundleInfoHtml(b);
          $('wa-info').classList.remove('wdab-hidden');
          refreshParse();
        })
        .catch(function (e) { showParseError(e.message); });
    });

    // Slots available in the layout the current article type would produce.
    var ctlGalleries = { slideshows: 0, slides: 0 };
    function ctlSlotCount() {
      if (!state.parsedData) return 0;
      var t = deepClone(TEMPLATES[state.parsedData.type]);
      var d = state.parsedData.type === 'crosshead'
        ? buildCrosshead(t, state.parsedData.meta, state.parsedData.entries)
        : buildNumbered(t, state.parsedData.meta, state.parsedData.entries, state.parsedData.type);
      ctlGalleries = gallerySummary(d);
      return (d.data.content || []).filter(function (c) {
        return c.identifier === 'image' || c.identifier === 'header-image';
      }).length;
    }

    function refreshParse() {
      var src = $('source').value;
      $('parse').disabled = src === 'docx' ? !$('file').files.length
        : src === 'url' ? !$('url').value.trim()
        : !(state.bundle && state.bundle.docStored);
      $('parse').textContent = src === 'docx' ? 'Read document'
        : src === 'url' ? 'Fetch article'
        : state.bundle && state.bundle.imported ? 'Re-import'
        : state.bundle && state.bundle.pictureStatus === 'built-by-hand' ? 'Import'
        : 'Load';
    }

    $('parse').addEventListener('click', function () {
      var source = $('source').value;
      var file = $('file').files[0];
      if (source === 'docx' && !file) return;
      if (source === 'url' && !$('url').value.trim()) return;
      if (source === 'whatsapp' && !(state.bundle && state.bundle.docStored)) return;
      state.fetchImage = null;
      state.bundleKey = null;
      state.placement = null;
      var type = $('type').value;
      $('parse').disabled = true;
      // Imported before, or held back as built by hand: fetch the pictures and
      // run the AI now, through the receiver's re-import.
      var needsImport = source === 'whatsapp' && (state.bundle.imported || state.bundle.pictureStatus === 'built-by-hand');
      $('parse').textContent = 'Working…';
      $('parse-error').style.display = 'none';
      $('type-note').classList.add('wdab-hidden');

      // What is happening, step by step, while the article loads
      var viaReceiverNow = source === 'docx' && $('docx-fetch').checked && !!receiverKey();
      var load = stageList($('load-list'), []
        .concat(needsImport ? [{ id: 'import', label: 'Fetch the pictures again and run the AI checks', detail: 'Usually 1–2 minutes' }] : [])
        .concat(viaReceiverNow ? [{ id: 'receiver', label: 'Fetch the linked pictures and run the AI checks', detail: 'Usually 1–2 minutes' }] : [])
        .concat(source === 'url' ? [{ id: 'fetch', label: 'Fetch the article from topgear.com' }] : [{ id: 'doc', label: 'Read the Word document' }])
        .concat([{ id: 'layout', label: 'Choose the article template' }]));
      $('load').classList.remove('wdab-hidden');

      // 'auto' reads the article type from the document itself, so nobody has
      // to open the Word file first. The editor can still pick one and re-parse.
      function showDetected(d) {
        if (!d) return;
        $('type-note').textContent = 'Template: ' + TYPE_LABELS[d.type] + ' — ' + d.reason;
        $('type-note').classList.remove('wdab-hidden');
      }
      function parseDocHtml(html) {
        load.start('layout');
        if (type === 'auto') { var d = detectArticleType(html); type = d.type; showDetected(d); }
        return type === 'crosshead' ? parseCrosshead(html) : parseNumbered(html, type);
      }

      // Pictures, their order and the AI's picks from a receiver bundle — the
      // same for a WhatsApp article and for a Word doc sent to the receiver.
      function useBundle(bundle) {
        var ordered = orderedBundleImages(bundle);
        state.imageUrls = ordered.map(function (im) { return bundleFileUrl(bundle, im.file); });
        state.placement = { byUrl: {}, heroUrl: null };
        ordered.forEach(function (im, i) {
          state.placement.byUrl[state.imageUrls[i]] = entryOfImage(im, bundle);
          if (bundle.selection && bundle.selection.hero === im.name) state.placement.heroUrl = state.imageUrls[i];
        });
        state.fetchImage = fetchFromReceiver;
        state.bundleKey = bundle.key;
        state.uploadedFilename = bundle.name;
      }

      var pipeline;
      if (source === 'url') {
        var articleUrl = $('url').value.trim();
        try { articleUrl = normaliseArticleUrl(articleUrl); $('url').value = articleUrl; } catch (e) { /* parseFromUrl reports it */ }
        load.start('fetch');
        pipeline = parseFromUrl(articleUrl, type).then(function (r) {
          load.start('layout');
          type = r.type;
          showDetected(r.detected);
          state.imageUrls = r.imageUrls || [];
          // A topgear.com list says which item each picture belongs to: place
          // by entry, and the page's own main picture as the opener (it may
          // also be an item's picture — then it is placed in both).
          if (r.imageEntries && state.imageUrls.length) state.placement = { byUrl: r.imageEntries, heroUrl: state.imageUrls[0] };
          state.uploadedFilename = slugFromUrl(articleUrl);
          return { meta: r.meta, entries: r.entries };
        });
      } else if (source === 'whatsapp') {
        var bundle = state.bundle;
        // Already in Studio: fetch its pictures again (the missing ones may
        // have arrived) and re-run the AI before loading it.
        if (needsImport) load.start('import');
        var ready = needsImport
          ? receiverFetch('/bundles/' + bundle.key + '/reimport', { method: 'POST' })
              .then(function () { return receiverFetch('/bundles/' + bundle.key); })
              .then(function (r) { return r.json(); })
              .then(function (fresh) {
                bundle = state.bundle = fresh;
                $('wa-info').innerHTML = bundleInfoHtml(fresh);
              })
          : Promise.resolve();
        pipeline = ready.then(function () {
          load.start('doc');
          return Promise.all([
            loadMammoth(),
            receiverFetch(bundleFileUrl(bundle, 'doc.docx')).then(function (r) { return r.arrayBuffer(); }),
          ]);
        })
          .then(function (x) { return x[0].convertToHtml({ arrayBuffer: x[1] }); })
          .then(function (result) {
            useBundle(bundle);
            return parseDocHtml(result.value);
          });
      } else {
        // Optionally send the doc to the receiver first: it fetches the Dropbox
        // pictures the doc links to, resizes them and runs the AI checks, as it
        // does for a WhatsApp article. If that fails the doc is still parsed
        // here, without pictures, and the reason is shown — nothing is lost.
        var viaReceiver = $('docx-fetch').checked;
        if (viaReceiver && !receiverKey()) {
          showParseError('Enter the receiver key to fetch pictures, or untick the box.');
          $('load').classList.add('wdab-hidden');
          $('parse').disabled = false; refreshParse();
          return;
        }
        load.start(viaReceiver ? 'receiver' : 'doc');
        var fresh = state.docFresh; state.docFresh = false;
        pipeline = file.arrayBuffer().then(function (buf) {
          var fromReceiver = !viaReceiver ? Promise.resolve(null)
            : receiverFetch('/docs?name=' + encodeURIComponent(file.name) + (fresh ? '&fresh=1' : ''), {
                method: 'POST', body: buf, headers: { 'Content-Type': 'application/octet-stream' },
              }).then(function (r) { return r.json(); })
                .catch(function (e) { return { failed: e.message }; });
          return Promise.all([loadMammoth(), fromReceiver]).then(function (x) {
            load.start('doc');
            var bundle = x[1];
            state.imageUrls = [];
            state.uploadedFilename = file.name.replace(/\.docx$/i, '');
            $('docx-info').classList.add('wdab-hidden');
            if (bundle && bundle.failed) {
              $('docx-info').innerHTML = '<div><strong>Pictures were not fetched.</strong> ' + esc(bundle.failed) + ' The document has been converted without them.</div>';
              $('docx-info').classList.remove('wdab-hidden');
            } else if (bundle) {
              state.bundle = bundle;
              useBundle(bundle);
              $('docx-info').innerHTML = bundleInfoHtml(bundle) +
                '<div>' + (bundle.reused ? 'Using the pictures and AI checks from a few minutes ago. ' : '') +
                '<button type="button" class="wdab-linkbtn" id="' + p + '-docx-again">Fetch again</button></div>';
              $('docx-info').classList.remove('wdab-hidden');
              $('docx-again').addEventListener('click', function () { state.docFresh = true; $('parse').click(); });
            }
            return x[0].convertToHtml({ arrayBuffer: buf });
          });
        }).then(function (result) { return parseDocHtml(result.value); });
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
            $('warn-list').innerHTML = parsed.meta.flagged.map(function (f) {
              var links = f.hrefs.filter(function (h) { return f.text.indexOf(h) === -1; }).map(function (h) {
                return ' → <a href="' + esc(h) + '" target="_blank" rel="noopener">' + esc(h) + '</a>';
              }).join('');
              return '<li>' + esc(f.text) + links + '</li>';
            }).join('');
            $('warn').style.display = 'block';
          } else {
            $('warn').style.display = 'none';
          }

          var imgs = state.imageUrls || [];
          $('images-row').classList.toggle('wdab-hidden', !imgs.length);
          if (imgs.length) {
            var slotCount = 0;
            try { slotCount = ctlSlotCount(); } catch (e) { slotCount = 0; }
            var placeable = Math.min(slotCount, imgs.length);
            $('images-label').textContent = 'Also add ' + imgs.length + ' article image' +
              (imgs.length === 1 ? '' : 's') + ' to this Dossier';
            var note = placeable + ' of ' + imgs.length + ' will be placed in the article' +
              ' (' + slotCount + ' image slot' + (slotCount === 1 ? '' : 's') + ' in this template).';
            if (ctlGalleries.slideshows) {
              note += ' The web galler' + (ctlGalleries.slideshows === 1 ? 'y becomes a slideshow' : 'ies become ' + ctlGalleries.slideshows + ' slideshows') +
                ' (' + ctlGalleries.slides + ' pictures).';
            } else if (imgs.length > slotCount) {
              note += ' The rest are added to the Dossier only' +
                (type === 'crosshead' ? ' — Type 1 or 2 has one image slot per entry.' : '.');
            }
            $('images-progress').textContent = note;
          }

          $('count').textContent = parsed.entries.filter(function (e) { return !e.section; }).length;
          $('entries').innerHTML = parsed.entries.slice(0, 60).map(function (e, i) {
            return type === 'crosshead'
              ? '<div><span class="n">' + (i + 1) + '.</span>' + esc(e.crosshead || e.name || '(no crosshead)') + '</div>'
              : e.section ? '<div><strong>' + esc(e.name) + '</strong> <span class="n">(section heading)</span></div>'
              : (e.sectionBefore ? '<div><strong>' + esc(e.sectionBefore.name) + '</strong> <span class="n">(section heading)</span></div>' : '') +
                '<div><span class="n">[' + e.number + ']</span>' + esc(e.name || e.crosshead || '') + '</div>';
          }).join('') + (parsed.entries.length > 60 ? '<div class="n">… and ' + (parsed.entries.length - 60) + ' more</div>' : '');

          load.finish();
          $('load').classList.add('wdab-hidden');
          setStage(2);
        })
        .catch(function (err) {
          load.fail(brief(err.message, 140));
          $('parse-error').textContent = (source === 'docx' ? 'Could not read the document: ' : '') + err.message;
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

    refreshParse(); // the button names the action for the source showing

    var controller = {
      $: $,
      setStage: setStage,
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
        // created: [{ id, url }] from createImagesInDossier, in upload order.
        // template: the article template read from Studio (loadTemplate);
        // the built-in copy when none is given (the Check stage's preview).
        function build(created, template) {
          template = deepClone(template || TEMPLATES[state.parsedData.type]);
          // Crosshead article from WhatsApp: pictures the AI matched to a
          // section go under that section; the rest become a gallery.
          var sectioned = created && state.bundleKey && state.parsedData.type === 'crosshead'
            ? sectionPictures(state.parsedData.entries, state.bundle) : null;
          var d = state.parsedData.type === 'crosshead'
            ? buildCrosshead(template, meta, sectioned ? sectioned.entries : state.parsedData.entries)
            : buildNumbered(template, meta, state.parsedData.entries, state.parsedData.type);
          var placed = applyImageIds(d, created ? assignSlots(d, created, state.placement) : null);
          // Web galleries → slideshows, pointing at the pictures just created
          placed.galleries = placeGalleries(placed.digital, created);
          placed.comments = 0;
          if ($('comments-add') && $('comments-add').checked) {
            var bundle = state.bundleKey ? state.bundle : null;
            // Frame notes only when pictures were expected (WhatsApp or a web
            // article); a plain Word import has its pictures added by hand later.
            var expectPictures = !!bundle || (state.imageUrls || []).length > 0;
            var opener = expectPictures && created && created.length ? openerNote(placed.digital, created, bundle) : null;
            var notes = flaggedNotes(pm, TOOL_PREFIX)
              .concat(pullQuoteNotes(pm, TOOL_PREFIX))
              .concat(bundle ? bundleNotes(bundle) : [])
              .concat(opener ? [opener] : [])
              .concat(expectPictures ? pictureNotes(placed.digital, bundle) : [])
              .concat(sectioned ? sectioned.notes : []);
            var withComments = addComments(placed.digital, notes, currentUserId());
            placed.digital = withComments.digital;
            placed.comments = withComments.placed.length;
          }
          return placed;
        }
        return {
          digital: build().digital, build: build, type: state.parsedData.type, meta: meta, filename: state.uploadedFilename,
          imageUrls: (state.imageUrls || []),
          fetchImage: state.fetchImage || null,
          bundleKey: state.bundleKey || null,
          addImages: !!($('images-add') && $('images-add').checked && (state.imageUrls || []).length),
        };
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
      // Pictures already uploaded by this dialog, so a retry after a failed
      // article (e.g. a rejected name) reuses them instead of adding a second set.
      var uploaded = { key: null, images: null };

      var content = '<div class="wdab-modal">' + formHtml('wdabm', 'Create article') + '</div>';

      dialogId = ContentStationSdk.openModalDialog({
        title: 'Word & Web → Digital',
        subtitle: 'Creates the digital article in Dossier “' + esc(dossier.Name || '') + '”',
        content: content,
        width: 640,
        buttons: [{ label: 'Close', class: 'pale' }],
      });

      // Digital styles (component set, Look and Feel, Twixl collection) come
      // from BRAND_DEFAULTS. A dossier in any other brand gets an article with
      // no Look and Feel — say so before anything is created.
      resolveDossierContext(dossier).then(function (ctx) {
        var el = document.getElementById('wdabm-brand-warn');
        if (!el || BRAND_DEFAULTS[ctx.pubId]) return;
        el.innerHTML = '<strong>No digital styles for this brand.</strong> This Dossier is in “' +
          esc(ctx.pubName || ('brand ' + ctx.pubId)) + '”, which has no Look and Feel set up in this plug-in, ' +
          'so the article will be created without the Top Gear styles. Use a Dossier in Top Gear for the styled article.';
        el.style.display = 'block';
      }).catch(function () { /* the warning is advisory; creation reports its own errors */ });

      wireForm('wdabm', function (ctl) {
        if (busy) return;
        var result = ctl.getResult();
        if (!result) return;
        busy = true;
        var btn = ctl.$('action');
        var errEl = ctl.$('action-error');
        var noteEl = ctl.$('note');
        btn.disabled = true;
        errEl.style.display = 'none';

        var name = result.meta.title || result.filename;

        // The server's name-length limit varies by install, so step down through
        // NAME_LIMITS rather than hard-coding a guess.
        var tpl = null; // { template, info } from loadTemplate
        function attempt(digital, i) {
          return createArticleInDossier(digital, name, result.meta.feedHeadline, dossier, NAME_LIMITS[i], tpl && tpl.info)
            .catch(function (err) {
              var tooLong = /S1026|too long|invalid characters/i.test(err.message || '');
              if (tooLong && i + 1 < NAME_LIMITS.length) return attempt(digital, i + 1);
              throw err;
            });
        }

        // Images are created first so the article can reference their object
        // IDs and arrive with pictures already in its image slots. Image failure
        // is never fatal — the article is still created, just without them.
        var imagesStep = Promise.resolve(null);
        var imagesKey = result.imageUrls.join('|');
        var nPics = result.imageUrls.length;
        // Stage 3: every step on screen, so a slow one is never mistaken for a hang
        ctl.$('run-title').textContent = 'Creating “' + brief(name, 60) + '”';
        ctl.$('run-foot').classList.add('wdab-hidden');
        ctl.setStage(3);
        var run = stageList(ctl.$('run-list'), []
          .concat(result.addImages ? [
            { id: 'images', label: 'Add ' + nPics + ' picture' + (nPics === 1 ? '' : 's') + ' to the Dossier' },
            { id: 'process', label: 'Wait for Studio to process the pictures' }] : [])
          .concat([{ id: 'template', label: 'Read the article template from Studio' },
            { id: 'article', label: 'Build and create the article' }]));
        if (result.addImages && uploaded.key === imagesKey && uploaded.images && uploaded.images.created.length) {
          imagesStep = Promise.resolve(uploaded.images);
          run.done('images', 'Using the ' + uploaded.images.created.length + ' already added');
        } else if (result.addImages) {
          run.start('images');
          imagesStep = createImagesInDossier(result.imageUrls, dossier, function (done, total) {
            run.progress('images', done, total, Math.min(done + 1, total) + ' of ' + total);
          }, result.fetchImage).catch(function (e) {
            return { created: [], failed: [], fatal: e.message };
          }).then(function (images) {
            if (images && images.created.length) uploaded = { key: imagesKey, images: images };
            if (images && images.fatal) run.failed('images', 'Not added: ' + brief(images.fatal, 100));
            else run.done('images', images ? images.created.length + ' added' + (images.failed.length ? ', ' + images.failed.length + ' failed' : '') : '');
            return images;
          });
        }

        // The Follow/Newsletter assets are shared objects; link them into this
        // dossier so the components resolve. Never fatal.
        var furnitureStep = linkFurnitureToDossier(String(dossier.ID || dossier.Id));

        // Read quietly while the pictures upload (starting its row now would
        // close the pictures' row); reported just before the article is built.
        var templateStep = loadTemplate(result.type, dossier).catch(function (e) {
          return { template: TEMPLATES[result.type], info: { source: 'built-in', error: e.message } };
        }).then(function (t) { tpl = t; return t; });

        Promise.all([furnitureStep, templateStep])
          .then(function () { return imagesStep; })
          .then(function (images) {
            // Don't build the article until Studio has finished ingesting the
            // images, or opening it straight away shows empty placements.
            if (!images || !images.created.length) return images;
            run.start('process');
            var ids0 = images.created.map(function (c) { return c.id; }).filter(Boolean);
            return waitForImagesReady(ids0, function (done, total) {
              run.progress('process', done, total, done + ' of ' + total + ' ready');
            }).then(function (r) {
              images.pending = r.pending;
              run.done('process', r.pending && r.pending.length ? r.pending.length + ' still processing' : 'All ready');
              return images;
            });
          })
          .then(function (images) {
            run.start('template');
            if (tpl.info.source === 'studio') run.done('template', '“' + brief(tpl.info.name, 60) + '”, version ' + tpl.info.version);
            else run.failed('template', 'Used the copy built into the plug-in' + (tpl.info.error ? ': ' + brief(tpl.info.error, 80) : ''));
            run.start('article');
            var created = images ? images.created.filter(function (c) { return c.id; }) : [];
            var built = result.build(created, tpl && tpl.template);
            return attempt(built.digital, 0).then(function (res) {
              return { res: res, images: images, placed: built.filled, slots: built.slots, comments: built.comments };
            });
          })
          .then(function (r) {
            var created = r.res && r.res.Objects && r.res.Objects[0];
            var bmd = created && created.MetaData && created.MetaData.BasicMetaData;
            var newName = bmd ? bmd.Name : name;
            var out = { newName: newName, images: r.images, placed: r.placed, slots: r.slots, comments: r.comments };
            if (!result.bundleKey) return out;
            // Tell the receiver this bundle is in Studio so a re-export never
            // offers it again. Never fatal — the article already exists.
            return receiverFetch('/bundles/' + result.bundleKey + '/imported', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ dossierId: String(dossier.ID || dossier.Id), articleId: bmd ? String(bmd.ID) : null }),
            }).then(function () { return out; }, function (e) {
              console.warn('[word-digital] could not mark WhatsApp bundle imported:', e.message);
              return out;
            });
          })
          .then(function (out) {
            var msg = 'Digital article “' + esc(out.newName) + '” created in Dossier “' + esc(dossier.Name || '') + '”.';
            var im = out.images;
            if (im) {
              if (im.fatal) msg += ' Images were not added: ' + esc(im.fatal);
              else {
                msg += ' ' + im.created.length + ' image' + (im.created.length === 1 ? '' : 's') + ' added';
                if (out.placed) msg += ' (' + out.placed + ' of ' + out.slots + ' slots filled)';
                if (im.failed.length) msg += ', ' + im.failed.length + ' failed';
                msg += '.';
                if (im.pending && im.pending.length) {
                  msg += ' ' + im.pending.length + ' were still processing — give Studio a moment before opening the article.';
                }
                if (im.failed.length) console.warn('[word-digital] image failures:', im.failed);
              }
            }
            if (out.comments) msg += ' ' + out.comments + ' comment' + (out.comments === 1 ? '' : 's') + ' added — see the Comments panel.';
            run.finish();
            ContentStationSdk.showNotification({ content: msg, icon: 'check' });
            try { ContentStationSdk.refreshCurrentSearch(); } catch (e) { /* non-fatal */ }
            if (dialogId !== null) ContentStationSdk.closeModalDialog(dialogId);
          })
          .catch(function (err) {
            // Say what really happened: the pictures are created before the article.
            var kept = uploaded.key === imagesKey && uploaded.images ? uploaded.images.created.length : 0;
            run.fail(brief(err.message, 140));
            ctl.$('run-foot').classList.remove('wdab-hidden');
            errEl.textContent = err.message + '\n' + (kept
              ? 'The article was not created, but ' + kept + ' image' + (kept === 1 ? ' is' : 's are') + ' already in this Dossier. ' +
                'Go back and create again — they will be reused, not uploaded twice.'
              : 'Nothing was created. Go back, fix it and try again.');
            errEl.style.display = 'block';
          })
          .then(function () {
            busy = false;
            btn.disabled = false;
            btn.textContent = 'Create article';
          });

        noteEl.textContent = result.meta.feedHeadline ? 'Feed headline will be saved to C_HEADLINE.' : '';
      });
    },
  });
})();
