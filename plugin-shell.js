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

  // Object names may not contain the characters Enterprise rejects (/ \ : * ? " < > |).
  // Only the Studio object name is sanitised — headlines inside the article keep them.
  // Enterprise rejects / \ : * ? " < > | and enforces a name-length limit that
  // varies by install — 60 chars has been seen to fail with S1026. Cut on a word
  // boundary so a shortened name still reads sensibly.
  function sanitizeObjectName(name, maxLen) {
    var out = String(name)
      .replace(/[\/\\:*?"<>|]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    var limit = maxLen || NAME_LIMITS[0];
    if (out.length > limit) {
      out = out.slice(0, limit);
      var cut = out.lastIndexOf(' ');
      if (cut > limit * 0.5) out = out.slice(0, cut);
    }
    return out.trim();
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

  // Create the digital article inside the given dossier.
  // Publication/Category are taken from the dossier; Targets are copied from
  // the dossier so the article lands on the same channel/issue. Component set,
  // Look and Feel and Twixl collection come from BRAND_DEFAULTS.
  function createArticleInDossier(digital, name, feedHeadline, dossier, nameLimit) {
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
      if (r.status === 401) throw new Error('The WhatsApp receiver refused the key — check it and try again.');
      if (!r.ok) throw new Error('WhatsApp receiver: HTTP ' + r.status);
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
    'built-by-hand': 'Built by hand in WhatsApp',
    'doc-missing': 'Word doc missing from the export',
  };

  function bundleInfoHtml(b) {
    var out = [];
    out.push('<strong>' + esc(PICTURE_STATUS[b.pictureStatus] || b.pictureStatus) + '</strong>' +
      (b.images && b.images.length ? ' — ' + b.images.length + ' image' + (b.images.length === 1 ? '' : 's') : '') +
      ' · sent by ' + esc(b.sender) + ' ' + esc(b.ts.replace('T', ' ')));
    if (b.statusNote) out.push(esc(b.statusNote));
    if (b.imported) {
      out.push('Already imported into Studio on ' + esc(String(b.imported.at || '').slice(0, 10)) +
        '. Loading it re-imports: the pictures are fetched from Dropbox again (so any added since are included) and the AI runs again.');
    } else if (b.builtByHand) {
      var h = b.builtByHand;
      out.push('Looks built by hand already: ' + esc(h.ts.replace('T', ' ').slice(0, 16)) + ' — “' + esc(h.text.slice(0, 80)) + '” ' +
        '<a href="' + esc(h.href) + '" target="_blank" rel="noopener">open draft</a>' +
        (h.rule === 'order' ? ' (matched by timing, not by name — check it’s this article)' : '') + '.');
      if (b.pictureStatus === 'built-by-hand') out.push('Loading it imports it anyway: the pictures are fetched from Dropbox and the AI runs then.');
    }
    (b.pictures || []).forEach(function (p) {
      if (p.type !== 'dropbox-folder' && p.type !== 'dropbox-file') {
        out.push('Fetch by hand: <a href="' + esc(p.href) + '" target="_blank" rel="noopener">' + esc(p.type) + ' link</a>' +
          (p.reason ? ' (' + esc(p.reason) + ')' : ''));
      } else if (p.from === 'chat') {
        out.push('Pictures from the chat: ' + esc(p.reason || ''));
      }
    });
    (b.suggestedPictures || []).forEach(function (p) {
      out.push('Possibly for this article (AI, ' + Math.round(p.confidence * 100) + '%): <a href="' + esc(p.href) + '" target="_blank" rel="noopener">link</a> — ' + esc(p.reason));
    });
    if (b.selection) {
      out.push(b.selection.hero
        ? 'AI picked the opener: ' + esc(b.selection.hero) + (b.selection.heroReason ? ' — ' + esc(b.selection.heroReason) : '')
        : 'AI found no picture that works as the opener' + (b.selection.heroReason ? ': ' + esc(b.selection.heroReason) : '.'));
    }
    var ins = b.instructions;
    if (ins) {
      (ins.missing || []).forEach(function (m) { out.push('Missing pictures: ' + esc(m.entry) + ' — ' + esc(m.note)); });
      (ins.embeds || []).forEach(function (e) { out.push('Embed: ' + esc(e.url) + ' (' + esc(e.where) + ')'); });
      if (ins.embargo) out.push('Embargo: ' + esc(ins.embargo));
      (ins.otherNotes || []).forEach(function (n) { out.push(esc(n)); });
    }
    (b.errors || []).forEach(function (e) { out.push('⚠ ' + esc(e.step) + ': ' + esc(e.message)); });
    return out.map(function (l) { return '<div>' + l + '</div>'; }).join('');
  }

  // ─── Comments in the created article ──────────────────────────────────────
  // Notes become Digital editor comments (see addComments in the engine). They
  // carry the creating editor's user id; the prefix says who raised them.
  var AI_PREFIX = 'WhatsApp AI: ';
  var TOOL_PREFIX = 'Word → Digital: ';

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
        text: AI_PREFIX + q.comment + (q.source && q.source !== 'the copy itself' ? '\n(' + q.source + ')' : ''),
      });
    });
    var head = [];
    // (The opener note is written after placement — see openerNote.)
    (bundle.pictures || []).forEach(function (p) {
      if (p.type !== 'dropbox-folder' && p.type !== 'dropbox-file') head.push('Pictures to fetch by hand (' + p.type + '): ' + p.href);
    });
    (bundle.suggestedPictures || []).forEach(function (p) {
      head.push('Possibly this article\'s pictures (' + Math.round(p.confidence * 100) + '% sure): ' + p.href + ' — ' + p.reason);
    });
    var ins = bundle.instructions;
    if (ins) {
      (ins.embeds || []).forEach(function (e) { head.push('Embed requested: ' + e.url + ' (' + e.where + ')'); });
      if (ins.embargo) head.push('Embargo: ' + ins.embargo);
      (ins.otherNotes || []).forEach(function (n) { head.push(n); });
    }
    if (head.length) notes.push({ anchor: {}, text: AI_PREFIX + head.join('\n') });
    return notes;
  }

  // Which numbered entry a bundle image shows: the AI's match first, then a
  // filename like "15-F90.jpg" or "15.jpg". Camera names such as
  // "03.02.2026-Geely…" are not read as entry 3 (digit after the separator).
  function entryOfImage(im, bundle) {
    var r = bundle.selection && (bundle.selection.ranked || []).filter(function (x) { return x.file === im.name; })[0];
    var fromAi = r && r.entry != null && String(r.entry).match(/\d+/);
    if (fromAi) return Number(fromAi[0]);
    var m = im.name.match(/^0*(\d{1,3})(?:[\s._-]+(?!\d)|\.[a-z]+$)/i);
    return m ? Number(m[1]) : null;
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
      slots.push({ header: c.identifier === 'header-image', entry: num ? Number(num) : null });
    });
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
      var why = sel && sel.hero === name && sel.heroReason ? ' — ' + sel.heroReason : '';
      return { anchor: {}, text: (why ? AI_PREFIX : TOOL_PREFIX) + 'Opening picture: ' + name + why +
        '. All ' + total + ' pictures are in the Dossier if you prefer another.' };
    }
    if (sel && sel.hero) {
      return { anchor: {}, text: AI_PREFIX + 'Suggested ' + sel.hero + ' as the opener, but it sits directly below the header, so it was not used.' };
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
        ? ' The writer\'s doc says: ' + ((bundle.doc && bundle.doc.picLineText) || 'no picture link given') + '.'
        : '';
      notes = [{ anchor: {}, text: TOOL_PREFIX + 'No pictures placed yet — all ' + frames + ' picture frames are empty.' + why }];
    }
    var emptyEntries = notes.map(function (n) { return n.text; }).join('\n');
    missing.forEach(function (m) {
      if (emptyEntries.indexOf(m.note) !== -1) return; // already said on the empty frame
      notes.push({ anchor: { entry: m.entry }, text: AI_PREFIX + 'The chat says this picture was missing (' + m.note + '). Check the picture placed here.' });
    });
    return notes;
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
    '.wdab .wdab-wa-info{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:8px 12px;margin-top:8px;font-size:12px;color:#334155}',
    '.wdab .wdab-wa-info div+div{margin-top:4px}',
    '.wdab select,.wdab input[type=text],.wdab input[type=password]{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:8px;padding:7px 10px;font:inherit;color:#1e293b;background:#fff}',
    '.wdab input[type=file]{width:100%;font:inherit}',
    '.wdab input[type=checkbox]{width:auto;margin:0 6px 0 0;vertical-align:middle}',
    '.wdab-row label input[type=checkbox]+span{font-weight:400;color:#334155}',
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

  function formHtml(p, actionLabel) {
    return '<div class="wdab">' +
      '  <div class="wdab-card">' +
      '    <h2>1 — Set up</h2>' +
      '    <div class="wdab-warn" id="' + p + '-brand-warn"></div>' +
      '    <div class="wdab-row">' +
      '      <label for="' + p + '-type">Article type</label>' +
      '      <select id="' + p + '-type">' +
      '        <option value="auto" selected>Auto-detect from the document</option>' +
      '        <option value="countdown">Type 1 — Numbered countdown (50 → 1)</option>' +
      '        <option value="ascending">Type 2 — Numbered ascending (1 → 50)</option>' +
      '        <option value="crosshead">Type 3 — Crosshead / generic article</option>' +
      '      </select>' +
      '      <p class="wdab-note wdab-hidden" id="' + p + '-type-note"></p>' +
      '    </div>' +
      '    <div class="wdab-row">' +
      '      <label for="' + p + '-source">Source</label>' +
      '      <select id="' + p + '-source">' +
      '        <option value="docx">Word document (.docx)</option>' +
      '        <option value="url">TopGear article URL</option>' +
      '        <option value="whatsapp">From WhatsApp</option>' +
      '      </select>' +
      '    </div>' +
      '    <div class="wdab-row wdab-hidden" id="' + p + '-wa-row">' +
      '      <label for="' + p + '-wa-key">Receiver key</label>' +
      '      <input type="password" id="' + p + '-wa-key" autocomplete="off">' +
      '      <label for="' + p + '-wa-bundle" style="margin-top:10px">Article from WhatsApp</label>' +
      '      <select id="' + p + '-wa-bundle"><option value="">—</option></select>' +
      '      <div class="wdab-wa-info wdab-hidden" id="' + p + '-wa-info"></div>' +
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
      '      <div class="wdab-row wdab-hidden" id="' + p + '-images-row">' +
      '        <label><input type="checkbox" id="' + p + '-images-add" checked> <span id="' + p + '-images-label"></span></label>' +
      '        <p class="wdab-note" id="' + p + '-images-progress"></p>' +
      '      </div>' +
      '      <div class="wdab-row">' +
      '        <label><input type="checkbox" id="' + p + '-comments-add" checked> <span>Add comments to the article for missing pictures and things to check</span></label>' +
      '        <p class="wdab-note" id="' + p + '-comments-note"></p>' +
      '      </div>' +
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
    $('wa-key').addEventListener('change', function () {
      setReceiverKey($('wa-key').value.trim());
      loadBundles();
    });

    function showParseError(msg) {
      $('parse-error').textContent = msg;
      $('parse-error').style.display = msg ? 'block' : 'none';
    }

    function loadBundles() {
      state.bundle = null;
      $('wa-info').classList.add('wdab-hidden');
      $('wa-bundle').innerHTML = '<option value="">Loading…</option>';
      refreshParse();
      receiverFetch('/bundles?all=1')
        .then(function (r) { return r.json(); })
        .then(function (j) {
          showParseError('');
          var list = j.bundles || [];
          // New first; then docs the chat shows were already built by hand
          // (BRS replied with an apple.news draft); then ones imported here.
          var waiting = list.filter(function (b) { return !b.imported && !b.builtByHand; });
          var byHand = list.filter(function (b) { return !b.imported && b.builtByHand; });
          var done = list.filter(function (b) { return b.imported; });
          var day = function (ts) { return String(ts || '').slice(0, 10); };
          var opt = function (b) {
            var label = b.imported ? 'imported ' + day(b.imported.at)
              : b.builtByHand ? 'draft posted ' + day(b.builtByHand.ts)
              : (PICTURE_STATUS[b.pictureStatus] || b.pictureStatus);
            return '<option value="' + esc(b.key) + '">' + esc(b.name) + ' — ' + esc(label) + '</option>';
          };
          var group = function (label, items) {
            return items.length ? '<optgroup label="' + esc(label) + '">' + items.map(opt).join('') + '</optgroup>' : '';
          };
          $('wa-bundle').innerHTML =
            '<option value="">' + (waiting.length ? 'Choose an article…' : 'Nothing new from WhatsApp') + '</option>' +
            group('Waiting', waiting) +
            group('Built by hand in WhatsApp — choose to import anyway', byHand) +
            group('Already in Studio — choose to re-import', done);
        })
        .catch(function (e) {
          $('wa-bundle').innerHTML = '<option value="">—</option>';
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
    function ctlSlotCount() {
      if (!state.parsedData) return 0;
      var t = deepClone(TEMPLATES[state.parsedData.type]);
      var d = state.parsedData.type === 'crosshead'
        ? buildCrosshead(t, state.parsedData.meta, state.parsedData.entries)
        : buildNumbered(t, state.parsedData.meta, state.parsedData.entries, state.parsedData.type);
      return (d.data.content || []).filter(function (c) {
        return c.identifier === 'image' || c.identifier === 'header-image';
      }).length;
    }

    function refreshParse() {
      var src = $('source').value;
      $('parse').disabled = src === 'docx' ? !$('file').files.length
        : src === 'url' ? !$('url').value.trim()
        : !(state.bundle && state.bundle.docStored);
      $('parse').textContent = src === 'docx' ? 'Parse Document'
        : src === 'url' ? 'Fetch & Parse Article'
        : state.bundle && state.bundle.imported ? 'Re-import from Dropbox & Load'
        : state.bundle && state.bundle.pictureStatus === 'built-by-hand' ? 'Import from Dropbox & Load'
        : 'Load from WhatsApp';
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
      $('parse').textContent = source === 'docx' ? 'Parsing…'
        : needsImport ? 'Importing from Dropbox…' : 'Fetching…';
      $('parse-error').style.display = 'none';
      $('type-note').classList.add('wdab-hidden');

      // 'auto' reads the article type from the document itself, so nobody has
      // to open the Word file first. The editor can still pick one and re-parse.
      function showDetected(d) {
        if (!d) return;
        $('type-note').textContent = 'Detected: ' + TYPE_LABELS[d.type] + ' (' + d.reason + '). ' +
          'Pick a type above and parse again to override.';
        $('type-note').classList.remove('wdab-hidden');
      }
      function parseDocHtml(html) {
        if (type === 'auto') { var d = detectArticleType(html); type = d.type; showDetected(d); }
        return type === 'crosshead' ? parseCrosshead(html) : parseNumbered(html, type);
      }

      var pipeline;
      if (source === 'url') {
        var articleUrl = $('url').value.trim();
        pipeline = parseFromUrl(articleUrl, type).then(function (r) {
          type = r.type;
          showDetected(r.detected);
          state.imageUrls = r.imageUrls || [];
          state.uploadedFilename = slugFromUrl(articleUrl);
          return { meta: r.meta, entries: r.entries };
        });
      } else if (source === 'whatsapp') {
        var bundle = state.bundle;
        // Already in Studio: fetch its pictures again (the missing ones may
        // have arrived) and re-run the AI before loading it.
        var ready = needsImport
          ? receiverFetch('/bundles/' + bundle.key + '/reimport', { method: 'POST' })
              .then(function () { return receiverFetch('/bundles/' + bundle.key); })
              .then(function (r) { return r.json(); })
              .then(function (fresh) {
                bundle = state.bundle = fresh;
                $('wa-info').innerHTML = bundleInfoHtml(fresh);
                $('parse').textContent = 'Fetching…';
              })
          : Promise.resolve();
        pipeline = ready.then(function () {
          return Promise.all([
            loadMammoth(),
            receiverFetch(bundleFileUrl(bundle, 'doc.docx')).then(function (r) { return r.arrayBuffer(); }),
          ]);
        })
          .then(function (x) { return x[0].convertToHtml({ arrayBuffer: x[1] }); })
          .then(function (result) {
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
            return parseDocHtml(result.value);
          });
      } else {
        pipeline = loadMammoth()
          .then(function (mammoth) { return file.arrayBuffer().then(function (buf) { return mammoth.convertToHtml({ arrayBuffer: buf }); }); })
          .then(function (result) {
            state.imageUrls = [];
            state.uploadedFilename = file.name.replace(/\.docx$/i, '');
            return parseDocHtml(result.value);
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
              ' (' + slotCount + ' image slot' + (slotCount === 1 ? '' : 's') + ' in this layout).';
            if (imgs.length > slotCount) {
              note += ' The rest are added to the Dossier only' +
                (type === 'crosshead' ? ' — Type 1 or 2 has one image slot per entry.' : '.');
            }
            $('images-progress').textContent = note;
          }

          $('count').textContent = parsed.entries.length;
          $('entries').innerHTML = parsed.entries.slice(0, 60).map(function (e, i) {
            return type === 'crosshead'
              ? '<div><span class="n">' + (i + 1) + '.</span>' + esc(e.crosshead || e.name || '(no crosshead)') + '</div>'
              : '<div><span class="n">[' + e.number + ']</span>' + esc(e.name || e.crosshead || '') + '</div>';
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
        // created: [{ id, url }] from createImagesInDossier, in upload order.
        function build(created) {
          var template = deepClone(TEMPLATES[state.parsedData.type]);
          var d = state.parsedData.type === 'crosshead'
            ? buildCrosshead(template, meta, state.parsedData.entries)
            : buildNumbered(template, meta, state.parsedData.entries, state.parsedData.type);
          var placed = applyImageIds(d, created ? assignSlots(d, created, state.placement) : null);
          placed.comments = 0;
          if ($('comments-add') && $('comments-add').checked) {
            var bundle = state.bundleKey ? state.bundle : null;
            // Frame notes only when pictures were expected (WhatsApp or a web
            // article); a plain Word import has its pictures added by hand later.
            var expectPictures = !!bundle || (state.imageUrls || []).length > 0;
            var opener = expectPictures && created && created.length ? openerNote(placed.digital, created, bundle) : null;
            var notes = flaggedNotes(pm, TOOL_PREFIX)
              .concat(bundle ? bundleNotes(bundle) : [])
              .concat(opener ? [opener] : [])
              .concat(expectPictures ? pictureNotes(placed.digital, bundle) : []);
            var withComments = addComments(placed.digital, notes, currentUserId());
            placed.digital = withComments.digital;
            placed.comments = withComments.placed.length;
          }
          return placed;
        }
        return {
          digital: build().digital, build: build, meta: meta, filename: state.uploadedFilename,
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

      var content = '<div class="wdab-modal">' + formHtml('wdabm', 'Create Digital Article in this Dossier') + '</div>';

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
        btn.textContent = 'Creating…';
        errEl.style.display = 'none';

        var name = result.meta.title || result.filename;

        // The server's name-length limit varies by install, so step down through
        // NAME_LIMITS rather than hard-coding a guess.
        function attempt(digital, i) {
          return createArticleInDossier(digital, name, result.meta.feedHeadline, dossier, NAME_LIMITS[i])
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
        if (result.addImages) {
          btn.textContent = 'Adding images…';
          var progEl = ctl.$('images-progress');
          imagesStep = createImagesInDossier(result.imageUrls, dossier, function (done, total) {
            if (progEl) progEl.textContent = 'Uploading image ' + Math.min(done + 1, total) + ' of ' + total + '…';
          }, result.fetchImage).catch(function (e) {
            return { created: [], failed: [], fatal: e.message };
          });
        }

        // The Follow/Newsletter assets are shared objects; link them into this
        // dossier so the components resolve. Never fatal.
        var furnitureStep = linkFurnitureToDossier(String(dossier.ID || dossier.Id));

        furnitureStep
          .then(function () { return imagesStep; })
          .then(function (images) {
            // Don't build the article until Studio has finished ingesting the
            // images, or opening it straight away shows empty placements.
            if (!images || !images.created.length) return images;
            btn.textContent = 'Processing images…';
            var progEl2 = ctl.$('images-progress');
            var ids0 = images.created.map(function (c) { return c.id; }).filter(Boolean);
            return waitForImagesReady(ids0, function (done, total) {
              if (progEl2) progEl2.textContent = 'Waiting for Studio to process images (' + done + ' of ' + total + ')…';
            }).then(function (r) {
              images.pending = r.pending;
              return images;
            });
          })
          .then(function (images) {
            btn.textContent = 'Creating…';
            var created = images ? images.created.filter(function (c) { return c.id; }) : [];
            var built = result.build(created);
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
            ContentStationSdk.showNotification({ content: msg, icon: 'check' });
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
