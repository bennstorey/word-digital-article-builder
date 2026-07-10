/**
 * Word → Digital Article Builder — WoodWing Studio plug-in
 *
 * Content Station SDK plug-in that adds a Custom App to Studio's Apps menu
 * for converting Top Gear AN+ Word documents (.docx) into digital articles
 * (.digital).
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
  /*__ENGINE__*/
  // ─── End conversion engine ────────────────────────────────────────────────

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

  // ─── UI ────────────────────────────────────────────────────────────────────
  var CSS = [
    '.wdab{max-width:640px;margin:0 auto;padding:24px 16px;font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;color:#1e293b}',
    '.wdab h2{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;color:#64748b;margin:0 0 14px}',
    '.wdab .wdab-card{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:20px;margin-bottom:16px;box-shadow:0 1px 2px rgba(0,0,0,.04)}',
    '.wdab label{display:block;font-weight:500;color:#334155;margin:0 0 4px}',
    '.wdab select,.wdab input[type=text]{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:8px;padding:7px 10px;font:inherit;color:#1e293b;background:#fff}',
    '.wdab input[type=file]{width:100%;font:inherit}',
    '.wdab .wdab-row{margin-bottom:14px}',
    '.wdab button{display:inline-block;border:0;border-radius:8px;padding:9px 16px;font:inherit;font-weight:600;cursor:pointer;background:#2563eb;color:#fff;width:100%}',
    '.wdab button:disabled{opacity:.4;cursor:not-allowed}',
    '.wdab .wdab-feed input{border-color:#fcd34d;background:#fffbeb}',
    '.wdab .wdab-feed-note{display:inline-block;font-size:11px;color:#b45309;background:#fffbeb;border:1px solid #fde68a;border-radius:4px;padding:1px 6px;margin-left:8px;font-weight:400}',
    '.wdab .wdab-error{color:#dc2626;margin-top:8px;display:none}',
    '.wdab .wdab-warn{background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:10px 12px;margin-bottom:14px;display:none}',
    '.wdab .wdab-warn ul{margin:6px 0 0;padding-left:18px;color:#92400e}',
    '.wdab .wdab-entries{max-height:280px;overflow-y:auto;border:1px solid #e2e8f0;border-radius:8px;padding:8px 12px;color:#334155}',
    '.wdab .wdab-entries .n{color:#94a3b8;margin-right:6px}',
    '.wdab .wdab-hidden{display:none}'
  ].join('\n');

  var HTML =
    '<style>' + CSS + '</style>' +
    '<div class="wdab">' +
    '  <div class="wdab-card">' +
    '    <h2>1 — Set up</h2>' +
    '    <div class="wdab-row">' +
    '      <label for="wdab-type">Article type</label>' +
    '      <select id="wdab-type">' +
    '        <option value="countdown">Type 1 — Numbered countdown (50 → 1)</option>' +
    '        <option value="ascending">Type 2 — Numbered ascending (1 → 50)</option>' +
    '        <option value="crosshead">Type 3 — Crosshead + text (reviews)</option>' +
    '      </select>' +
    '    </div>' +
    '    <div class="wdab-row">' +
    '      <label for="wdab-file">Word document (.docx)</label>' +
    '      <input type="file" id="wdab-file" accept=".docx">' +
    '    </div>' +
    '    <button id="wdab-parse" disabled>Parse Document</button>' +
    '    <p class="wdab-error" id="wdab-parse-error"></p>' +
    '  </div>' +
    '  <div id="wdab-preview" class="wdab-hidden">' +
    '    <div class="wdab-card">' +
    '      <h2>2 — Detected metadata (editable)</h2>' +
    '      <div class="wdab-row wdab-feed wdab-hidden" id="wdab-feed-row">' +
    '        <label>Feed Headline<span class="wdab-feed-note">Copy manually to Studio → C_HEADLINE</span></label>' +
    '        <input type="text" id="wdab-feed">' +
    '      </div>' +
    '      <div class="wdab-row"><label>Article title</label><input type="text" id="wdab-title"></div>' +
    '      <div class="wdab-row"><label>Subtitle</label><input type="text" id="wdab-subtitle"></div>' +
    '      <div class="wdab-row"><label>Author name</label><input type="text" id="wdab-author"></div>' +
    '      <div class="wdab-warn" id="wdab-warn"><strong>Unrecognised lines (not included in output):</strong><ul id="wdab-warn-list"></ul></div>' +
    '    </div>' +
    '    <div class="wdab-card">' +
    '      <h2>3 — Entries (<span id="wdab-count">0</span>)</h2>' +
    '      <div class="wdab-entries" id="wdab-entries"></div>' +
    '    </div>' +
    '    <div class="wdab-card">' +
    '      <h2>4 — Download</h2>' +
    '      <button id="wdab-download">Download .digital file</button>' +
    '    </div>' +
    '  </div>' +
    '</div>';

  function esc(str) {
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function wireApp() {
    var $ = function (id) { return document.getElementById(id); };
    var fileInput = $('wdab-file');
    var parseBtn = $('wdab-parse');
    var parseError = $('wdab-parse-error');
    var preview = $('wdab-preview');

    var parsedData = null;
    var uploadedFilename = '';

    fileInput.addEventListener('change', function () {
      parseBtn.disabled = !fileInput.files.length;
      preview.classList.add('wdab-hidden');
      parsedData = null;
      parseError.style.display = 'none';
    });

    parseBtn.addEventListener('click', function () {
      var file = fileInput.files[0];
      if (!file) return;
      var type = $('wdab-type').value;
      parseBtn.disabled = true;
      parseBtn.textContent = 'Parsing…';
      parseError.style.display = 'none';

      loadMammoth()
        .then(function (mammoth) { return file.arrayBuffer().then(function (buf) { return mammoth.convertToHtml({ arrayBuffer: buf }); }); })
        .then(function (result) {
          var parsed = type === 'crosshead' ? parseCrosshead(result.value) : parseNumbered(result.value);
          parsedData = { meta: parsed.meta, entries: parsed.entries, type: type };
          uploadedFilename = file.name.replace(/\.docx$/i, '');

          $('wdab-feed').value = parsed.meta.feedHeadline;
          $('wdab-feed-row').classList.toggle('wdab-hidden', !parsed.meta.feedHeadline);
          $('wdab-title').value = parsed.meta.title;
          $('wdab-subtitle').value = parsed.meta.subtitle;
          $('wdab-author').value = parsed.meta.author;

          var warn = $('wdab-warn');
          if (parsed.meta.unrecognized && parsed.meta.unrecognized.length) {
            $('wdab-warn-list').innerHTML = parsed.meta.unrecognized.map(function (t) { return '<li>' + esc(t) + '</li>'; }).join('');
            warn.style.display = 'block';
          } else {
            warn.style.display = 'none';
          }

          $('wdab-count').textContent = parsed.entries.length;
          $('wdab-entries').innerHTML = parsed.entries.slice(0, 60).map(function (e, i) {
            return type === 'crosshead'
              ? '<div><span class="n">' + (i + 1) + '.</span>' + esc(e.crosshead || '(no crosshead)') + '</div>'
              : '<div><span class="n">[' + e.number + ']</span>' + esc(e.name) + '</div>';
          }).join('') + (parsed.entries.length > 60 ? '<div class="n">… and ' + (parsed.entries.length - 60) + ' more</div>' : '');

          preview.classList.remove('wdab-hidden');
        })
        .catch(function (err) {
          parseError.textContent = 'Error parsing document: ' + err.message;
          parseError.style.display = 'block';
        })
        .then(function () {
          parseBtn.disabled = false;
          parseBtn.textContent = 'Parse Document';
        });
    });

    $('wdab-download').addEventListener('click', function () {
      if (!parsedData) return;
      var pm = parsedData.meta;
      var pd = pm.deltas || {};
      var meta = {
        feedHeadline: $('wdab-feed').value,
        title: $('wdab-title').value,
        subtitle: $('wdab-subtitle').value,
        author: $('wdab-author').value,
        score: pm.score,
        intro: pm.intro,
        introParts: pm.introParts,
        titleDeltas: $('wdab-title').value === pm.title ? pd.title : null,
        subtitleDeltas: $('wdab-subtitle').value === pm.subtitle ? pd.subtitle : null,
      };
      var template = deepClone(TEMPLATES[parsedData.type]);
      var digital = parsedData.type === 'crosshead'
        ? buildCrosshead(template, meta, parsedData.entries)
        : buildNumbered(template, meta, parsedData.entries, parsedData.type);

      var blob = new Blob([JSON.stringify(digital)], { type: 'application/json' });
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = uploadedFilename + '.digital';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    });
  }

  ContentStationSdk.registerCustomApp({
    name: 'word-digital-article-builder',
    title: 'Word → Digital Article',
    content: HTML,
    onInit: wireApp,
  });
})();
