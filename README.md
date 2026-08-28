# Word / Web → Digital Article Builder

Converts TG AN+ Word documents (.docx) **or live topgear.com article URLs** into
WoodWing Studio digital articles (.digital).

Word parsing via [mammoth.js](https://github.com/mwilliamson/mammoth.js) (CDN).
The Studio plug-in creates the digital article directly in the current Dossier via
the workflow API.

## Sources

| Source | Needs | Notes |
|--------|-------|-------|
| `.docx` | nothing — runs fully in the browser | Original path, unchanged |
| topgear.com URL | the proxy in `../server.js` | topgear.com sends no CORS headers, and its Akamai edge 403s non-browser user agents, so the fetch must be server-side |

Both sources feed the **same** parsers and builders and produce the same
`{ meta, entries }` shape — `parseFromUrl()` is the URL-side adapter. Article
images (hero + carousel + listicle items + inline) are collected on the URL path
and downloadable as a ZIP.

### Proxy

The URL source needs `server.js` running (locally) or deployed (Fly.io — see
`fly.toml`). The standalone page defaults to same-origin `/proxy`; the Studio
plug-in sets `window.PROXY_BASE` to the deployed proxy because it runs on
Studio's origin. **Set `PROXY_BASE` in `plugin-shell.js` to your real Fly
hostname before shipping.**

## Supported article types

| Type | Layout | Word doc convention |
|------|--------|---------------------|
| 1 — Countdown | Numbered 50 → 1 | Entries as `50. Name` … `1. Name` |
| 2 — Ascending | Numbered 1 → 50 | Entries as `1. Name` … `50. Name` |
| 3 — Crosshead / generic article | Review Q&A or plain prose | Crossheads as bold paragraphs or Word headings; a doc with no crossheads (e.g. "First Look" pieces) becomes one body component per paragraph |

Metadata is picked up from `Feed headline:`, `Article headline:`, `Article subhead:` and `Words:` lines. Bold, italics and links are preserved everywhere, and **each paragraph becomes its own body component**.

**Nothing is ever dropped.** Lines starting with `pics:`, `web gallery`,
`embed`/`imbed`, `embargo` and bare URLs are treated as editorial instructions:
they are listed in the review box *and still placed inline as body text where they
appeared*, so copy can't vanish silently. (Before Aug 2026 these were discarded.)
They are deliberately not read as crossheads even when bold, so an instruction
can't open a spurious section.

Type 1/2 docs written with **bold entry names and no `1. ` prefix** are detected
and auto-numbered by position rather than parsing to nothing.

## Files

- `index.html` — standalone web version (also the single source of truth for the conversion engine)
- `word-digital-plugin.js` — **generated** Content Station SDK plug-in for WoodWing Studio
- `plugin-shell.js` — plug-in UI, Studio API integration and SDK wiring (template for the generated file)
- `build-plugin.js` — extracts the conversion engine from `index.html` and produces `word-digital-plugin.js`

After changing `index.html` **or** `plugin-shell.js`, regenerate and deploy:

```
node build-plugin.js
git add -A && git commit && git push   # GitHub Pages redeploys in ~1 minute
```

Studio users must hard-refresh (Cmd+Shift+R) to pick up a new plug-in version.

## Installing in WoodWing Studio (cloud, self-service)

The plug-in runs inside Studio's own page (same origin), so it needs no CORS changes and no WoodWing involvement.

1. Host this repo on GitHub Pages (or any HTTPS host). Current home: `https://bennstorey.github.io/word-digital-article-builder/`
2. In the Studio Server **Management Console** go to **Integrations → Studio → Plug-ins → Studio** and click **Add new**.
3. Enter the absolute URL of the plug-in file: `https://bennstorey.github.io/word-digital-article-builder/word-digital-plugin.js`, and make sure it is enabled.
4. Refresh Studio.

Note: `{SESSION_ID}` URL-app wildcards are deprecated since Studio 10.40 / Enterprise Server 10.7 — the SDK plug-in route is the supported integration path.

## Workflow

**Happy path (Dossier button):** open a Dossier → click **Word → Digital Article** in the toolbar → pick type + .docx → parse → check metadata → **Create Digital Article in this Dossier**. The article is created in the dossier with:

- the dossier's Brand/Category, its channel/issue Targets, and the brand's first Article workflow status
- Studio object name sanitised (Enterprise rejects `/ \ : * ? " < > |` in names)
- `C_HEADLINE` set from the feed headline (no manual copy needed)
- component set, Look and Feel and Twixl Collection ID from `BRAND_DEFAULTS` in `plugin-shell.js` (Top Gear: Default set, "TG-custom-styles-ISSUE-APPLE 2026", Twixl `102069`), plus `C_CS_FILEFORMATVERSION` / `C_CS_DE_COMPONENT_NAMES`

**Fallback (standalone web version):** `index.html` on GitHub Pages offers the same converter with a `.digital` file download instead of direct creation; feed headline must be copied to C_HEADLINE manually. (The plug-in's Apps-menu entry was removed by request — the Dossier button is the only trigger inside Studio.)

## Studio Server integration notes (hard-won)

- Auth is **cookie-based**: every request needs the `X-WoodWing-Application: Content Station` header (CSRF guard), `credentials: 'same-origin'`, and `Ticket: null` in JSON-RPC payloads. On ticket-based setups `ContentStationSdk.getInfo().Ticket` is used instead (handled automatically).
- File upload: client-generated `fileguid` GUID, `PUT` to `transferindex.php?fileguid=…&ww-app=Content%2BStation&format=<mime>`; the PUT URL doubles as the Attachment `FileUrl` in `CreateObjects`. (The `uploadtokens` flow from the integration guide is not what this Studio version uses.)
- Component set / Look and Feel live in object ExtraMetaData (`C_CS_COMPONENTSET`, `C_CS_STYLEID` — GUIDs), **not** in the `.digital` file.
- SDK docs live on the server: `https://<studio-host>/app/sdk/content-station-11-sdk.md` and `plugins.md`.

## Roadmap

- **Next: use a Word doc already in the Dossier.** In the dossier modal, offer any selected/contained .docx object as the default source (download its native file via the workflow API, parse as usual), with the file picker as the alternative for new uploads. The `onAction(config, selection, dossier)` handler already receives the selection; docx objects have Format `application/vnd.openxmlformats-officedocument.wordprocessingml.document`.
- **Brand-hosted templates instead of embedded ones.** The three article templates are currently baked into the plug-in (extracted from the original `.digitmpl` files). Investigate sourcing them from the brand the dossier belongs to instead: query for `ArticleTemplate` objects (Format `application/ww-digitmpl+json`) in the dossier's Publication, download the native `.digitmpl` same-origin, and build against that. Templates would then be maintained in Studio per brand, with no plug-in redeploy when a template changes.
- Workflow-status picker in the modal (currently: first Article status for the brand).
- Brand defaults for more titles in `BRAND_DEFAULTS`.
