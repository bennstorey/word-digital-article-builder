# Word / Web → Digital Article Builder

Converts TG AN+ Word documents (.docx), **live topgear.com article URLs**, or
**articles sent in the TG AN+ WhatsApp group** into WoodWing Studio digital
articles (.digital).

Word parsing via [mammoth.js](https://github.com/mwilliamson/mammoth.js) (CDN).
The Studio plug-in creates the digital article directly in the current Dossier via
the workflow API.

## Sources

| Source | Needs | Notes |
|--------|-------|-------|
| `.docx` | nothing — runs fully in the browser | Original path, unchanged |
| topgear.com URL | the proxy in `../server.js` | topgear.com sends no CORS headers, and its Akamai edge 403s non-browser user agents, so the fetch must be server-side |
| From WhatsApp | the receiver on the same Fly app, and its key | Plug-in only. See **From WhatsApp** below |

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

## From WhatsApp

The desk's WhatsApp chat export, saved to `/TG WhatsApp exports` in Dropbox, is
picked up by the receiver (`~/Documents/claude-project-web-digital-topgear`,
deployed on the same Fly app as the proxy). The receiver turns every Word doc
from the last 14 days into a **bundle**: the doc, its Dropbox pictures (resized
to max 2000px wide), and Claude's reading of the chat. It never writes to
Studio. The plug-in's **From WhatsApp** source lists the bundles in three
groups:

- **Waiting**: new, with pictures and AI already done.
- **Already drafted — apple.news link posted in the chat**: the chat shows an
  apple.news draft reply for the doc, so the article has already been built.
  Choosing one imports it then (pictures and AI run on demand).
- **Already in Studio — choose to re-import**: fetches the pictures again (e.g.
  once missing ones arrive) and re-runs the AI.

The **Receiver key** field takes the receiver's `RECEIVER_KEY`. It is stored in
the browser's localStorage (demo arrangement; see the receiver README).

**Picture placement.** Numbered lists place by entry: the AI's match, or a
filename such as `15 - F90.jpg`. A missing picture leaves its own frame empty
instead of shifting the rest. **Opener** (header image): the AI's pick of the
picture that best tells the story and works as a centred square crop (Apple
News feed) and as landscape. It is never the picture directly below the header,
but it may reuse another entry's picture. Without a pick, the header is left
empty with a comment.

**Comments.** Created articles carry Digital-editor comments (turn them off
with the checkbox):
- on the text next to each empty picture frame, with the chat's reason where known;
- on the exact phrase for each AI copy query;
- on leftover editor-instruction lines (`PICS:` etc.);
- a summary on the headline (opener, links to fetch by hand, embeds).

Comments attach to text only (format 2.2+: a `comment` attribute on the run plus
a root `comments` object), so a note about a frame sits on the text beside it.
They are labelled "WhatsApp AI:" or "Word → Digital:" and carry the editor's
user id.

## Supported article types

| Type | Layout | Word doc convention |
|------|--------|---------------------|
| 1 — Countdown | Numbered 50 → 1 | Entries as `50. Name` … `1. Name` |
| 2 — Ascending | Numbered 1 → 50 | Entries as `1. Name` … `50. Name` |
| 3 — Crosshead / generic article | Review Q&A or plain prose | Crossheads as bold paragraphs or Word headings; a doc with no crossheads (e.g. "First Look" pieces) becomes one body component per paragraph |

**Auto-detect** (the default) reads the type from the document:
- three or more `N. Name` entries give countdown or ascending, by the direction most steps take;
- Word's automatic numbering gives ascending;
- anything else gives crosshead.

The dialog says what it found ("Detected: Type 2 — 18 numbered entries, 1 → 18"),
and picking a type and parsing again overrides it.

**Word's automatic numbering** (the number is Word list formatting, not typed)
is read as a numbered list. mammoth drops the numbers and emits each entry as
its own `<ol><li>` before the entry's copy; three or more of those make the
entries, numbered by position. This fixed "51 Worst Cars" and "Ugliest F1 cars",
which previously parsed to 2 and 0 entries.

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

**Happy path (Dossier button):** open a Dossier → click **Word & Web → Digital** in the toolbar → pick the source (.docx, topgear.com URL or From WhatsApp; the article type auto-detects) → parse → check metadata → **Create Digital Article in this Dossier**. The article is created in the dossier with:

- the dossier's Brand/Category, its channel/issue Targets, and the brand's first Article workflow status
- Studio object name sanitised (Enterprise rejects `/ \ : * ? " < > |` in names)
- `C_HEADLINE` set from the feed headline (no manual copy needed)
- component set, Look and Feel and Twixl Collection ID from `BRAND_DEFAULTS` in `plugin-shell.js` (Top Gear: Default set, "TG-custom-styles-ISSUE-APPLE 2026", Twixl `102069`), plus `C_CS_FILEFORMATVERSION` / `C_CS_DE_COMPONENT_NAMES`. **Only Top Gear (brand 3) has defaults.** In any other brand (e.g. the lab's *WW Development Sandbox*) the article has no Look and Feel, and the dialog warns before anything is created

**Fallback (standalone web version):** `index.html` on GitHub Pages offers the same converter with a `.digital` file download instead of direct creation; feed headline must be copied to C_HEADLINE manually. (The plug-in's Apps-menu entry was removed by request — the Dossier button is the only trigger inside Studio.)

## Studio Server integration notes (hard-won)

- Auth is **cookie-based**: every request needs the `X-WoodWing-Application: Content Station` header (CSRF guard), `credentials: 'same-origin'`, and `Ticket: null` in JSON-RPC payloads. On ticket-based setups `ContentStationSdk.getInfo().Ticket` is used instead (handled automatically).
- File upload: client-generated `fileguid` GUID, `PUT` to `transferindex.php?fileguid=…&ww-app=Content%2BStation&format=<mime>`; the PUT URL doubles as the Attachment `FileUrl` in `CreateObjects`. (The `uploadtokens` flow from the integration guide is not what this Studio version uses.)
- Component set / Look and Feel live in object ExtraMetaData (`C_CS_COMPONENTSET`, `C_CS_STYLEID` — GUIDs), **not** in the `.digital` file.
- SDK docs live on the server: `https://<studio-host>/app/sdk/content-station-11-sdk.md` and `plugins.md`.

## Roadmap

- **Snippet-list template** for unnumbered lists (bold headings, no numbers, e.g. "50 silliest American cars"), hooked into auto-detect. See KNOWN-ISSUES §1.
- **Comments in the web version**: the comment functions are in the shared engine, but only the plug-in calls them.

- **Next: use a Word doc already in the Dossier.** In the dossier modal, offer any selected/contained .docx object as the default source (download its native file via the workflow API, parse as usual), with the file picker as the alternative for new uploads. The `onAction(config, selection, dossier)` handler already receives the selection; docx objects have Format `application/vnd.openxmlformats-officedocument.wordprocessingml.document`.
- **Brand-hosted templates instead of embedded ones.** The three article templates are currently baked into the plug-in (extracted from the original `.digitmpl` files). Investigate sourcing them from the brand the dossier belongs to instead: query for `ArticleTemplate` objects (Format `application/ww-digitmpl+json`) in the dossier's Publication, download the native `.digitmpl` same-origin, and build against that. Templates would then be maintained in Studio per brand, with no plug-in redeploy when a template changes.
- Workflow-status picker in the modal (currently: first Article status for the brand).
- Brand defaults for more titles in `BRAND_DEFAULTS`.
