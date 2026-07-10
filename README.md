# Word → Digital Article Builder

Converts Top Gear AN+ Word documents (.docx) into WoodWing Studio digital articles (.digital).

Runs entirely in the browser — no backend. Word parsing via [mammoth.js](https://github.com/mwilliamson/mammoth.js), UI via Tailwind CSS (both loaded from CDN).

## Supported article types

| Type | Layout | Word doc convention |
|------|--------|---------------------|
| 1 — Countdown | Numbered 50 → 1 | Entries as `50. Name` … `1. Name` |
| 2 — Ascending | Numbered 1 → 50 | Entries as `1. Name` … `50. Name` |
| 3 — Crosshead | Review Q&A | Crossheads as bold paragraphs or Word headings |

Metadata is picked up from `Feed headline:`, `Article headline:`, `Article subhead:` and `Words:` lines. Bold in body text and italics everywhere are preserved. Lines starting with `pics:`, `web gallery`, `embed`/`imbed`, `embargo` and bare URLs are treated as editorial instructions and skipped.

## Files

- `index.html` — standalone web version (also the single source of truth for the conversion engine)
- `word-digital-plugin.js` — **generated** Content Station SDK plug-in for WoodWing Studio; adds the converter to Studio's Apps menu
- `plugin-shell.js` — plug-in UI and SDK wiring (template for the generated file)
- `build-plugin.js` — extracts the conversion engine from `index.html` and produces `word-digital-plugin.js`

After changing `index.html`, regenerate the plug-in with:

```
node build-plugin.js
```

## Installing in WoodWing Studio (cloud, self-service)

The plug-in runs inside Studio's own page (same origin), so it needs no CORS changes and no WoodWing involvement.

1. Host this repo on GitHub Pages (or any HTTPS host).
2. In the Studio Server **Management Console** go to **Integrations → Studio → Plug-ins → Studio** and click **Add new**.
3. Enter the absolute URL of the plug-in file, e.g. `https://<pages-url>/word-digital-plugin.js`, and make sure it is enabled.
4. Refresh Studio — **Word → Digital Article** appears in the Apps menu.

Note: `{SESSION_ID}` URL-app wildcards are deprecated since Studio 10.40 / Enterprise Server 10.7 — the SDK plug-in route above is the supported way to integrate, and is what enables a future version to create the digital article directly in Studio via the workflow API (same-origin session).

## Workflow

1. Open the app from Studio's Apps menu.
2. Pick the article type, choose the .docx, and parse.
3. Check the detected metadata and entries; copy the feed headline to C_HEADLINE in Studio manually.
4. Download the .digital file and upload it into the target dossier.
