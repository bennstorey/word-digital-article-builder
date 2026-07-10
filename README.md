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

## Using inside WoodWing Studio

Registered as a custom app via the Management Console (Integrations → Studio → Apps menu page):

- **Name:** Word → Digital Article
- **Application URL:** `https://<pages-url>/index.html?ticket={SESSION_ID}&server={SERVER_URL}`
- **thisTab:** false

The `ticket`/`server` wildcards are filled in by Studio Server. They are not used yet — they're in place for a future version that creates the digital article directly in Studio (requires the app origin to be added to Studio Server's CORS allowlist by WoodWing support).

## Workflow

1. Open the app from Studio's Apps menu.
2. Pick the article type, choose the .docx, and parse.
3. Check the detected metadata and entries; copy the feed headline to C_HEADLINE in Studio manually.
4. Download the .digital file and upload it into the target dossier.
