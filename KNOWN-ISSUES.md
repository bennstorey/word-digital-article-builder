# Known issues

Open problems with evidence, so they can be picked up without re-deriving them.

---

## 1. Web images downloaded but not placed (layout has too few slots)

**Reported:** 2026-08-28
**Article:** https://www.topgear.com/car-news/hot-hatch/these-are-12-best-hot-hatches-all-time
**Status:** understood — behaves as designed, but the design is wrong for listicles

Images were fetched and added to the dossier correctly; they were not placed into
the article.

Measured on that URL:

| Article type chosen | contentType | Entries | Images found | Image slots | Placed |
|---------------------|-------------|---------|--------------|-------------|--------|
| Type 3 — Crosshead  | listicle    | 12      | 13           | **2**       | 2      |
| Type 1 — Countdown  | listicle    | 12      | 13           | **13**      | 13     |

Cause: the crosshead template has one `image` component plus one `header-image`.
The numbered templates have one `image` per entry (50). A listicle converted as
Type 3 therefore has nowhere to put all but two of its pictures.

Workaround: convert listicles as Type 1 or Type 2.

Proper fix (not done): when `content.contentType === 'listicle'`, select a numbered
layout automatically instead of leaving it to the dropdown — the parse already
knows it is a listicle (`debug.listicle`). Alternatively let `buildCrosshead`
append an image component per entry.

---

## 2. Follow / Newsletter furniture images do not resolve

**Reported:** 2026-08-28
**Status:** ROOT CAUSE CONFIRMED against the live Studio — fix not yet built

### What is actually going on

The furniture images are **per-article objects, not shared assets.** Querying
Studio for `Name contains "tg-follow"` returns **983 objects**, named
`tg-follow-newsletter-signup-{light,dark}-wide-<N>`, each one `PlacedOn` its own
dossier and created by the `WoodWing Production` user. Every article gets its own
copy.

Of the 13 image IDs baked into our three templates, `GetObjects` returns only 4:

| ID | Exists | Belongs to |
|----|--------|-----------|
| 72514, 72515 | yes | an old article (countdown newsletter pair) |
| 72517, 72520 | yes | an old article (ascending newsletter pair) |
| 72509, 72510, 72511, 72512 | **no** | crosshead's set — gone |
| 72513, 72516, 72518, 72519 | **no** | gone |
| 48816 | **no** | gone |

So the template IDs point at furniture belonging to other, older articles — some
since deleted. Nothing in the create flow makes new copies, which is exactly why
the images are "not linked or present in the dossier".

This was always going to rot: object IDs are per-article and per-environment, so
hard-coding them into a template cannot work for long.

### Fix options (needs a decision)

1. **Copy from a recent article.** Use the workflow API's copy operation to
   duplicate a known-good furniture pair into the new dossier and reference the
   new IDs. Keeps a single source of truth in Studio, but depends on a "known
   good" article existing.
2. **Ship the PNGs in the plug-in.** Embed the four images as base64 and create
   them per dossier using the image-upload path that already exists. Self
   contained, no dependency on other articles; adds roughly 200KB to the plug-in.
3. **Source them from the brand**, alongside the templates — the cleanest, and
   already on the roadmap, but the largest change.

Whichever is chosen, the `apple-news-follow` components must be rewritten with the
newly created IDs at build time rather than carrying literals.

### Superseded

An earlier reading of this issue guessed that only the countdown/ascending IDs
were wrong and that crosshead's were good, based on a known-good exported article
using `72512/72509` and `72510/72511`. Studio says those four no longer exist —
that article referenced furniture that has since been deleted. Do not "fix" this
by copying crosshead's IDs to the other templates.

The `apple-news-follow` components carry hard-coded Studio **object IDs** for their
branded images, and those IDs differ per template:

| Template  | Follow image / dark | Newsletter image / dark |
|-----------|---------------------|-------------------------|
| crosshead | **72512 / 72509**   | **72510 / 72511**       |
| countdown | 72516 / 72513       | 72514 / 72515           |
| ascending | 72519 / 72518       | 72517 / 72520           |

A known-good article exported from Studio (`Want money off a new EV.digital`, built
from the **crosshead** template) uses `72512 / 72509` and `72510 / 72511`, and its
export bundle contains exactly `48816, 72509, 72510, 72511, 72512` — nothing in the
`72513–72520` range.

So the crosshead template's IDs resolve, and the countdown/ascending ones appear
not to. That matches the report: the furniture broke on an article converted with a
numbered layout.

**Caveat:** an export bundle only contains images that article actually used, so the
absence of `72513–72520` is strong evidence but not proof they don't exist. Confirm
by searching Studio for object IDs 72516 and 72513 before changing anything.

Proposed fix: point the countdown and ascending `apple-news-follow` components at
the crosshead template's IDs (`72512 / 72509`, `72510 / 72511`). Leave crosshead
alone — it works.

Longer term these IDs are environment-specific and should not be baked into the
templates at all; sourcing templates from the brand (already on the roadmap) would
remove the problem.

---

## 3. Images appear unplaced if the article is opened immediately (FIXED)

**Reported:** 2026-08-28
**Status:** fixed in build `0dc7a5a6`

Images placed correctly, but only if the article was not opened too soon after
creation.

Cause: `CreateObjects` returns an object ID as soon as the record exists, but
Studio generates renditions asynchronously afterwards. The article was being built
and created immediately, so opening it before ingest finished showed empty
placements.

Fix: `waitForImagesReady()` polls `GetObjects` with `Rendition: 'thumb'` after
upload and waits until every new image reports a file, before the article is
built. It shows "Waiting for Studio to process images (n of m)…", times out after
60s rather than blocking forever, and a failed poll never prevents article
creation. Anything still pending is named in the completion notification.
