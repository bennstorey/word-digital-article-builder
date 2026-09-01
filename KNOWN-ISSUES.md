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

The `apple-news-follow` components in our templates carry hard-coded object IDs:

| Template  | Follow image / dark | Newsletter image / dark |
|-----------|---------------------|-------------------------|
| crosshead | 72512 / 72509       | 72510 / 72511           |
| countdown | 72516 / 72513       | 72514 / 72515           |
| ascending | 72519 / 72518       | 72517 / 72520           |

Of those 13 IDs, `GetObjects` returns only **four**:

| IDs | Exists | Note |
|-----|--------|------|
| 72514, 72515, 72517, 72520 | yes | newsletter pairs belonging to old articles |
| 72509, 72510, 72511, 72512 | **no** | crosshead's set — gone |
| 72513, 72516, 72518, 72519 | **no** | gone |
| 48816 | **no** | gone |

So the template IDs point at furniture belonging to other, older articles, most of
which have since been deleted. Nothing in the create flow makes new copies, which
is exactly why the images are "not linked or present in the dossier".

This was always going to rot: object IDs are per-article and per-environment, so
hard-coding them into a template cannot work for long.

### Superseded reading

An earlier note here guessed that only the countdown/ascending IDs were wrong and
crosshead's were good, because a known-good exported article used `72512/72509`
and `72510/72511` and its bundle contained those PNGs. Studio says those four no
longer exist — that article referenced furniture deleted since. **Do not "fix"
this by copying crosshead's IDs onto the other templates.**

### Fix options (needs a decision)

1. **Copy from a recent article.** Duplicate a known-good furniture pair into the
   new dossier via the workflow API and reference the new IDs. Single source of
   truth in Studio, but depends on a "known good" article existing.
2. **Ship the PNGs in the plug-in.** Embed the images as base64 and create them
   per dossier using the image-upload path that already exists. Self-contained,
   no dependency on other articles; adds roughly 200KB to the plug-in.
3. **Source them from the brand**, alongside the templates — cleanest, already on
   the roadmap, largest change.

Whichever is chosen, the `apple-news-follow` components must be rewritten with the
newly created IDs at build time rather than carrying literals.

Open question: only `tg-follow-newsletter-signup-*` objects were found by name.
The "For more content follow this channel" image uses a different naming
convention that has not been identified yet — `Name contains "follow-channel"`
returns nothing.

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
