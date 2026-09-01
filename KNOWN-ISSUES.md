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
**Status:** root cause identified, fix not applied (needs confirmation in Studio)

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
