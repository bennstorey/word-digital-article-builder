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
**Status:** FIXED in build `7c8ac957`

### What was going on

The `apple-news-follow` components in our templates carried hard-coded object IDs,
a different set in each template:

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

So the template IDs pointed at objects that mostly no longer exist, and nothing in
the create flow linked any furniture into the dossier — which is exactly why the
images were "not linked or present in the dossier".

A search for `Name contains "tg-follow"` returns 983 objects
(`tg-follow-newsletter-signup-{light,dark}-wide-<N>`). Those are per-issue
variants, which initially suggested the furniture was copied per article. It is
not — see the fix below.

### Superseded reading

An earlier note here guessed that only the countdown/ascending IDs were wrong and
crosshead's were good, because a known-good exported article used `72512/72509`
and `72510/72511` and its bundle contained those PNGs. Studio says those four no
longer exist — that article referenced furniture deleted since. **Do not "fix"
this by copying crosshead's IDs onto the other templates.**

### The fix

The furniture is **not** copied per article: `GetObjects` on the shared assets shows the *same object ID* related to
both the `general-furniture` dossier (52254) and to each article's dossier. Studio
makes an asset appear in an article by adding a `Contained` relation to the
existing object, not by duplicating it.

The canonical assets, all in Brand `Top Gear` / Category `Edit`, targeted at
Apple News / "Shared Assets":

| ID | Name | Use |
|----|------|-----|
| 91357 | `tg-logo-tech-blue4x-88` | Follow, light |
| 91358 | `tg-logo-white4x-100` | Follow, dark |
| 91356 | `tg-follow-newsletter-signup-light-wide-88` | Newsletter, light |
| 91355 | `tg-follow-newsletter-signup-dark-wide-88` | Newsletter, dark |

So the fix is two parts:

1. All three templates now reference those IDs (previously each template carried
   its own stale set). The 983 `tg-follow-newsletter-signup-*-wide-<N>` objects
   are per-issue variants, not something to create per article.
   This includes the **footer block** — the `container` component with
   `styles.style = "_option14"` ("14: AN Footer"), whose image is nested inside
   `containers.main` and referenced the long-deleted `48816`. It now uses
   `91358` (`tg-logo-white4x-100`), the same asset as the Follow dark variant.
   Note `applyImageIds()` only fills top-level `image`/`header-image`
   components, so article pictures can never leak into the footer container.
2. `linkFurnitureToDossier()` adds a `Contained` relation from the new article's
   dossier to each of the four objects during create, so they are present in the
   dossier. It runs before the article is built and never fails the run.

Note the "Follow this channel" image is the Top Gear **logo** (`tg-logo-*`), which
is why searching for `follow-channel` found nothing.

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
