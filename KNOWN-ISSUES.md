# Known issues

Open problems with evidence, so they can be picked up without re-deriving them.

---

## 1. Web images downloaded but not placed (layout has too few slots)

**Reported:** 2026-08-28
**Article:** https://www.topgear.com/car-news/hot-hatch/these-are-12-best-hot-hatches-all-time
**Status:** FIXED 2026-09-25 — Type 4 "snippet list" (see the note under NEXT SESSION below)

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

### NEXT SESSION — likely needs a fourth template (decided 2026-09-18, not started)

> **Done 2026-09-25: Type 4 — Snippet list.** It is the ascending template with each title's number and space removed (Benn: name alone, in the title's normal style), keeping document order.
> - **Auto-selected:** from topgear.com's `numberedList: false`, and for Word docs whose bold item names match a count in the headline.
> - **Results:** both articles above now give 12 entries with 13 frames for 13 images, where only 2 were placed before. "50 silliest American cars" gives 50 items and 51 frames.
> - **Placement:** pictures are placed by item position, or by filename number such as "7 - x.jpg".
>
> Earlier note: **2026-09-25:** article-type **auto-detect** now exists (`detectArticleType` / `detectTypeFromNumbers` in `index.html`). The snippet-list template should hook in there. Unnumbered bold-heading lists are detected as crosshead today. Word test doc: "50 of the silliest… American cars" in the TG AN+ WhatsApp export (50 bold headings, no numbers).

Reproduced again on the V12 engines article
(https://www.topgear.com/car-news/supercars/12-greatest-and-strangest-v12-engines-ever-made):
converted as Type 3, all 13 images uploaded but only 2 were placed (hero ->
`header-image`, entry 1 -> the single `image`), 11 left loose in the dossier.

These articles are **snippet-style entries**: a list of named items, each with a
picture and a paragraph, but **no numbers**. Neither existing layout fits:
Type 1/2 would add numbers that aren't in the source, and Type 3 has no image
slot per entry.

TopGear already flags the difference in `__NEXT_DATA__` —
`props.pageProps.content.numberedList`:

| Article | contentType | numberedList | items |
|---------|-------------|--------------|-------|
| 30 most tasteless cars | listicle | **true** | 30 |
| 12 best hot hatches | listicle | **false** | 12 |
| 12 greatest V12 engines | listicle | **false** | 12 |

So the plan to pick up:

1. A **fourth template, "snippet list"**: per entry an image, the item title
   (no number), then body — i.e. the numbered group without the coloured
   number op. Needs a `.digitmpl` built in Studio, or derived from the
   countdown template by stripping the number from the title component.
2. **Auto-select from the source**: listicle + `numberedList: true` -> Type 1/2
   (use `reversedList` for countdown vs ascending); listicle +
   `numberedList: false` -> snippet list; anything else -> crosshead. Keep the
   dropdown as an override, and say in the UI why it was chosen.
3. Decide whether .docx input can hit this case too (bold names, no numbers —
   today's auto-numbering fallback in `parseNumbered` would wrongly number it).

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

---

## 4. Some article images silently fail to create in Studio (FIXED)

**Reported:** 2026-09-18
**Article:** https://www.topgear.com/car-news/supercars/12-greatest-and-strangest-v12-engines-ever-made
**Status:** fixed in build `82a517d9`

Only 3 of 13 images reached the dossier (`v12`, `Untitled-1_46`,
`Untitled-2_24`). All 13 fetched fine, directly and through the proxy.

Cause: 10 of the 13 stored filenames carry a Drupal path parameter, e.g.
`_V2A0009V2.jpg;jsessionid=null_1.jpg`. `imageNameFromUrl()` only stripped the
final extension, so the object name became `_V2A0009V2.jpg;jsessionid=null_1`,
which `CreateObjects` rejects (S1026, invalid characters). Exactly the three
images without the parameter succeeded.

Fix: the image name is cut at the first `;`, giving `_V2A0009V2` — the same name
Studio already holds for earlier copies of that photo. If Studio still rejects a
name, one retry uses a strict `[A-Za-z0-9 _-]` form, falling back to
`topgear-image-<n>`. The standalone ZIP download was never affected: it names
files positionally.

## 5. Word-numbered list articles parsed to almost nothing (FIXED)

**Reported:** 2026-09-25 (found while testing auto-detect on the 13 Word docs in the TG AN+ WhatsApp export)
**Docs:** "51 Worst Cars AN+", "Ugliest F1 cars"
**Status:** fixed in build `7fd0f596`

The entries were numbered with Word's automatic numbering. mammoth drops the
numbers and emits each entry heading as its own one-item `<ol><li>` (the body
copy between them breaks the list), so the `N. Name` rule never matched.
"51 Worst Cars" parsed to 2 entries and "Ugliest F1 cars" to 0. The docx
numbering (`word/numbering.xml`) showed a single decimal list starting at 1,
so the documents read 1 → N.

Fix: `wordListEntries()` treats an `<ol>` item followed by ordinary paragraphs
as an entry heading, numbered by position. It needs three or more, so a short
numbered list inside the copy is left alone. The name is taken whole, so
"1972 Eifelland" keeps its year. Result: 51 and 15 entries. Auto-detect reports
these as "Word numbered list, 1 → N" (ascending).

## 6. Comment author inside Studio not verified (OPEN)

**Reported:** 2026-09-25
**Status:** open, low risk

Comments added by the plug-in carry `userId` from
`ContentStationSdk.getInfo().CurrentUser.UserID` (falling back to `User`,
`ShortName`, then empty). The shape of `getInfo()` inside Studio hasn't been
inspected. Comments written with `userId: "benn.storey"` render correctly
(lab object 93514). If a plug-in-created comment shows no author, log
`ContentStationSdk.getInfo()` in Studio's console and adjust `currentUserId()`
in `plugin-shell.js`.

