# Translation — Svelte Route Map

## Main Presentation

| Route | File | Source | Description |
|-------|------|--------|-------------|
| `/translation` | `+page.svelte` | — | PresentationProvider (4 slides) |
| | `slides/Slide1.svelte` | Slide 1 | Translation diagram: Incarnation/Excarnation columns, E=MC², Enter/Exit arrows |
| | `slides/Slide2.svelte` | Slide 2 | Genesis 6:1-8 (NIV) scripture with bold highlight overlay |
| | `slides/Slide3.svelte` | Slide 3 | Nephilim analysis: Sons of God = Angels = Fallen Ones, FLOOD, Ark, post-flood questions |
| | `slides/Slide4.svelte` | Slide 4 | End of lesson (background image) |

## Standalone Scripture Drill Routes

| Route | Files | Source | Custom Show |
|-------|-------|--------|-------------|
| `/translation/genesis-5-21-24` | Content.svelte, +page.svelte | Linked slide 5 | CS 12 "genesis.5.21-24" |
| `/translation/2-kings-2-11-12` | Content.svelte, +page.svelte | Linked slide 6 | CS 13 "2kings.2.11-12" |
| `/translation/matthew-17-1-3` | Content.svelte, +page.svelte | Linked slide 7 | CS 14 "matthew.17.1-3" |
| `/translation/luke-24-50-52` | Content.svelte, +page.svelte | Linked slide 8 | CS 15 "luke.24.50-52" |
| `/translation/genesis-3-1-5` | Content.svelte, +page.svelte | Linked slide 9 | CS 5 "genesis.3.1-5" |
| `/translation/exodus-33-21-23` | Content.svelte, +page.svelte | Linked slide 10 | CS 7 "exodus.33.21-23" |
| `/translation/acts-10-30-31` | Content.svelte, +page.svelte | Linked slide 11 | CS 10 "acts.10.30-31" |
| `/translation/genesis-18-1-3` | Content.svelte, +page.svelte | Linked slide 14 | CS 6 "genesis.18.1-3" |
| `/translation/numbers-22-29-33` | Content.svelte, +page.svelte | Linked slide 15 | CS 8 "numbers.22.29-33" |
| `/translation/acts-22-6-9` | Content.svelte, +page.svelte | Linked slide 16 | CS 11 "acts.22.6-8" |

## Standalone Non-Scripture Drill Routes

| Route | Files | Source | Custom Show |
|-------|-------|--------|-------------|
| `/translation/angels-sin` | +page.svelte | Linked slide 20 | CS 0 "AngelsSin" |
| `/translation/genesis-6-1` | +page.svelte | Slide 2 (main) | CS 1 "Genesis6.1" |

## Multi-Slide CustomShowProvider Routes

| Route | Files | Source Slides | Custom Show |
|-------|-------|---------------|-------------|
| `/translation/ezekiel-1-daniel-10` | Content1.svelte, Content2.svelte, +page.svelte | Linked 12, 13 | CS 9 "ezekiel.1-daniel.10" |
| `/translation/post-flood` | Content1.svelte, Content2.svelte, +page.svelte | Linked 21, 22 | CS 2 "post_flood" |
| `/translation/gates` | Content1.svelte, Content2.svelte, +page.svelte | Linked 23, 24 | CS 3 "John10.1" |
| `/translation/angels` | Content1.svelte, Content2.svelte, Content3.svelte, +page.svelte | Linked 18, 19, 17 | CS 4 "Angels" |

## Aggregator CustomShowProvider Routes

| Route | Aggregated Content From | Custom Show |
|-------|------------------------|-------------|
| `/translation/incarnation` | genesis-3-1-5, genesis-18-1-3, exodus-33-21-23, numbers-22-29-33, ezekiel-1-daniel-10 (Content1 + Content2), acts-10-30-31, acts-22-6-9 | CS 16 "Incarnation" |
| `/translation/excarnation` | genesis-5-21-24, 2-kings-2-11-12, matthew-17-1-3, luke-24-50-52 | CS 17 "Excarnation" |

## Drill Relationships (Slide 1 → Drill Routes)

| Step | Fragment Text | DrillTo Route |
|------|--------------|---------------|
| 3 | Incarnation | `/translation/incarnation` |
| 4 | Genesis 3:1-5, 8 | `/translation/genesis-3-1-5` |
| 4 | Genesis 18:1-2  3 | `/translation/genesis-18-1-3` |
| 4 | Exodus 33:21-23 | `/translation/exodus-33-21-23` |
| 4 | Numbers 22:29-33 | `/translation/numbers-22-29-33` |
| 4 | Ezekiel 1:25;8:1 Daniel 10:4 | `/translation/ezekiel-1-daniel-10` |
| 4 | Acts 10:30-31 | `/translation/acts-10-30-31` |
| 4 | Acts 22:6 | `/translation/acts-22-6-9` |
| 5 | Excarnation | `/translation/excarnation` |
| 5 | Luke 24:50-52 | `/translation/luke-24-50-52` |
| 5 | Matthew 17:1-3 | `/translation/matthew-17-1-3` |
| 5 | 2 Kings 2:11-12 | `/translation/2-kings-2-11-12` |
| 5 | Genesis 5:21-24 | `/translation/genesis-5-21-24` |

## Drill Relationships (Slide 3 → Drill Routes)

| Step | Fragment Text | DrillTo Route |
|------|--------------|---------------|
| static | Genesis 6:1-8 (title) | `/translation/genesis-6-1` |
| 1 | Used 5 times in OT | `/translation/angels` |
| 9 | Sins of Angels: Jude 6, 2 Peter 2:4 | `/translation/angels-sin` |
| 14 | Numbers 13:32-33, Deut 1:27-28... | `/translation/post-flood` |
| 15 | Gates: John 10:1-7 | `/translation/gates` |

## Notes

- **CS 1 "Genesis6.1"** references main slide 2 (not a linked slide). Created as standalone `genesis-6-1` drill route with the same scripture content.
- **Slide 1 seq 9** has hyperlink to slide 4 (End of lesson) — ignored as non-functional.
- **Slide 1 seq 15** "Moses protected from glory of God" has hyperlink to linked slide 11 (Acts 10:30-31) — rendered as plain text without drillTo (likely a PowerPoint authoring error; Exodus 33 is the contextually correct reference).
- **Slide 3 seq 3** (Line 10) has degenerate line_endpoints (same from/to point). Rendered using layout-derived coordinates as a horizontal separator line.
- **Slide 2 seq 1** `is_scripture` flag with `E=MC²` formula (seq 2 on Slide 1) was a false positive from the extractor scripture detector — rendered as regular text.
