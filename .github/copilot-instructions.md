## JSON Structure Reference

The JSON output captures the presentation's animation sequence and hyperlink relationships for reconstructing the PowerPoint functionality in a web-based format.

### Top-Level Structure

```json
{
  "file_path": "/path/to/presentation.pptx",
  "file_name": "presentation.pptx",
  "total_slides": 17,
  "custom_shows": { ... },
  "slides": [ ... ]
}
```

### Custom Shows

Custom shows are named collections of slides that can be linked from the main presentation. They function as "pop-up" content - when a user clicks a hyperlinked element, the custom show slides are displayed, then the user returns to the main slide.

```json
"custom_shows": {
  "2": {
    "name": "Gen12.1",
    "id": 2,
    "slides": [
      {
        "slide_file": "slides/slide7.xml",
        "texts": [
          "1 Now the LORD said to Abram...",
          "Genesis 12:1-3"
        ]
      }
    ]
  }
}
```

- **Key**: The custom show ID (referenced by hyperlinks)
- **name**: Display name of the custom show
- **slides**: Array of slides in the custom show, each containing:
  - **slide_file**: Original slide file reference
  - **texts**: All text content from that slide

### Slides

Each slide contains an ordered animation sequence - the exact order elements appear when presenting.

```json
"slides": [
  {
    "slide_number": 1,
    "animation_sequence": [ ... ],
    "static_content": [ ... ]
  }
]
```

### Animation Sequence

The `animation_sequence` array lists text elements in the order they appear during the presentation (via click/animation).

```json
"animation_sequence": [
  {
    "sequence": 1,
    "text": "Genesis 12:1-3",
    "shape_name": "Text Box 5",
    "timing": "click",
    "layout": { "x": 59.9, "y": 9.0, "width": 259.2, "height": 36.0 },
    "font": { "font_size": 25.2, "bold": true, "color": "#0000CC" },
    "hyperlink": {
      "type": "customshow",
      "id": 2
    },
    "linked_content": {
      "name": "Gen12.1",
      "id": 2,
      "slides": [ ... ]
    }
  },
  {
    "sequence": 2,
    "text": "Great Nation",
    "shape_name": "Text Box 12",
    "timing": "with",
    "layout": { "x": 112.8, "y": 115.5, "width": 180.0, "height": 28.9 },
    "font": { "font_size": 21.6, "bold": true }
  },
  {
    "sequence": 3,
    "text": "Land of Canaan",
    "shape_name": "Text Box 13",
    "timing": "after",
    "delay": 500,
    "layout": { "x": 112.8, "y": 145.5, "width": 180.0, "height": 28.9 },
    "font": { "font_size": 21.6, "bold": true }
  }
]
```

- **sequence**: The order number (1, 2, 3...) for when this element appears
- **text**: The text content of the element
- **shape_name**: Original PowerPoint shape name (for reference)
- **timing**: Animation timing type:
  - `"click"`: On Click - requires a new click to appear (new step)
  - `"with"`: With Previous - appears simultaneously with the previous element (same step)
  - `"after"`: After Previous - appears after the previous element's animation finishes (same step but delayed)
- **delay**: Delay in milliseconds (only present for `"after"` timing, e.g., `500` = 500ms delay)
- **layout**: Position and dimensions in pixels (960×540 canvas):
  - **x**, **y**: Top-left position
  - **width**, **height**: Dimensions
  - **rotation**: Optional rotation in degrees
- **font**: Typography settings:
  - **font_size**: Size in CSS pixels (PowerPoint points × 1.333)
  - **font_name**: Font family name
  - **bold**, **italic**: Boolean flags
  - **color**: Hex color (e.g., `"#0000CC"`)
  - **alignment**: `"left"`, `"center"`, or `"right"`
- **fill**: Background color (hex)
- **line**: Border styling: `{ color, width }`
- **arc_path** (optional): For freeform arc shapes (e.g., "Arc 192"), contains path data:
  - **from**: Start point `{ x, y }` in slide coordinates
  - **to**: End point `{ x, y }` in slide coordinates
  - **curve**: Vertical offset for arc (negative = curves up, positive = curves down)
  - **flip**: Boolean indicating horizontal flip
- **hyperlink** (optional): If the element is clickable:
  - **type**: `"customshow"` for links to custom shows
  - **id**: The custom show ID to display
- **linked_content** (optional): When `hyperlink.type` is `"customshow"`, this contains the full content of the linked custom show, including all slide texts

### Static Content

Elements that appear immediately (not animated) are listed in `static_content`:

```json
"static_content": [
  {
    "text": "The Promises",
    "shape_name": "Rectangle 63",
    "layout": { "x": 0, "y": 0, "width": 960, "height": 50 },
    "font": { "font_size": 36, "bold": true, "alignment": "center" },
    "static": true
  }
]
```

---

## Reconstructing Presentation Behavior

To recreate the PowerPoint experience in a web format:

1. **Initial State**: Display `static_content` elements immediately
2. **Animation with Timing**:
   - `timing: "click"` - Wait for user click, then show element (new step number)
   - `timing: "with"` - Show simultaneously with the previous click element (same step number)
   - `timing: "after"` - Show after `delay` milliseconds following the previous element (decimal step, e.g., step 1.1 for 500ms delay)
3. **Hyperlinks**: When an element has a `hyperlink` with `type: "customshow"`:
   - Display the `linked_content.slides`
   - After viewing, return to the main slide on final click after all content is visible.
4. **Navigation**: After all animation_sequence items are revealed, advance to the next slide

### Converting Timing to MBS Step Values

For MBS (SvelteKit presentation system), convert timing to step values:

| JSON timing | JSON delay | MBS step |
|-------------|-----------|----------|
| `"click"` | - | New integer (`step={1}`, `step={2}`, etc.) |
| `"with"` | - | Same integer as previous click (`step={2}` if previous was `step={2}`) |
| `"after"` | 500 | Previous step + 0.1 (`step={2.1}`) |
| `"after"` | 1000 | Previous step + 0.2 (`step={2.2}`) |

**Example conversion:**
```json
// JSON output:
{ "sequence": 1, "timing": "click", ... }   // → step={1}
{ "sequence": 2, "timing": "with", ... }    // → step={1}
{ "sequence": 3, "timing": "after", "delay": 500 }  // → step={1.1}
{ "sequence": 4, "timing": "click", ... }   // → step={2}
```

### Converting Arc Shapes to MBS

Freeform arc shapes (like "Arc 192") have an `arc_path` field with `from`, `to`, and `curve` values. These map directly to the MBS `Arc` component.

**Important:** The extracted coordinates are in PowerPoint's slide dimensions (e.g., 1536×864). MBS uses a 960×540 canvas. Scale the coordinates by the ratio: `960 / slide_width` (typically 0.625).

**JSON arc_path example:**
```json
{
  "shape_name": "Arc 192",
  "shape_type": "freeform",
  "arc_path": {
    "from": { "x": 582.7, "y": 611.4 },
    "to": { "x": 387.0, "y": 610.7 },
    "curve": -53.8,
    "flip": true
  },
  "line": { "width": 8.0, "color": "#0000FF" }
}
```

**MBS Svelte conversion (with scale factor 0.625):**
```svelte
<Fragment step={48} animate="draw">
  <Arc
    from={{ x: 582.7 * 0.625, y: 611.4 * 0.625 }}
    to={{ x: 387.0 * 0.625, y: 610.7 * 0.625 }}
    curve={-53.8 * 0.625}
    stroke={{ width: 8 * 0.625, color: '#0000FF' }}
    arrow
  />
</Fragment>
```

Or with pre-calculated values:
```svelte
<Fragment step={48} animate="draw">
  <Arc from={{ x: 364, y: 382 }} to={{ x: 242, y: 382 }} curve={-34} stroke={{ width: 5, color: '#0000FF' }} arrow />
</Fragment>
```

### Example Flow (Slide 1 of "The Promises")

1. Show static: "The Promises" (title)
2. Click → Show "Genesis 12:1-3" (clickable link to custom show, displays scripture, returns to previous position at end of sub-show)
3. Click → Show "Great Nation"
4. Click → Show "Land of Canaan"
5. ... continue through sequence
6. After all 35 items revealed, next click advances to slide 2