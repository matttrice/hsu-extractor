## JSON Structure Reference

The JSON output captures the presentation's animation sequence and hyperlink relationships for reconstructing the PowerPoint functionality in a web-based format.

### Top-Level Structure

```json
{
  "file_path": "/path/to/presentation.pptx",
  "file_name": "presentation.pptx",
  "total_slides": 17,
  "source_dimensions": {
    "width": 1536.0,
    "height": 864.0
  },
  "target_canvas": {
    "width": 960,
    "height": 540
  },
  "scale_factor": 0.625,
  "custom_shows": { ... },
  "slides": [ ... ]
}
```

**Coordinate Scaling**: All layout coordinates (x, y, width, height) are automatically scaled from the source PowerPoint dimensions to a 960×540 pixel canvas. The `scale_factor` shows the conversion ratio applied.

### Custom Shows

Custom shows are named collections of slides that can be linked from the main presentation. When a user clicks a hyperlinked element, the custom show slides are displayed, then the user returns to the main slide.

**Structure**: Custom show slides have the **same structure as regular slides**, including `animation_sequence` and `static_content` with full visual data (layout, font, fill, line, etc.).

```json
"custom_shows": {
  "2": {
    "name": "Gen12.1",
    "id": 2,
    "slides": [
      {
        "slide_file": "slides/slide7.xml",
        "animation_sequence": [
          {
            "sequence": 1,
            "text": "1 Now the LORD said to Abram...",
            "shape_name": "Text Box 3",
            "timing": "click",
            "layout": { "x": 50, "y": 80, "width": 860, "height": 400 },
            "font": { "font_size": 20, "wrap": true }
          }
        ],
        "static_content": [
          {
            "text": "Genesis 12:1-3",
            "shape_name": "Title 1",
            "static": true,
            "layout": { "x": 50, "y": 20, "width": 860, "height": 40 },
            "font": { "font_size": 28, "bold": true }
          }
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
  - **animation_sequence**: Ordered list of animated elements (same structure as regular slides)
  - **static_content**: Non-animated elements that appear immediately (same structure as regular slides)
  
**Note**: Simple scripture reference drills typically have only `static_content` with text. Complex drills with animations (like "Flesh" in Sin_Death) have full `animation_sequence` with visual data.

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
- **text**: The text content of the element (only present if shape has text content)
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
- **font**: Typography settings (only present if shape has text content):
  - **font_size**: Size in CSS pixels (PowerPoint points × 1.333). May be omitted if the text uses PowerPoint theme defaults.
  - **font_name**: Font family name
  - **bold**, **italic**: Boolean flags
  - **color**: Hex color (e.g., `"#0000CC"`)
  - **align**: `"left"`, `"center"`, or `"right"` (horizontal alignment)
  - **v_align**: `"top"`, `"middle"`, or `"bottom"` (vertical alignment, from PowerPoint's anchor property)
  - **wrap**: Boolean flag indicating if text wraps within the shape bounds (defaults to false if omitted)
- **fill**: Background color (hex)
- **line**: Border/stroke styling:
  - **width**: Stroke width in pixels
  - **color**: Hex color (e.g., `"#0000FF"`)
  - **theme_color**: Theme color reference (e.g., `"TEXT_1 (13)"`)
  - **dash**: Dash pattern type (`"dash"`, `"sysDash"`, `"dot"`, `"lgDash"`, etc.)
- **arrow_ends** (optional): For lines with arrowheads:
  - **headEnd**: Arrow type at start (`"triangle"`, `"stealth"`, `"diamond"`, `"oval"`, `"arrow"`)
  - **tailEnd**: Arrow type at end (same values)
- **line_endpoints** (optional): For line/connector shapes, pre-calculated endpoints accounting for rotation:
  - **from**: Start point `{ x, y }` in slide coordinates  
  - **to**: End point `{ x, y }` in slide coordinates
  - Use these values directly instead of calculating from layout + rotation
- **arc_path** (optional): For freeform arc shapes (e.g., "Arc 192"), contains path data:
  - **from**: Start point `{ x, y }` in slide coordinates
  - **to**: End point `{ x, y }` in slide coordinates
  - **curve**: Vertical offset for arc (negative = curves up, positive = curves down)
  - **flip**: Boolean indicating horizontal flip
- **hyperlink** (optional): If the element is clickable:
  - **type**: `"customshow"` for links to custom shows
  - **id**: The custom show ID to display
- **linked_content** (optional): When `hyperlink.type` is `"customshow"`, this contains the full content of the linked custom show, including all slide texts which will translate to a drillTo in MBS svelte.

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

### Coordinate Scaling (PowerPoint → MBS)

The extractor automatically scales all coordinates from PowerPoint's slide dimensions to MBS's **960×540 canvas**. Coordinates in the JSON are already pre-calculated to final pixel values rounded to 1 decimal place.

**No manual scaling needed** - use JSON coordinates directly in Svelte components:

```svelte
<!-- JSON: "x": 81.1, "y": 12.0, "width": 229.1, "height": 56.4 -->
<Fragment layout={{ x: 81.1, y: 12, width: 229.1, height: 56.4 }}>
```

The JSON includes metadata showing the conversion applied:
- `source_dimensions`: Original PowerPoint slide size (e.g., 1536×864)
- `target_canvas`: MBS canvas size (always 960×540)
- `scale_factor`: Conversion ratio (e.g., 0.625 for 1536→960)

### Converting Arc Shapes to MBS

Freeform arc shapes (like "Arc 192") have an `arc_path` field with `from`, `to`, and `curve` values. These map directly to the MBS `Arc` component.

**JSON arc_path example (coordinates already scaled to 960×540):**
```json
{
  "shape_name": "Arc 192",
  "shape_type": "freeform",
  "arc_path": {
    "from": { "x": 364.2, "y": 382.1 },
    "to": { "x": 241.9, "y": 381.7 },
    "curve": -33.6,
    "flip": true
  },
  "line": { "width": 5.0, "color": "#0000FF" }
}
```

**MBS Svelte conversion (use coordinates directly):**
```svelte
<Fragment step={48} animate="draw">
  <Arc from={{ x: 364.2, y: 382.1 }} to={{ x: 241.9, y: 381.7 }} curve={-33.6} stroke={{ width: 5, color: '#0000FF' }} arrow />
</Fragment>
```

### Converting Lines with Arrows and Dashes

When the JSON has an `arrow_ends` property with `headEnd` or `tailEnd`, use the MBS `Arrow` component instead of `Line`:

**JSON with arrow_ends (coordinates already scaled to 960×540):**
```json
{
  "shape_name": "Line 74",
  "layout": { "x": 191.5, "y": 285.0, "width": 0.6, "height": 558.6, "rotation": 90.0 },
  "line": { "width": 9.4, "theme_color": "TEXT_1 (13)" },
  "arrow_ends": { "tailEnd": "triangle" },
  "line_endpoints": { "from": { "x": 0, "y": 285.3 }, "to": { "x": 750.3, "y": 285.3 } }
}
```

**MBS conversion (use coordinates directly):**
```svelte
<Fragment step={6} animate="wipe">
  <Arrow from={{ x: 0, y: 285.3 }} to={{ x: 750.3, y: 285.3 }} stroke={{ width: 9.4, color: '#000000' }} zIndex={30} />
</Fragment>
```

### Understanding Line Rotation

PowerPoint stores lines as rotated rectangles. The `line_endpoints` field provides pre-calculated `from` and `to` coordinates that account for rotation:

- **rotation: 0°** → Vertical line (top to bottom)
- **rotation: 90°** → Horizontal line (left to right)  
- **rotation: 270°** → Horizontal line (right to left)
- **Other angles** → Diagonal line

**Always use `line_endpoints.from` and `line_endpoints.to`** instead of trying to interpret `layout.x/y` directly. The layout values represent the bounding box before rotation, not the actual line position.

**JSON with rotation (coordinates already scaled to 960×540):**
```json
{
  "shape_name": "Line 72",
  "layout": { "x": 585.1, "y": 38.9, "width": 1.3, "height": 252.7, "rotation": 270.0 },
  "line": { "width": 3.8, "dash": "sysDash" },
  "line_endpoints": { "from": { "x": 459.4, "y": 165.3 }, "to": { "x": 712.1, "y": 165.3 } }
}
```

**MBS conversion (use coordinates directly):**
```svelte
<Fragment step={16.1} animate="draw">
  <Line from={{ x: 459.4, y: 165.3 }} to={{ x: 712.1, y: 165.3 }} stroke={{ width: 3.8, dash: '10,5' }} />
</Fragment>
```

For dashed lines, add the `dash` property to the stroke:

**JSON with dash:**
```json
{
  "shape_name": "Line 126",
  "line": { "width": 7.7, "color": "#0000FF", "dash": "sysDash" }
}
```

**MBS conversion:**
```svelte
<Fragment>
  <Line from={{ x: 0, y: 260 }} to={{ x: 960, y: 260 }} stroke={{ width: 4.8, color: '#0000FF', dash: '10,5' }} zIndex={3} />
</Fragment>
```

**Dash pattern mapping:**
| JSON dash | MBS stroke.dash |
|-----------|-----------------|
| `"sysDash"` | `"10,5"` |
| `"dash"` | `"8,4"` |
| `"dot"` | `"2,2"` |
| `"lgDash"` | `"16,6"` |

### Example Flow (Slide 1 of "The Promises")

1. Show static: "The Promises" (title)
2. Click → Show "Genesis 12:1-3" (clickable link to custom show, displays scripture, returns to previous position at end of sub-show)
3. Click → Show "Great Nation"
4. Click → Show "Land of Canaan"
5. ... continue through sequence
6. After all 35 items revealed, next click advances to slide 2