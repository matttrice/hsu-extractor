## JSON Structure Reference

The JSON output captures the presentation's animation sequence and hyperlink relationships for reconstructing the PowerPoint functionality in a web-based format.

### Top-Level Structure

```json
{
  "file_path": "/path/to/presentation.pptx",
  "file_name": "presentation.pptx",
  "total_slides": 17,
  "source_dimensions": {
    "width": 1536,
    "height": 864
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
    "layout": { "x": 60, "y": 9, "width": 259, "height": 36 },
    "font": { "font_size": 25, "bold": true, "color": "#0000CC" },
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
    "layout": { "x": 113, "y": 116, "width": 180, "height": 29 },
    "font": { "font_size": 22, "bold": true }
  },
  {
    "sequence": 3,
    "text": "Land of Canaan",
    "shape_name": "Text Box 13",
    "timing": "after",
    "delay": 500,
    "layout": { "x": 113, "y": 146, "width": 180, "height": 29 },
    "font": { "font_size": 22, "bold": true }
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
  - **curve**: Perpendicular offset from midpoint (negative = curve left/up, positive = curve right/down)
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
3. **Hyperlinks (Drills)**: When an element has a `hyperlink` with `type: "customshow"`:
   - Display the `linked_content.slides`
   - After viewing all content, return to the **origin slide** (not intermediate drills)
   - Multi-level drill chains (e.g., hebrews-3-14 → hebrews-4-1) all return directly to origin like a custom_show   
5. **Navigation**: After all animation_sequence items are revealed, advance to the next slide

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

The extractor automatically scales all coordinates from PowerPoint's slide dimensions to MBS's **960×540 canvas**. Coordinates in the JSON are already pre-calculated to final pixel values rounded to whole numbers.

**No manual scaling needed** - use JSON coordinates directly in Svelte components:

```svelte
<!-- JSON: "x": 81, "y": 12, "width": 229, "height": 56 -->
<Fragment layout={{ x: 81, y: 12, width: 229, height: 56 }}>
```

The JSON includes metadata showing the conversion applied:
- `source_dimensions`: Original PowerPoint slide size (e.g., 1536×864)
- `target_canvas`: MBS canvas size (always 960×540)
- `scale_factor`: Conversion ratio (e.g., 0.625 for 1536→960)

### Converting Arc Shapes to MBS

Freeform arc shapes (like "Arc 192") have an `arc_path` field with `from`, `to`, and `curve` values. These map directly to the MBS `Arc` component.

The `curve` value is the perpendicular offset from the midpoint of the line between `from` and `to`:
- **Negative curve** = arc curves "left" relative to the from→to direction (typically upward for left-to-right arcs)
- **Positive curve** = arc curves "right" relative to the from→to direction (typically downward for left-to-right arcs)

**JSON arc_path example:**
```json
{
  "shape_name": "Arc 192",
  "shape_type": "freeform",
  "arc_path": {
    "from": { "x": 364, "y": 382 },
    "to": { "x": 242, "y": 382 },
    "curve": -34
  },
  "line": { "width": 5, "color": "#0000FF" }
}
```

**MBS Svelte conversion:**
```svelte
<Fragment step={48} animate="draw">
  <Arc from={{ x: 364, y: 382 }} to={{ x: 242, y: 382 }} curve={-34} stroke={{ width: 5, color: '#0000FF' }} arrow />
</Fragment>
```

### Converting Lines with Arrows

When the JSON has an `arrow_ends` property with `headEnd` or `tailEnd`, use the MBS `Arrow` component instead of `Line`:

**JSON with arrow_ends:**
```json
{
  "shape_name": "Line 74",
  "layout": { "x": 192, "y": 285, "width": 1, "height": 559, "rotation": 90 },
  "line": { "width": 9 },
  "arrow_ends": { "tailEnd": "triangle" },
  "line_endpoints": { "from": { "x": 0, "y": 285 }, "to": { "x": 750, "y": 285 } }
}
```

**MBS conversion:**
```svelte
<Fragment step={6} animate="wipe">
  <Arrow from={{ x: 0, y: 285 }} to={{ x: 750, y: 285 }} stroke={{ width: 9 }} zIndex={30} />
</Fragment>
```

**Arrow component props:**
- `from`, `to`: Point-to-point mode with `{ x, y }` coordinates
- `fromBox`, `toBox`: Box-to-box mode with `{ x, y, width, height }` for curved arrows between rectangles
- `bow`: Curvature amount (0 = straight, 0.1-0.5 = curved)
- `flip`: Reverse curve direction
- `headSize`: Arrow head size multiplier (default: 3, use 0 for no head)
- `startMarker`, `endMarker`: Optional circle markers `{ radius, fill? }`

### Converting Lines (No Arrowhead)

Use `Line` for connectors without arrowheads:

**JSON line:**
```json
{
  "shape_name": "Line 72",
  "layout": { "x": 585, "y": 39, "width": 1, "height": 253, "rotation": 270 },
  "line": { "width": 4, "dash": "sysDash" },
  "line_endpoints": { "from": { "x": 459, "y": 165 }, "to": { "x": 712, "y": 165 } }
}
```

**MBS conversion:**
```svelte
<Fragment step={16.1} animate="draw">
  <Line from={{ x: 459, y: 165 }} to={{ x: 712, y: 165 }} stroke={{ width: 4, dash: '10,5' }} />
</Fragment>
```

**Line component props:**
- `from`, `to`: Start and end points with `{ x, y }` coordinates
- `stroke`: Stroke styling `{ width?, color?, dash? }`
- `startMarker`, `endMarker`: Optional circle markers `{ radius, fill? }`
- `zIndex`: Stacking order

### Understanding Line Rotation

PowerPoint stores lines as rotated rectangles. The `line_endpoints` field provides pre-calculated `from` and `to` coordinates that account for rotation:

- **rotation: 0°** → Vertical line (top to bottom)
- **rotation: 90°** → Horizontal line (left to right)  
- **rotation: 270°** → Horizontal line (right to left)
- **Other angles** → Diagonal line

**Always use `line_endpoints.from` and `line_endpoints.to`** instead of trying to interpret `layout.x/y` directly. The layout values represent the bounding box before rotation, not the actual line position.

### Dash Pattern Mapping

| JSON dash | MBS stroke.dash |
|-----------|-----------------|
| `"sysDash"` | `"10,5"` |
| `"dash"` | `"8,4"` |
| `"dot"` | `"2,2"` |
| `"lgDash"` | `"16,6"` |

### Converting Rectangles to MBS

Rectangle shapes (background columns, boxes) use the `Rect` component:

**JSON rectangle:**
```json
{
  "shape_name": "Rectangle 68",
  "layout": { "x": 75, "y": 48, "width": 275, "height": 464 },
  "fill": "#B3B3B3"
}
```

**MBS conversion:**
```svelte
<Fragment step={5} animate="wipe-down">
  <Rect x={75} y={48} width={275} height={464} fill="var(--color-level1)" zIndex={5} />
</Fragment>
```

**Rect component props:**
- `x`, `y`: Position on canvas (960×540)
- `width`, `height`: Dimensions
- `fill`: Background color
- `stroke`: Border styling `{ width?, color?, dash? }`
- `radius`: Corner radius for rounded rectangles
- `zIndex`: Stacking order

### Converting Ellipses to MBS

Ellipse shapes use the `Ellipse` component (self-positioning like Arrow/Arc/Rect):

**JSON ellipse:**
```json
{
  "shape_name": "Oval 12",
  "layout": { "x": 100, "y": 150, "width": 200, "height": 100 },
  "fill": "#FFD700"
}
```

**MBS conversion:**
```svelte
<Fragment step={3} animate="fade">
  <Ellipse cx={200} cy={200} rx={100} ry={50} fill="#FFD700" zIndex={5} />
</Fragment>
```

Note: Convert layout bounds to center coordinates:
- `cx` = x + width/2
- `cy` = y + height/2
- `rx` = width/2
- `ry` = height/2

**Ellipse component props:**
- `cx`, `cy`: Center position on canvas (960×540)
- `rx`, `ry`: Horizontal and vertical radii
- `fill`: Fill color
- `stroke`: Border styling `{ width?, color?, dash? }`
- `zIndex`: Stacking order

### Example Flow (Slide 1 of "The Promises")

1. Show static: "The Promises" (title)
2. Click → Show "Genesis 12:1-3" (clickable link to custom show, displays scripture, returns to previous position at end of sub-show)
3. Click → Show "Great Nation"
4. Click → Show "Land of Canaan"
5. ... continue through sequence
6. After all 35 items revealed, next click advances to slide 2