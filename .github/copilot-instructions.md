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
    "shape_name": "Text Box 12"
  }
]
```

- **sequence**: The order number (1, 2, 3...) for when this element appears
- **text**: The text content of the element
- **shape_name**: Original PowerPoint shape name (for reference)
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
    "static": true
  }
]
```

---

## Reconstructing Presentation Behavior

To recreate the PowerPoint experience in a web format:

1. **Initial State**: Display `static_content` elements immediately
2. **Animation**: On each click/advance, reveal the next item in `animation_sequence` by sequence number
3. **Hyperlinks**: When an element has a `hyperlink` with `type: "customshow"`:
   - Display the `linked_content.slides`
   - After viewing, return to the main slide on final click afer all conent is visible.
4. **Navigation**: After all animation_sequence items are revealed, advance to the next slide

### Example Flow (Slide 1 of "The Promises")

1. Show static: "The Promises" (title)
2. Click → Show "Genesis 12:1-3" (clickable link to custom show, displays scripture, returns to previous position at end of sub-show)
3. Click → Show "Great Nation"
4. Click → Show "Land of Canaan"
5. ... continue through sequence6. 
5. After all 35 items revealed, next click advances to slide 2