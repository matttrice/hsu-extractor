## PowerPoint Extractor

Parses PowerPoint files and extracts complete presentation data including:

- **Animation sequences** - Ordered list of elements as they appear during presentation
- **Layout data** - Exact positions, dimensions, and styling (font, color, borders)
- **Custom shows** - Hyperlinked drill content with full slide data
- **Static content** - Non-animated elements that appear immediately

**Automatic Conversions:**
- **Coordinates**: All layout coordinates (x, y, width, height) are automatically scaled from the source PowerPoint dimensions to a 960×540 pixel canvas (16:9 aspect ratio). No manual scaling needed.
- **Font sizes**: Automatically converted from PowerPoint points to CSS pixels (× 1.333).
- **Metadata**: The JSON includes `source_dimensions`, `target_canvas`, and `scale_factor` for reference.

## Install
    
```bash
# install and activate venv
$ python -m venv .venv
$ source .venv/bin/activate

# install packages
$ pip install -r requirements.txt
```

## Usage

Add a directory named `pptx` (or `hsu-pptx`) at the same level, next to this repo. Add PowerPoint files to be parsed by the extractor in that folder. Then run the script.

```bash
python extractor.py
```

You will be presented with the list of pptx files in the directory. Select the file you want to parse and the script will generate a `.json` file in the `extracted/` folder.

```
hsu-extractor/
├── extractor.py
├── extracted/           ← JSON output files go here
│   ├── 09-The_Promises.json
│   └── ...
```

## Prompt for Reproduction to MBS
Steps to reproduce Power Point to Svelte [MBS](https://github.com/matttrice/mbs). When running the prompt you should see it reference copilot-instructions from both the mbs repo and hsu-extractor.

1. Export PPTX PNG Slides to [mbs/static](../mbs/static/export) for ReferenceOverlay(s). Add to context.
2. Run `extractor.py` for pptx to json, add to context. 
3. Prompt: 
    - Use the new exported pptx "<json-file>" to create a new presentation route "<url-name>". The ReferenceOverlays are exported to "<location>". Add a new link to the main navigation, create all slides, animmations and drill refererences and adhere to current standards.   