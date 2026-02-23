## PowerPoint Extractor for Master Bible Study

Utility tool that parses PowerPoint files and extracts presentation data for [Master Bible Study (MBS)](https://github.com/matttrice/mbs).

Extracts:

- **Animation sequences** - Ordered list of elements as they appear during presentation
- **Layout data** - Exact positions, dimensions, and styling (font, color, borders)
- **Custom shows** - Hyperlinked drill content with full slide data
- **Static content** - Non-animated elements that appear immediately
- **Images** - Automatically extracted to a folder with the same name as the JSON

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
│   ├── 08-The_Ark.json
│   ├── 08-The_Ark/      ← Images extracted from the PPTX
│   │   ├── image1.jpg
│   │   ├── image2.png
│   │   └── ...
│   └── ...
```

## Prompt for Reproduction to MBS
Steps to reproduce Power Point to Svelte [MBS](https://github.com/matttrice/mbs). 

The prompt will reference copilot-instructions from both the mbs repo and hsu-extractor.

1. Pre-scale pptx to 16:9 WideScreen:
    -  PPTX -> Design -> Slide Size -> Widescreen, scale up/down for uniformity if prompted.
2. Before extraction, mark all **non-top-level** drill slides as **Hidden** in PowerPoint.
	- This applies to both decks using `custom_shows` and decks using pure hyperlink chains.
	- Extractor classification is deterministic: Hidden → `linked_slides`, Non-hidden → `slides[]`.
3. Run PowerPoint -> Export -> PNG Slides to [mbs/static](../mbs/static/export) for ReferenceOverlay(s). 
4. Run `extractor.py` for pptx to json, move images to static/export folder (if they exist).
	- move extracted images to static/export/
5. Add json to context.
	- For decks with `custom_shows`, treat each `custom_shows[id].slide_numbers[]` value as a reference to a `linked_slides` `slide_number`.
6. Prompt: [/create-presentation](../mbs/.github/prompts/create-lesson.prompt.md) route-name.

## License

This project is licensed under the Creative Commons Attribution 4.0 International License (CC-BY 4.0).

You are free to use, modify, and distribute this work with proper attribution to the original source.

**How to attribute:** Include a link to this repository and reference the CC-BY 4.0 license. This allows anyone to compare your version with the original if modifications are made.

See the [LICENSE](LICENSE) file for complete details.
