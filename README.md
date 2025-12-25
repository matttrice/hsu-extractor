## PowerPoint Extractor

Parses PowerPoint files and extracts complete presentation data including:

- **Animation sequences** - Ordered list of elements as they appear during presentation
- **Layout data** - Exact positions, dimensions, and styling (font, color, borders)
- **Custom shows** - Hyperlinked drill content with full slide data
- **Static content** - Non-animated elements that appear immediately

Font sizes are automatically converted from PowerPoint points to CSS pixels (× 1.333).

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