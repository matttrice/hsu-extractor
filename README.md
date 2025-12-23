## Powerpoint Text Extractor

Parses a directory of PowerPoint files and extracts:
2. **JSON structure** - Complete presentation structure with animation sequences and hyperlinks (including custom shows)

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

You will be presented with the list of pptx files in the directory. Select the file you want to parse and the script will:
1. Generate a `.json` file with the same name as the PowerPoint file
2. Output the HTML to the console