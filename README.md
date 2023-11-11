## Powerpoint Text Extractor

Parses a directory of powerpoint files and extracts the text to html divs with the [slidev](https://sli.dev/) ` v-click=number` format.
# Install
    
```bash
    # install and activate venv
    $ python -m venv .venv
    $ source .venv/bin/activate
    
    # install packages
    $ pip install -r requirements.txt
```

## Usage

Add a directory named `pptx` at the same level, next to this repo. Add powerpoint files to be parsed by the extractor in that folder. Then run the script.
You will be presented with the list of pptx files in the directory. Select the file you want to parse and the script will output the html to the console.

```bash
python extractr.py
```