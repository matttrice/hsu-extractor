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

Add a folder named `pptx` in the same directory as `extractr.py` and add the powerpoint files to be parsed in that folder. Then run the script.

```bash
python extractr.py
```