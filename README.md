# Article Rebrand Tool 

A Python tool for rebranding old Word document articles into newly branded Word documents that using structured parsing and re-writing.

---

## What It Does

This tool:
- Extracts **titles** and detects **authors**
- Preserves **headings**, **paragraph styles**, **hyperlinks**, **bold/italic/underline**, and **tables**
- Supports **unordered list detection** 
- Converts documents into a consistent pre-determined style using a **Word template**

A manual review would still be needed at this point, see Current Specifics and Limitations below

> Currently supports `.docx` input. 

I've populated the input folder with fake articles that contain common elements that need formatting (in the structure I was working with for my actual project; no content is real though). These can be used or modified for testing purposes.

---

## Project Structure

Article-Rebrand/
├── input_docs/         # Drop your original Word files here
├── output_docs/        # Reformatted documents are saved here
├── templates/          # Holds the Word template (e.g., Article-Template.docx)
├── authors.json        # List of known authors for detection
├── parser.py           # Parses original .docx into structured content
├── writer.py           # Writes structured content into new documents
├── convert.py          # Convert a single file
├── convert_batch.py    # Convert all files in input_docs/
├── requirements.txt    # Python package dependencies
├── .gitignore          # Keeps virtual env & build junk out of Git
└── venv/               # Local virtual environment (not tracked by Git)

---

## Setup

```bash
# Clone this repo and navigate to the folder
git clone https://github.com/your-username/article-rebrand.git
cd article-rebrand

# (Optional) Create virtual environment
python3 -m venv venv
source venv/bin/activate  # macOS/Linux
venv\Scripts\activate     # Windows

# Install dependencies
pip install -r requirements.txt

# Replace author names in authors.json

```

## Usage

To convert a single file:
```bash
python convert.py
```

To batch-convert all .docx files in input_docs/:
```bash
python convert_batch.py
```

## Current Specifics and Limitations

- Only unordered lists (bullets) are auto-detected. Ordered (numbered) lists must currently be fixed manually, they currently show as unordered lists.
- Tables with merged cells or complex formatting will need cleanup. Current tables are outputted with a soft grey border, visually comfortable cell spacing, and bold/italics preserved. Colours, column widths, merged cells, and varying text sizes are not supported.
- The structure of MY use case had the article titles in all the old documents inside text boxes, and a logo and byline at the top of the page and in another text box. This version as-is has this built in, but identified in parser.py
- I built this with author detection, but currently am not using it; the author was already written in the body text so I've just kept that as is. The functionality is there though, for assigning the author byline to somewhere else.

## Contributions

Feel free to fork and contribute! Submit pull requests with helpful changes or refinements.

## Credits

Created by Sarah Brown as a document reformatting and Python learning project. (www.questadon.com)
