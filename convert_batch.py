import os
from parser import parse_doc, load_known_authors
from writer import write_new_doc

INPUT_FOLDER = "input_docs"
OUTPUT_FOLDER = "output_docs"
TEMPLATE_PATH = "templates/Article-Template.docx"

authors = load_known_authors()

for filename in os.listdir(INPUT_FOLDER):
    if filename.lower().endswith(".docx") and not filename.startswith("~$"):
        input_path = os.path.join(INPUT_FOLDER, filename)
        output_path = os.path.join(OUTPUT_FOLDER, filename + " - Converted.docx")

        print(f"🔄 Converting: {filename}")
        try:
            parsed = parse_doc(input_path, known_authors=authors)
            write_new_doc(parsed, output_path, template_path=TEMPLATE_PATH)
            print(f"✅ Saved to {output_path}")
        except Exception as e:
            print(f"❌ Failed to convert {filename}: {e}")