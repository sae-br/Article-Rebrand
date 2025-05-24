from parser import parse_doc, load_known_authors
from writer import write_new_doc

# 1. Input file
input_path = "input_docs/02 How to Make Good Decisions - Numbered List, Hyperlink.docx"

# 2. Output file
output_path = "output_docs/02 How to Make Good Decisions - Numbered List, Hyperlink - Converted.docx"

# 3. Template file
template_path = "templates/Article-Template.docx"

# 4. Load authors
authors = load_known_authors()

# 5. Parse the original file
parsed = parse_doc(input_path, known_authors=authors)

# 6. Write the new formatted document
write_new_doc(parsed, output_path, template_path=template_path)

print("✅ Conversion complete. Output saved to:", output_path)