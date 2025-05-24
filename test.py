from parser import parse_doc, load_known_authors

authors = load_known_authors()
result = parse_doc("input_docs/BA-CL A Key to Board Member Excellence - 2004.docx", authors)

print("Title:", result["title"])
print("Author:", result.get("author")) # Added ".get" so if there's no author, no errors are thrown
print("Body Preview:", result["structured_body"][:10])  # Show first 10 blocks