from template_extractor import BOLTemplateExtractor
import json
import csv
from pathlib import Path

# Load config
with open("config.json", "r") as f:
    config = json.load(f)

# Ensure folders exist
new_docs_folder_path = Path(config["new_docs_folder"])
new_docs_folder_path.mkdir(parents=True, exist_ok=True)

processed_docs_folder_path = Path(config["processed_docs_folder"])
processed_docs_folder_path.mkdir(parents=True, exist_ok=True)

output_file_path = Path(config["output_file"])

# Initialize components
extractor = BOLTemplateExtractor(templates_dir="templates", confidence_threshold=0.5)
print("Configuration and components initialized.")

# Step 1: Extract templates → save to processed_docs_folder
for document in new_docs_folder_path.glob("*"):  # process all files
    result = extractor.extract_bol_text(document)

    processed_filename = processed_docs_folder_path / f"extracted_{Path(document).stem}_template.md"
    processed_filename.write_text(result, encoding="utf-8")

print("Template extraction completed.")

# Step 2: Run SmartLayer on processed docs → save outputs to CSV
write_header = not output_file_path.exists()
with output_file_path.open("a", newline="", encoding="utf-8") as csvfile:
    print("Starting SmartLayer processing...")
    writer = csv.writer(csvfile)

    if write_header:
        writer.writerow(["filename", "result"])  # adjust if multi-query with multiple columns

    print("Processing documents with SmartLayer...")
    for document in processed_docs_folder_path.glob("*.md"):
        print("Reading document:", document.name)
        document_text = document.read_text(encoding="utf-8")

        # Normalize output for CSV
        if isinstance(document_text, list):
            output_str = " | ".join(document_text)
        else:
            output_str = document_text

        # Print to terminal
        print(f"\nResults for {document.name}:\n{output_str}")

        # Save to CSV
        writer.writerow([document.name, output_str])
        csvfile.flush()  # optional safety flush per row
