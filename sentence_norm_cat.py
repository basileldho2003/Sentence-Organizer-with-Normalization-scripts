from openpyxl import load_workbook, Workbook
from string_normalizer import TextProcessor
import os
import re

text_processor = TextProcessor()


def normalize_text(text):
    """Normalize text using TextProcessor."""
    return text_processor.process_text(text)


def split_sentences(text):
    """
    Splits text into sentences while preserving full stops, initials, abbreviations, and decimal numbers.
    """
    sentence_pattern = (
        r"(?<!\b[A-Z])(?<!\b[A-Z]\.)(?<!\b[A-Z]\.[A-Z])(?<!\b\d)([.?!])\s+"
    )

    sentences = re.split(sentence_pattern, text)

    # Recombine the punctuation marks that were split
    result = []
    for i in range(0, len(sentences) - 1, 2):
        result.append(sentences[i] + sentences[i + 1])

    # Append the last sentence if it wasn’t processed in the loop
    if len(sentences) % 2 == 1:
        result.append(sentences[-1])

    return [sentence.strip() for sentence in result if sentence.strip()]


# Load data from english_news_articles.xlsx
input_file = "english_news_articles.xlsx"
wb_input = load_workbook(input_file)
sheet = wb_input.active  # Read the first sheet

# Extract headers and convert them to lowercase for safe indexing
headers = [str(cell.value).strip().lower() for cell in sheet[1]]

# Ensure the required columns exist
if "domain" not in headers or "text" not in headers:
    raise ValueError("Missing 'Domain' or 'text' columns in the Excel file.")

category_idx = headers.index("domain")
text_idx = headers.index("text")

# Read and process data
data = {}
for row in sheet.iter_rows(min_row=2, values_only=True):
    if row[category_idx] is None or row[text_idx] is None:
        continue  # Skip empty rows

    category = str(row[category_idx]).strip()
    normalized_text = normalize_text(str(row[text_idx]).strip())
    sentences = split_sentences(normalized_text)

    if category not in data:
        data[category] = []
    data[category].extend(sentences)

# Load or create output Excel file
output_file = "Categorized_Sentences.xlsx"
if os.path.exists(output_file):
    wb_output = load_workbook(output_file)
else:
    wb_output = Workbook()
    wb_output.remove(wb_output.active)  # Remove default sheet if new file

# Categorize and write to appropriate sheets
for category, sentences in data.items():
    if category not in wb_output.sheetnames:
        ws = wb_output.create_sheet(title=category)
        ws.append(["Sentences"])  # Add header row
    else:
        ws = wb_output[category]

    for sentence in sentences:
        ws.append([sentence])

# Save changes
wb_output.save(output_file)
print("Sentences have been split, categorized, and saved successfully.")
