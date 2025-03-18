from openpyxl import load_workbook, Workbook
from string_normalizer import TextProcessor
import os

text_processor = TextProcessor()

def normalize_text(text):
    """Normalize text using TextProcessor."""
    return text_processor.process_text(text)

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

# Read data
data = {}
for row in sheet.iter_rows(min_row=2, values_only=True):
    if row[category_idx] is None or row[text_idx] is None:
        continue  # Skip empty rows

    category = str(row[category_idx]).strip()
    text = normalize_text(str(row[text_idx]).strip())

    if category not in data:
        data[category] = []
    data[category].append(text)

# Load or create output Excel file
output_file = "Categorized_Sentences.xlsx"
if os.path.exists(output_file):
    wb_output = load_workbook(output_file)
else:
    wb_output = Workbook()
    wb_output.remove(wb_output.active)  # Remove default sheet if new file

# Categorize and write to appropriate sheets
for category, texts in data.items():
    if category not in wb_output.sheetnames:
        ws = wb_output.create_sheet(title=category)
        ws.append([""])  # Add header row
    else:
        ws = wb_output[category]

    for text in texts:
        ws.append([text])

# Save changes
wb_output.save(output_file)
print("Sentences have been categorized and saved successfully.")
