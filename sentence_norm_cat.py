from openpyxl import load_workbook, Workbook
import os
from string_normalizer import TextProcessor
from openpyxl.styles import Font


def classify_sentence_length(sentence):
    """Classifies a sentence based on the number of words."""
    word_count = len(sentence.split())

    if 1 <= word_count <= 4:
        return "Very Short (1-4 words)"
    elif 5 <= word_count <= 8:
        return "Short (5-8 words)"
    elif 9 <= word_count <= 11:
        return "Medium (9-11 words)"
    elif 12 <= word_count <= 15:
        return "Long (12-15 words)"
    elif word_count >= 16:
        return "Very Long (16+ words)"
    return "Unknown"


# Initialize text processor
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
    length_category = classify_sentence_length(text)

    if category not in data:
        data[category] = {
            "Very Short (1-4 words)": [],
            "Short (5-8 words)": [],
            "Medium (9-11 words)": [],
            "Long (12-15 words)": [],
            "Very Long (16+ words)": [],
        }
    data[category][length_category].append(text)

# Load or create output Excel file
output_file = "Categorized_Sentences.xlsx"
if os.path.exists(output_file):
    wb_output = load_workbook(output_file)
else:
    wb_output = Workbook()
    wb_output.remove(wb_output.active)  # Remove default sheet if new file

# Categorize and write to appropriate sheets
for category, length_data in data.items():
    if category not in wb_output.sheetnames:
        ws = wb_output.create_sheet(title=category)
    else:
        ws = wb_output[category]

    # Remove any existing content in the sheet
    ws.delete_rows(1, ws.max_row)

    # Write header row with bold formatting
    header_row = [
        "Very Short (1-4 words)",
        "Short (5-8 words)",
        "Medium (9-11 words)",
        "Long (12-15 words)",
        "Very Long (16+ words)",
    ]
    for col_idx, header in enumerate(header_row, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)

    max_rows = max(len(texts) for texts in length_data.values()) if length_data else 0

    for row_idx in range(max_rows):
        row_data = []
        for length_label in header_row:
            if row_idx < len(length_data.get(length_label, [])):
                row_data.append(length_data[length_label][row_idx])
            else:
                row_data.append("")  # Empty cell if no sentence available
        ws.append(row_data)

# Save changes
wb_output.save(output_file)
print("Sentences have been normalized, categorized by length, and saved successfully.")
