from openpyxl import load_workbook

# File paths
english_news_path = "english_news_articles.xlsx"
categorized_sentences_path = "Categorized_Sentences.xlsx"

# Load the English news articles workbook
wb_articles = load_workbook(english_news_path)
ws_articles = wb_articles[wb_articles.sheetnames[0]]  # First sheet

# Extract headers
headers = [cell.value for cell in ws_articles[1]]

# Find column indices for 'text' and 'Domain'
text_col_idx = headers.index("text") + 1  # Convert to 1-based index
domain_col_idx = headers.index("Domain") + 1

# Dictionary to store categorized sentences
categorized_sentences = {}

# Read data row by row (starting from the second row, skipping headers)
for row in ws_articles.iter_rows(min_row=2, values_only=True):
    text = row[text_col_idx - 1]  # Convert index back to 0-based
    category = row[domain_col_idx - 1]

    # Normalize text (remove commas and hyphens)
    if isinstance(text, str):
        normalized_text = text.replace(",", "").replace("-", " ")
    else:
        continue

    # Add to the appropriate category
    if category not in categorized_sentences:
        categorized_sentences[category] = []
    categorized_sentences[category].append([normalized_text])  # Wrap in list for row format

# Load the Categorized Sentences workbook
wb_categorized = load_workbook(categorized_sentences_path)

# Ensure existing sheets are correctly populated
for sheet_name in wb_categorized.sheetnames:
    ws = wb_categorized[sheet_name]

    # Check if sheet is empty (only header exists)
    if ws.max_row == 1:
        ws.append(["Sentence"])  # Ensure the header exists

    # Clear existing data but keep the header intact
    ws.delete_rows(2, ws.max_row)  

    # Append normalized sentences if available for the category
    if sheet_name in categorized_sentences:
        for sentence in categorized_sentences[sheet_name]:
            ws.append(sentence)

# Save the workbook
wb_categorized.save(categorized_sentences_path)

print("Categorization and normalization completed successfully.")
