import fitz  # PyMuPDF
import sys
import json
import os

# Get the PDF path, JSON file path, and completion file path from arguments
pdf_path = sys.argv[1]
json_file_path = sys.argv[2]
html_json_file_path = sys.argv[3]
completion_file_path = sys.argv[4]

# Ensure that the completion file path expands the '~'
# completion_file_path = os.path.expanduser(completion_file_path)

# Read the JSON file
with open(json_file_path, 'r', encoding='utf-8') as json_file:
    data = json.load(json_file)

# Read the HTML JSON file
with open(html_json_file_path, 'r', encoding='utf-8') as html_json_file:
    html_data = html_json_file.read()  # Read the HTML JSON file as a string

# Open the PDF
pdf_document = fitz.open(pdf_path)

# Iterate over the pages and form fields
for page_num in range(len(pdf_document)):
    page = pdf_document.load_page(page_num)
    form_fields = page.widgets()

    for widget in form_fields:
        field_name = widget.field_name
        if field_name in data:
            widget.reset()  # Reset the field before setting the value
            widget.field_value = data[field_name]  # Set the form field value
            widget.update()  # Update the widget to save the changes
            widget.field_value = data[field_name] + ' '
            widget.update()

# Create a new hidden text field named "store" on the first page
first_page = pdf_document.load_page(0)

# Define a rectangle for the hidden field (can be very small or off the visible area)
# Create a 1x1 rectangle at the top-left corner
store_field_rect = fitz.Rect(0, 0, 2, 2)

# Create a new widget
store_widget = fitz.Widget()
store_widget.rect = store_field_rect
store_widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
store_widget.field_name = "store"

store_widget.field_value = html_data

# Set flags to hide the field
store_widget.field_flags = fitz.PDF_ANNOT_IS_HIDDEN
store_widget.field_flags = fitz.PDF_TX_FIELD_IS_COMB

# Add the widget to the page
first_page.add_widget(store_widget)

store_widget.text_fontsize = 0  # Make the text size 0 to ensure it's invisible
# Set the text color to white (invisible on a white page)
store_widget.text_color = (1, 1, 1)
store_widget.update()  # Update the widget to save the changes

# Set the PageMode to display the thumbnails (pages panel)
pdf_document.set_pagemode("UseThumbs")

# This seems to be needed to update the appearances of the fields... (in preview)
pdf_document.need_appearances(True)

# Save the updated PDF
output_pdf = pdf_path.replace(".pdf", "_filled.pdf")
pdf_document.save(output_pdf)


# Close the PDF
pdf_document.close()

# Create a completion file to signal that the script is finished
with open(completion_file_path, 'w', encoding='utf-8') as completion_file:
    completion_file.write("Python script completed successfully!")

print(f"PDF saved to {output_pdf}")
