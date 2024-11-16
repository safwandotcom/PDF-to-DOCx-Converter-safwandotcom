# -----------------------------------------------------------------------------
# Code created by: Mohammed Safwanul Islam @safwandotcom®
# Project: Data Science File Conversion 
# Date created: 15th November 2024
# Organization: N/A
# -----------------------------------------------------------------------------
# Description:
#   This code offers a user-friendly solution for converting PDF content into a DOCX format, 
#   preserving textual information while structuring it into paragraphs for better readability within a Word document.
# License:
# This code belongs to @safwandotcom®.
# Code can be freely used for any purpose with proper attribution.
# -----------------------------------------------------------------------------

import PyPDF2  
from docx import Document                                     
import os  
import re  # Library for working with regular expressions

# Function to extract text from a PDF file
def pdf_to_text(pdf_file):
    # Open the PDF file in binary mode ('rb' - read binary)
    with open(pdf_file, 'rb') as file:
        # Initialize the PdfReader to read the PDF content
        reader = PyPDF2.PdfReader(file)
        text = ""  
        
        # Iterate through each page in the PDF
        for page_num in range(len(reader.pages)):
            # Extract text from the current page
            page = reader.pages[page_num]
            page_text = page.extract_text()
            
            # For debugging, print the text extracted from each page
            print(f"Page {page_num + 1} text:\n", page_text)
            
            # Append the text of the current page to the complete text
            text += page_text

    
    return text

# Function to convert the extracted PDF text into a DOCX file
def pdf_to_docx(pdf_file, output_file):
    # Extract the text from the PDF using the pdf_to_text() function
    text = pdf_to_text(pdf_file)        
    
    # Create a new Document object (represents a .docx file)
    doc = Document()

    # Split the extracted text into paragraphs based on common punctuation (e.g., ., !, ?)
    # This helps format the content into more natural paragraphs when adding to the DOCX
    paragraphs = re.split(r'(?<=[.!?]) +', text)  # Regular expression to split text after punctuation

    # Loop through each paragraph in the split text
    for paragraph in paragraphs:
        # Check if the paragraph contains non-whitespace characters
        if paragraph.strip():
            # Add the cleaned-up paragraph to the DOCX document
            # The strip() method removes leading/trailing whitespace
            doc.add_paragraph(paragraph.strip())
    
    # Save the DOCX file with the specified output filename
    doc.save(output_file)
    print(f"PDF content has been saved as {output_file}")

# Main function to handle user input and control the program flow
def main():
    # Ask the user for the path to the PDF file they want to convert
    pdf_file = input("Enter the path of the PDF file: ").strip('"')
    
    # Ask the user for the name of the output DOCX file (e.g., "output.docx")
    output_file = input("Enter the name for the output DOCX file (e.g., output.docx): ")
    
    # Call the pdf_to_docx function to convert the PDF into a DOCX
    pdf_to_docx(pdf_file, output_file)

# Ensure the main function is executed when the script is run directly
if __name__ == "__main__":
    main()
