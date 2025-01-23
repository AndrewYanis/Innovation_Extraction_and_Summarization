# PDF Terminology Extraction and Summarization

This project provides a Python script for extracting and summarizing specific contextual information from PDF documents based on predefined terminologies. The script identifies sentences or paragraphs containing specified terms from an Excel file, extracts the content, and summarizes the extracted text using the FLAN-T5-small model.

## Features
- Extracts paragraphs containing terminologies from PDF documents.
- Ensures token limits are respected while extracting and summarizing text.
- Summarizes the extracted context using the locally stored FLAN-T5-small model.
- Saves the extracted and summarized data into an Excel file.

## Requirements

### Python Libraries
- pandas
- re
- transformers
- docling
- openpyxl

### Files Needed
#### FLAN-T5-small model and tokenizer:
- Downloaded locally and stored in the following directories:
  - `/Users/andrew/PDF Extractor/flan-t5-small-tokenizer`
  - `/Users/andrew/PDF Extractor/flan-t5-small-model`

#### Excel File:
- An Excel file containing a column named `Terminology` with the terms to search for in the PDF.
- File path: `/Users/andrew/PDF Extractor/Innovation Terminology.xlsx`

## How to Use
1. Place the PDF file to be processed in a directory accessible to the script.
2. Run the script and provide the path to the PDF file when prompted:

   ```bash
   python script_name.py
   ```

3. The extracted contexts and summaries will be saved to `relevant_paragraphs_with_summary.xlsx` in the working directory.

## File Descriptions

### Script
- **script_name.py**: Main script for extracting and summarizing data.

### Input Files
- **PDF File**: The document from which to extract contexts.
- **Innovation Terminology.xlsx**: Excel file with terminologies to search for in the PDF.

### Output File
- **relevant_paragraphs_with_summary.xlsx**:
  - **Paragraph**: Extracted text containing the terminology.
  - **Summary**: Summarized version of the paragraph.
  - **Matched Terminologies**: The terms found in the paragraph.

## Notes
- Ensure that the FLAN-T5-small model and tokenizer are downloaded and stored locally before running the script.
- The script dynamically adjusts summarization settings based on the length of the extracted text.
- Only non-empty paragraphs are processed and summarized.

## Example

### Input:
- **PDF file**: `/Users/andrew/PDF Extractor/2.pdf`
- **Excel file**: `/Users/andrew/PDF Extractor/Innovation Terminology.xlsx`

### Output:
- **File**: `relevant_paragraphs_with_summary.xlsx`
  - **Columns**:
    - **Paragraph**: Extracted text containing the terminology.
    - **Summary**: Summarized version of the paragraph.
    - **Matched Terminologies**: The terms found in the paragraph.

## Contact
For questions or support, please reach out via GitHub issues in this repository.
