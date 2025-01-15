# PDF Terminology Extraction and Summarization

This project provides a Python script for extracting and summarizing specific contextual information from PDF documents based on predefined terminologies. The script identifies sentences containing specified terms from an Excel file, extracts sentences within a 200-word range before and after, and summarizes the extracted text using the FLAN-T5-small model.

## Features
- Extracts sentences containing terminologies from PDF documents.
- Includes 200 words before and after the sentences containing the terminology.
- Ensures that only complete sentences are extracted within the range.
- Summarizes the extracted context using the locally stored FLAN-T5-small model.
- Saves the extracted and summarized data into an Excel file.

## Requirements
### Python Libraries
- `pandas`
- `PyPDF2`
- `re`
- `transformers`

### Files Needed
1. **FLAN-T5-small model and tokenizer**:
   - Downloaded locally and stored in the following directories:
     - `/Users/andrew/PDF Extractor/flan-t5-small-tokenizer`
     - `/Users/andrew/PDF Extractor/flan-t5-small-model`

2. **Excel File**:
   - An Excel file containing a column named `Terminology` with the terms to search for in the PDF.
   - File path: `/Users/andrew/PDF Extractor/Innovation Terminology.xlsx`

## How to Use
1. Place the PDF file to be processed in a directory accessible to the script.
2. Run the script and provide the path to the PDF file when prompted.
   ```bash
   python script_name.py
   ```
3. The extracted contexts and summaries will be saved to `relevant_contexts_with_summary.xlsx` in the working directory.

## File Descriptions
### Script
- **`script_name.py`**: Main script for extracting and summarizing data.

### Input Files
- **PDF File**: The document from which to extract contexts.
- **`Innovation Terminology.xlsx`**: Excel file with terminologies to search for in the PDF.

### Output File
- **`relevant_contexts_with_summary.xlsx`**: Contains:
  - Extracted context.
  - Summary of the context.
  - Matched terminologies.

## Notes
- Ensure that the FLAN-T5-small model and tokenizer are downloaded and stored locally before running the script.
- The script ensures that the extracted text is within 200 words before and after the matched terminology, maintaining complete sentences.

## Example
### Input:
- PDF file: `/Users/andrew/PDF Extractor/2.pdf`
- Excel file: `/Users/andrew/PDF Extractor/Innovation Terminology.xlsx`

### Output:
- File: `relevant_contexts_with_summary.xlsx`
- Columns:
  - **Context**: Extracted sentences containing the terminology.
  - **Summary**: Summarized version of the context.
  - **Matched Terminologies**: The terms found in the context.

## Contact
For questions or support, please reach out via GitHub issues in this repository.

