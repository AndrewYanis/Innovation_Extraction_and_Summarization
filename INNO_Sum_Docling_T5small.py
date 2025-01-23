import pandas as pd
from docling.document_converter import DocumentConverter
import re
from transformers import pipeline, T5Tokenizer, T5ForConditionalGeneration

def pdf_to_markdown(source_path):
    """
    Convert a PDF document to Markdown and return the content.

    Args:
        source_path (str): The path to the source PDF file (local path).

    Returns:
        str: The Markdown content extracted from the PDF.
    """
    try:
        # Initialize the DocumentConverter
        converter = DocumentConverter()

        # Convert the PDF source to a docling document
        result = converter.convert(source_path)

        # Export the document content to Markdown format
        markdown_content = result.document.export_to_markdown()

        return markdown_content

    except Exception as e:
        print(f"An error occurred during PDF to Markdown conversion: {e}")
        raise

# ---- PART 1: Load Model ----

# Load the locally downloaded flan-t5-small model and tokenizer
def load_local_t5_model():
    model_dir_tokenizer = "/Users/andrew/PDF Extractor/flan-t5-small-tokenizer"
    model_dir_model = "/Users/andrew/PDF Extractor/flan-t5-small-model"
    try:
        print("Loading flan-t5-small model and tokenizer from local directories...")
        tokenizer = T5Tokenizer.from_pretrained(model_dir_tokenizer)
        model = T5ForConditionalGeneration.from_pretrained(model_dir_model)
        summarizer = pipeline("summarization", model=model, tokenizer=tokenizer)
        print("Model and tokenizer loaded successfully.")
        return summarizer, tokenizer
    except Exception as e:
        print("Error loading the model and tokenizer from local directories:", e)
        raise

# ---- PART 2 ----

# Load the terminology from the Excel file
def load_terminology(excel_file):
    df = pd.read_excel(excel_file)
    return set(df['Terminology'].dropna().str.lower())

# Extract text blocks from the Markdown content
def extract_text_blocks_from_markdown(markdown_content):
    blocks = re.split(r"\n\n|\n- |\n\\*", markdown_content)  # Split by blank lines or bullet points
    return [block.strip() for block in blocks if block.strip()]  # Remove empty blocks

# Filter paragraphs that contain terminology and respect token limit
def filter_paragraphs_with_terminology(blocks, terminology, tokenizer):
    extracted_contexts = []
    for block in blocks:
        matched_terms = [term for term in terminology if re.search(rf"\b{re.escape(term)}\b", block.lower())]
        if matched_terms:
            # Ensure token limit is respected
            tokenized_context = tokenizer(block, return_tensors="pt", truncation=True, max_length=512)
            truncated_context = tokenizer.decode(tokenized_context['input_ids'][0], skip_special_tokens=True)

            extracted_contexts.append((truncated_context, ", ".join(matched_terms)))
    return extracted_contexts

# Simplified summarization function
def summarize_paragraph(paragraph, summarizer, tokenizer):
    """
    Summarize a paragraph using the summarizer.

    Args:
        paragraph (str): The text to summarize.
        summarizer: The summarization pipeline.
        tokenizer: The tokenizer used with the model.

    Returns:
        str: The summary of the paragraph.
    """
    try:
        # Skip very short paragraphs
        input_length = len(tokenizer(paragraph, return_tensors="pt", truncation=True).input_ids[0])
        if input_length < 10:
            print(f"Skipping short paragraph (length: {input_length}): {paragraph}")
            return ""

        # Dynamically adjust max_length based on the input length
        adjusted_max_length = min(max(10, input_length - 1), 128)  # Ensure max_length is valid

        # Summarize the paragraph
        summary = summarizer(
            paragraph,
            max_length=adjusted_max_length,
            min_length=max(5, adjusted_max_length // 2),
            truncation=True
        )[0]['summary_text']
        return summary

    except Exception as e:
        print(f"Error during summarization: {e}")
        return "Summary not generated due to error."

# Main function to process the documents
def extract_and_summarize_to_excel(pdf_file, output_excel_file):
    # Convert PDF to Markdown
    markdown_content = pdf_to_markdown(pdf_file)

    # Load flan-t5-small summarizer and tokenizer
    summarizer, tokenizer = load_local_t5_model()

    # Load terminology and blocks
    terminology = load_terminology("/Users/andrew/PDF Extractor/Innovation Terminology.xlsx")
    blocks = extract_text_blocks_from_markdown(markdown_content)
    relevant_paragraphs = filter_paragraphs_with_terminology(blocks, terminology, tokenizer)

    # Summarize paragraphs
    summarized_data = []
    for paragraph, terms in relevant_paragraphs:
        if paragraph.strip():  # Ensure the paragraph is not empty
            summary = summarize_paragraph(paragraph, summarizer, tokenizer)
            if summary:  # Only add non-empty summaries
                summarized_data.append((paragraph, summary, terms))

    # Save to Excel
    df = pd.DataFrame(summarized_data, columns=["Paragraph", "Summary", "Matched Terminologies"])
    df.to_excel(output_excel_file, index=False)
    print(f"Processed {len(summarized_data)} paragraphs saved to {output_excel_file}.")

if __name__ == "__main__":
    # Step 2: Process PDF and Excel
    pdf_file_path = input("Please enter the path to your PDF file (e.g., '/Users/Admin/Reports/China.pdf'): ").strip("\"'")
    output_excel_file_path = "relevant_paragraphs_with_summary.xlsx"

    extract_and_summarize_to_excel(pdf_file_path, output_excel_file_path)