import pandas as pd
from PyPDF2 import PdfReader
import re
from transformers import pipeline, T5Tokenizer, T5ForConditionalGeneration
import os

# ---- PART 1: Load Model ----

# Load the locally downloaded FLAN-T5-small model and tokenizer
def load_local_t5_model():
    model_dir_tokenizer = "/Users/andrew/PDF Extractor/flan-t5-small-tokenizer"
    model_dir_model = "/Users/andrew/PDF Extractor/flan-t5-small-model"
    try:
        print("Loading FLAN-T5-small model and tokenizer from local directories...")
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

# Extract sentences from the PDF file
def extract_sentences_from_pdf(pdf_file):
    reader = PdfReader(pdf_file)
    sentences = []
    for page in reader.pages:
        text = page.extract_text()
        sentences.extend(re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s', text))  # Split by sentences
    return [sentence.strip() for sentence in sentences if sentence.strip()]

# Filter sentences around the terminology and include complete sentences within 200 words before and after
def filter_sentences_with_context(sentences, terminology, context_range=200):
    extracted_contexts = []
    for i, sentence in enumerate(sentences):
        matched_terms = []
        for term in terminology:
            if re.search(rf"\b{re.escape(term)}\b", sentence.lower()):
                matched_terms.append(term)
        if matched_terms:
            # Count words and collect sentences within the range
            word_count = 0
            start, end = i, i

            # Extend range backwards
            while start > 0 and word_count <= context_range:
                word_count += len(sentences[start - 1].split())
                if word_count <= context_range:
                    start -= 1

            # Extend range forwards
            word_count = 0  # Reset word count for forward direction
            while end < len(sentences) - 1 and word_count <= context_range:
                word_count += len(sentences[end + 1].split())
                if word_count <= context_range:
                    end += 1

            # Join sentences within the range
            context = " ".join(sentences[start:end + 1]).strip()
            extracted_contexts.append((context, ", ".join(matched_terms)))
    return extracted_contexts

# Simplified summarization function
def summarize_paragraph(paragraph, summarizer, tokenizer):
    # Dynamically adjust max_length based on the input length
    input_length = len(tokenizer(paragraph, return_tensors="pt", truncation=True).input_ids[0])
    adjusted_max_length = min(50, max(10, input_length // 2))

    # Summarize the paragraph
    summary = summarizer(
        paragraph,
        max_length=adjusted_max_length,
        min_length=max(5, adjusted_max_length // 2),
        truncation=True
    )[0]['summary_text']

    return summary

# Main function to process the documents
def extract_and_summarize_to_excel(pdf_file, output_excel_file):
    # Load FLAN-T5-small summarizer and tokenizer
    summarizer, tokenizer = load_local_t5_model()

    # Load terminology and sentences
    terminology = load_terminology("/Users/andrew/PDF Extractor/Innovation Terminology.xlsx")
    sentences = extract_sentences_from_pdf(pdf_file)
    relevant_contexts = filter_sentences_with_context(sentences, terminology)

    # Summarize contexts
    summarized_data = []
    for context, terms in relevant_contexts:
        summary = summarize_paragraph(context, summarizer, tokenizer)
        summarized_data.append((context, summary, terms))

    # Save to Excel
    df = pd.DataFrame(summarized_data, columns=["Context", "Summary", "Matched Terminologies"])
    df.to_excel(output_excel_file, index=False)
    print(f"Processed {len(summarized_data)} contexts saved to {output_excel_file}.")

if __name__ == "__main__":
    # Step 2: Process PDF and Excel
    pdf_file_path = input("Please enter the path to your PDF file (e.g., '/Users/Admin/Reports/China.pdf'): ").strip("\"'")
    output_excel_file_path = "relevant_contexts_with_summary.xlsx"

    extract_and_summarize_to_excel(pdf_file_path, output_excel_file_path)