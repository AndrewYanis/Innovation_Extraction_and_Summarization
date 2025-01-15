from transformers import T5Tokenizer, T5ForConditionalGeneration

# Specify the correct model name
model_name = "google/flan-t5-small"

# Download and save the model and tokenizer
tokenizer = T5Tokenizer.from_pretrained(model_name)
model = T5ForConditionalGeneration.from_pretrained(model_name)

# Save locally
tokenizer.save_pretrained("./flan-t5-small-tokenizer")
model.save_pretrained("./flan-t5-small-model")