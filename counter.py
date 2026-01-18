from docx import Document
from collections import Counter
import re

# ===============================
# USER INPUT (relative file path and top count)
# ===============================
DOCX_FILE_PATH = ".text.docx"   # <-- change this
TOP_N = 5

# ===============================
# Load and extract text
# ===============================
doc = Document(DOCX_FILE_PATH)
text = " ".join(p.text for p in doc.paragraphs)

# ===============================
# Stop words (unigrams only)
# ===============================
STOP_WORDS = {
    "the", "and", "to", "of", "a", "in", "is", "it", "that", "for",
    "on", "with", "as", "was", "were", "be", "by", "this", "are",
    "or", "at", "from", "an", "which", "but", "not", "have", "has",
    "had", "they", "their", "them", "its", "we", "our", "you",
    "your", "i", "me", "my"
}

# ===============================
# Clean and tokenize text
# ===============================
# Convert to lowercase and keep only letters and spaces
text = text.lower()
text = re.sub(r"[^a-z\s]", "", text)

words = text.split()

# ===============================
# Generate n-grams
# ===============================
unigrams = words
bigrams = zip(words, words[1:])
trigrams = zip(words, words[1:], words[2:])

# ===============================
# Count frequencies
# ===============================
# Filter stop words ONLY for unigrams
filtered_unigrams = [w for w in unigrams if w not in STOP_WORDS]

# ===============================
# Count frequencies
# ===============================
unigram_counts = Counter(filtered_unigrams)
bigram_counts = Counter([" ".join(bg) for bg in bigrams])
trigram_counts = Counter([" ".join(tg) for tg in trigrams])

# ===============================
# Display results
# ===============================

print("\nTop Unigrams:")
for word, count in unigram_counts.most_common(TOP_N):
    print(f"{word}: {count}")

print("\nTop Bigrams:")
for phrase, count in bigram_counts.most_common(TOP_N):
    print(f"{phrase}: {count}")

print("\nTop Trigrams:")
for phrase, count in trigram_counts.most_common(TOP_N):
    print(f"{phrase}: {count}")