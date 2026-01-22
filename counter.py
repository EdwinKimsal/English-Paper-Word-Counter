from docx import Document
from collections import Counter
import re

# ===============================
# USER INPUT
# ===============================
DOCX_FILE_PATH = ".text.docx"   # <-- your path
TOP_N = 5                       # <-- length of lists

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
# Custom exclusions
# ===============================
EXCLUDED_WORDS = {"got", "very", "things", "stuff"}
EXCLUDED_PHRASES = {"this is"}

# ===============================
# Load and extract text
# ===============================
doc = Document(DOCX_FILE_PATH)
raw_text = " ".join(p.text for p in doc.paragraphs)

# ===============================
# Clean and tokenize text
# ===============================
# Replace punctuation with spaces to preserve word boundaries
text = raw_text.lower()
text = re.sub(r"[^a-z\s]", " ", text)
words = text.split()

# ===============================
# Generate n-grams
# ===============================
bigrams = [" ".join(bg) for bg in zip(words, words[1:])]
trigrams = [" ".join(tg) for tg in zip(words, words[1:], words[2:])]

# ===============================
# Filter unigrams (exact word matching only)
# ===============================
filtered_unigrams = [
    w for w in words
    if w not in STOP_WORDS and w not in EXCLUDED_WORDS
]

# ===============================
# Count frequencies
# ===============================
unigram_counts = Counter(filtered_unigrams)
bigram_counts = Counter(bg for bg in bigrams if bg not in EXCLUDED_PHRASES)
trigram_counts = Counter(trigrams)

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

# ===============================
# Display excluded-word / phrase counts (iff present)
# ===============================
print("\nExcluded Word / Phrase Counts:")

excluded_word_counts = Counter(
    w for w in words if w in EXCLUDED_WORDS
)

excluded_phrase_counts = Counter(
    bg for bg in bigrams if bg in EXCLUDED_PHRASES
)

for word in EXCLUDED_WORDS:
    if excluded_word_counts[word] > 0:
        print(f"{word}: {excluded_word_counts[word]}")

for phrase in EXCLUDED_PHRASES:
    if excluded_phrase_counts[phrase] > 0:
        print(f"{phrase}: {excluded_phrase_counts[phrase]}")
