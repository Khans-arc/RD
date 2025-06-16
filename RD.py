from docx import Document 
from fuzzywuzzy import fuzz 
import string

def read_paragraphs(filepath):
    doc = Document(filepath)
    paragraphs = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            paragraphs.append(text)
    return paragraphs

def normalize_text(text):
    text = text.lower()
    text = text.translate(str.maketrans('', '', string.punctuation))
    text = ' '.join(text.split())
    return text

def find_exact_duplicates(paragraphs):
    text_to_indices = {}
    for i, text in enumerate(paragraphs):
        if text in text_to_indices:
            text_to_indices[text].append(i + 1)
        else:
            text_to_indices[text] = [i + 1]
    return text_to_indices

def find_near_duplicates(paragraphs, threshold=80):
    near_duplicates = []
    for i, text1 in enumerate(paragraphs):
        for j, text2 in enumerate(paragraphs):
            if i < j:
                normalized_text1 = normalize_text(text1)
                normalized_text2 = normalize_text(text2)
                similarity_score = fuzz.ratio(normalized_text1, normalized_text2)
                if similarity_score >= threshold:
                    near_duplicates.append((i + 1, j + 1, similarity_score, text1))
    return near_duplicates

if __name__ == "__main__":
    filepath = r"C:\Users\Muzammil\Desktop\Project\FileName.docx"
    paragraphs = read_paragraphs(filepath)

    print(f"Total paragraphs in the document: {len(paragraphs)}\n")

    duplicates = find_exact_duplicates(paragraphs)
    if duplicates:
        print("Exact Duplicates Found:\n")
        for text, indices in duplicates.items():
            if len(indices) > 1:
                print(f"Repeated at: {', '.join(map(str, indices))}")
                print(f"Text: {text}\n")
    else:
        print("No exact duplicates found.\n")

    near_duplicates = find_near_duplicates(paragraphs)
    if near_duplicates:
        print("Near Duplicates Found (Similar Paragraphs):\n")
        for index1, index2, score, text in near_duplicates:
            print(f"Paragraph {index1} and Paragraph {index2} have {score}% similarity.")
            print(f"Text: {text}\n")
    else:
        print("No near duplicates found.")
