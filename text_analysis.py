import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
import os
import pandas as pd
import re

nltk.download('punkt')

def load_stop_words(filepaths):
    stop_words = set()
    for filepath in filepaths:
        with open(filepath, 'r', encoding='ISO-8859-1') as file:  # Changed encoding to ISO-8859-1
            words = file.read().split()
            stop_words.update(words)
    return stop_words

stop_word_files = [
    'StopWords/StopWords_Auditor.txt',
    'StopWords/StopWords_Currencies.txt',
    'StopWords/StopWords_DatesandNumbers.txt',
    'StopWords/StopWords_Generic.txt',
    'StopWords/StopWords_GenericLong.txt',
    'StopWords/StopWords_Geographic.txt',
    'StopWords/StopWords_Names.txt'
]
stop_words = load_stop_words(stop_word_files)

positive_words = set(open('MasterDictionary/positive-words.txt').read().split())
negative_words = set(open('MasterDictionary/negative-words.txt').read().split())

def clean_text(text):
 
    words = word_tokenize(text.lower())
    
    words = [word for word in words if word.isalpha() and word not in stop_words]
    return words

def calculate_scores(words):
    positive_score = sum(1 for word in words if word in positive_words)
    negative_score = (-1) * sum(-1 for word in words if word in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(words) + 0.000001)
    return positive_score, negative_score, polarity_score, subjectivity_score

def analyze_text(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()

    sentences = sent_tokenize(text)
    words = clean_text(text)
    positive_score, negative_score, polarity_score, subjectivity_score = calculate_scores(words)

    avg_sentence_length = len(words) / len(sentences)
    complex_words = [word for word in words if len([char for char in word if char in 'aeiou']) > 2]
    percentage_complex_words = len(complex_words) / len(words)
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    avg_words_per_sentence = len(words) / len(sentences)
    syllable_per_word = sum(len([char for char in word if char in 'aeiou']) for word in words) / len(words)
    personal_pronouns = len(re.findall(r'\b(I|we|my|ours|us)\b', text, re.I))
    avg_word_length = sum(len(word) for word in words) / len(words)

    return {
        'URL_ID' : url_id,
        'URL' : url,
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVITY SCORE': subjectivity_score,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
        'COMPLEX WORD COUNT': len(complex_words),
        'WORD COUNT': len(words),
        'SYLLABLE PER WORD': syllable_per_word,
        'PERSONAL PRONOUNS': personal_pronouns,
        'AVG WORD LENGTH': avg_word_length
    }

input_file = 'Input.xlsx'
output_file = 'Output Data Structure.xlsx'
df_input = pd.read_excel(input_file)
analysis_results = []

for index, row in df_input.iterrows():
    url_id = row['URL_ID']
    url=row['URL']
    file_path = f'articles/{url_id}.txt'
    analysis_result = analyze_text(file_path)
    analysis_result.update(row.to_dict())
    analysis_results.append(analysis_result)

df_output = pd.DataFrame(analysis_results)
df_output.to_excel(output_file, index=False)
