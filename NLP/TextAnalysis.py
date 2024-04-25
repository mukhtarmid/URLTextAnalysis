import pandas as pd
import requests
from bs4 import BeautifulSoup # type: ignore
import string
import nltk # type: ignore
import os
nltk.download('punkt')
import re
import openpyxl # type: ignore
from statistics import mean
from nltk.tokenize import word_tokenize, sent_tokenize # type: ignore
from nltk.corpus import stopwords # type: ignore
nltk.download('stopwords')

# Step 1: Read the input.xlsx &  file to extract the URLs
df_input = pd.read_excel(r'D:/USER/Desktop/Projects/NLP/Input.xlsx') # Change the directory
df_output = pd.read_excel(r'D:/USER/Desktop/Projects/NLP/Output Data Structure.xlsx') # Change the directory
urls = df_input['URL'].tolist()
url_ids = df_input['URL_ID'].tolist()

article_dict = {}  # Initialize an empty dictionary to store URL and article text

# Step 2: Extract the article title and text from each URL and save them in the dictionary
for url, url_id in zip(urls, url_ids):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Extract the article title
    title = soup.find('h1')
    if title is None:
        title = ''
    else:
        title = title.text.strip()

    # Extract the article text
    article_text = ' '.join([p.text for p in soup.find_all('p')])

    # Combine title and text
    content = f'{title}\n\n{article_text}'

    # Save the extracted article title, text, and url_id in the dictionary with the URL as the key
    article_dict[url] = {'url_id': url_id, 'content': content }

# Create a directory named TextFolder if it doesn't exist
if not os.path.exists('TextFolder'):
    os.makedirs('TextFolder')

# Save each URL and its content in separate files within the TextFolder directory
for url, data in article_dict.items():
    url_id = data['url_id']
    content = data['content']
    file_path = os.path.join('TextFolder', f'{url_id}.txt')
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)

# Function to read stop words from a file and tokenize them
def read_words(file_path, encoding='iso-8859-1'):
    with open(file_path, 'r', encoding=encoding) as f:
        stop_words = f.read()
    stop_words = nltk.word_tokenize(stop_words)
    stop_words = [x.lower() for x in stop_words]
    return stop_words

stpwrdsgenericlong = read_words(r'D:/USER/Desktop/Projects/NLP/StopWords/StopWords_GenericLong.txt', encoding='iso-8859-1') # Change the directory
stpwrdsauditor = read_words(r'D:/USER/Desktop/Projects/NLP/StopWords/StopWords_Auditor.txt', encoding='iso-8859-1') # Change the directory
stpwrdscurrencies = read_words(r'D:/USER/Desktop/Projects/NLP/StopWords/StopWords_Currencies.txt', encoding='iso-8859-1') # Change the directory
stpwrdsdatesandnumbers = read_words(r'D:/USER/Desktop/Projects/NLP/StopWords/StopWords_DatesandNumbers.txt', encoding='iso-8859-1') # Change the directory
stpwrdsgeographic = read_words(r'D:/USER/Desktop/Projects/NLP/StopWords/StopWords_Geographic.txt', encoding='iso-8859-1') # Change the directory
stpwrdsnames = read_words(r'D:/USER/Desktop/Projects/NLP/StopWords/StopWords_Names.txt', encoding='iso-8859-1') # Change the directory
stpwrdsgeneric = read_words(r'D:/USER/Desktop/Projects/NLP/StopWords/StopWords_Generic.txt', encoding='iso-8859-1') # Change the directory
positivewrds = read_words(r'D:/USER/Desktop/Projects/NLP/MasterDictionary/positive-words.txt', encoding='iso-8859-1') # Change the directory
negativewrds = read_words(r'D:/USER/Desktop/Projects/NLP/MasterDictionary/negative-words.txt', encoding='iso-8859-1') # Change the directory

# Combine all stop words into a single list
stop_words = stpwrdsgenericlong + stpwrdsauditor + stpwrdscurrencies + stpwrdsdatesandnumbers + stpwrdsgeographic + stpwrdsnames + stpwrdsgeneric

# Convert list to set for faster lookups
stop_words = set(stop_words)

# Function to remove those stopwords from positive & negative words
def remove_common_elements(list1, list2):
    for elem in list1:
        if elem in list2:
            list1.remove(elem)
    return list1

positivewrds = remove_common_elements(positivewrds, stop_words)
negativewrds = remove_common_elements(negativewrds, stop_words)

# Initialize scores dictionary
score_dict = {}

# Get the set of stopwords
stop_words = set(stopwords.words('english'))

# Define personal pronouns
personal_pronouns = ['i', 'we', 'my', 'ours', 'us']

# Iterate over the values of text_dict
for key, data in article_dict.items():
    # Tokenize the text into words and sentences
    words = word_tokenize(data['content'].lower())
    sentences = sent_tokenize(data['content'])

    # Initialize scores
    positive_score = 0
    negative_score = 0

    # Calculate Positive and Negative Scores
    for word in words:
        if word in positivewrds:
            positive_score += 1
        elif word in negativewrds:
            negative_score += 1

    # Calculate Polarity and Subjectivity Scores
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(words) + 0.000001)

    # Remove stopwords and punctuation
    cleaned_words = [word for word in words if word not in stop_words and word.isalnum()]

    # Calculate the number of cleaned words
    cleaned_word_count = len(cleaned_words)

    # Calculate the number of words and sentences
    word_count = len(words)
    sentence_count = len(sentences)

    # Calculate the average sentence length
    average_sentence_length = word_count / sentence_count

    # Calculate the syllable count for all words
    syllable_count = [len(set(re.findall(r'[aeiouy]+', word))) for word in words]

    # Calculate the number of complex words
    complex_words_count = sum(1 for word in words if len(word) > 6 and syllable_count[words.index(word)] >= 3)

    # Calculate the percentage of complex words
    percentage_complex_words = complex_words_count / word_count * 100

    # Calculate the average number of words per sentence
    average_words_per_sentence = word_count / sentence_count

    # Calculate the Fog Index
    fog_index = 0.4 * (average_sentence_length + percentage_complex_words)

    # Calculate the syllable count per words
    syllable_count_per_word = [sum(1 for char in word if char.lower() in 'aeiou') - (len(re.findall(r'[^aeiou][eo]$', word.lower())) if re.search(r'[^aeiou][eo]$', word.lower()) else 0) for word in words]
    syllable_count_per_word = mean(syllable_count_per_word)

    # Calculate personal pronouns
    personal_pronouns_count = sum(1 for word in words if word in personal_pronouns and word != 'US')

    # Calculate the total number of characters in all words
    total_characters = sum(len(word) for word in words)

    # Calculate the average word length
    average_word_length = total_characters / word_count

    # Save the scores in score_dict
    score_dict[key] = {'positive_score': positive_score, 'negative_score': negative_score,'polarity_score': polarity_score, 'subjectivity_score': subjectivity_score,
                       'average_sentence_length': average_sentence_length, 'complex_words_count': complex_words_count,'percentage_complex_words': percentage_complex_words,
                       'average_words_per_sentence': average_words_per_sentence, 'fog_index': fog_index, 'word_count': cleaned_word_count, 'syllable_count_per_word': syllable_count_per_word,
                       'personal_pronouns_count': personal_pronouns_count, 'average_word_length': average_word_length}

# List to save the score
positive_score_list = []
negative_score_list = []
polarity_score_list = []
subjectivity_score_list = []
average_sentence_length_list = []
percentage_complex_words_list = []
fog_index_list = []
average_words_per_sentence_list = []
complex_words_count_list = []
cleaned_word_count_list = []
syllable_count_per_word_list = []
personal_pronouns_count_list = []
average_word_length_list = []

# Print the scores
for key, value in score_dict.items():
  positive_score_list.append(value['positive_score'])
  negative_score_list.append(value['negative_score'])
  polarity_score_list.append(value['polarity_score'])
  subjectivity_score_list.append(value['subjectivity_score'])
  average_sentence_length_list.append(value['average_sentence_length'])
  complex_words_count_list.append(value['complex_words_count'])
  percentage_complex_words_list.append(value['percentage_complex_words'])
  average_words_per_sentence_list.append(value['average_words_per_sentence'])
  fog_index_list.append(value['fog_index'])
  cleaned_word_count_list.append(value['word_count'])
  syllable_count_per_word_list.append(value['syllable_count_per_word'])
  personal_pronouns_count_list.append(value['personal_pronouns_count'])
  average_word_length_list.append(value['average_word_length'])

# Update it in the dataframe
columns_to_update = ['POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE',
                     'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX',
                     'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 'WORD COUNT',
                     'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']

for column, score_list in zip(columns_to_update, [positive_score_list, negative_score_list, polarity_score_list,
                                                  subjectivity_score_list, average_sentence_length_list,
                                                  percentage_complex_words_list, fog_index_list,
                                                  average_words_per_sentence_list, complex_words_count_list,
                                                  cleaned_word_count_list, syllable_count_per_word_list,
                                                  personal_pronouns_count_list, average_word_length_list]):
    df_output[column] = score_list
df_output.to_excel('Output Data Structure.xlsx', engine = 'openpyxl')