from bs4 import BeautifulSoup;
import requests,openpyxl;
from nltk.stem import PorterStemmer
from nltk.tokenize import word_tokenize
import pandas as pd
import math,openpyxl,re,textstat
urls=pd.read_csv("input.csv")

#
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Title"
sheet.append(["URL", "POSITIVE_SCORE", "NEGATIVE_SCORE", "POLARITY_SCORE", "SUBJECTIVITY_SCORE", "AVG_SENTENCE_LENGTH",
         "PERCENTAGE_OF_COMPLEX_WORDS", "FOG_INDEX", "AVG_NUMBER_OF_WORDS_PER_SENTENCE", "COMPLEX_WORD_COUNT",
         "WORD_COUNT", "SYLLABLE_PER_WORD", "PERSONAL_PRONOUNS", "AVG_WORD_LENGTH"])

# Loop over the URLs
all_paragraphs=[]
for url in urls.URL:
    # Send a GET request to the URL
    response = requests.get(url)

    # Check if the request was successful
    if response.status_code == 200:
        # Create BeautifulSoup object
        soup = BeautifulSoup(response.text, 'html.parser')

        paragraphs = soup.find_all('p')
        all_paragraphs.append(paragraphs)
    else:
        print("Error accessing URL:", url)

for index in range(len(all_paragraphs)):
    print("Start number wise....................................................................................................................",index)
    combine_array = []
    for para in all_paragraphs[index]:
        if para.get("class") != ['entry-title', 'td-module-title']:
            if para.get("class") != ["tdm-descr"]:
                # Remove the content of <strong> tags within the <p> tag
                for strong_tag in para.find_all('strong'):
                    strong_tag.decompose()

                print("para :", para)
                texts = para.text.lower()
                texts.strip().replace("  ", " ")
                array = texts.split(".")
                clean_array = [sentence for sentence in array if sentence != '']
                print("clean_array :", clean_array)
                for arr in [clean_array]:
                    combine_array.extend(arr)

    print("combine_array :", combine_array)

    # Specify the path to the stop words list file
    stop_words_files = ["StopWords/StopWords_Auditor.txt", "StopWords/StopWords_Currencies.txt",
                        "StopWords/StopWords_DatesandNumbers.txt", "StopWords/StopWords_Generic.txt",
                        "StopWords/StopWords_GenericLong.txt", "StopWords/StopWords_Geographic.txt"]

    # Combined all stop words files into single file.
    words = []
    for file_path in stop_words_files:
        with open(file_path, "r") as file:
            words.extend(file.read().splitlines())

    stop_words = []
    for text in words:
        t = text.lower()
        t = t.replace("|", ",")
        stop_words.append(t)
    # print(stop_words)

    # //////////////////////////////////////////////////////////////////////////////////////////////////////////
    # clean data
    entire_data = []
    Total_sentence = 0
    for i in combine_array:
        array = i.split(".")
        Total_sentence += len(array)
        clean_array = [sentence for sentence in array if sentence != '']


        #     print(clean_array)
        # Cleaned the text by removing stop words
        def clean_text(text):
            text[0] = text[0].replace("“", "").replace("”", "").replace("(", "").replace(")", "").replace("‘", "")
            words = text[0].split()
            return entire_data.extend(words)


        clean_text(clean_array)
    # print(entire_data)

    # //////////////////////////////////////////////////////////////////////////////////////////////////////////
    # Clean the text by removing stop words
    clean_data = [];


    def clean_text(words):
        cleaned_words = [word for word in words if word.lower() not in stop_words]
        cleaned_text = " ".join(cleaned_words)
        return cleaned_text


    sen = clean_text(entire_data)
    clean_data.append(sen)
    # print(clean_data)
    # /////////////////////////////////////////////////////////////////////////////////////////////////////////
    # remove Digits from data
    def remove_digits(data):
        if isinstance(data, str):
            return ''.join(char for char in data if not char.isdigit())
        else:
            return data


    clean_data = [remove_digits(clean_data[0])]
    # print(clean_data)
    # /////////////////////////////////////////////////////////////////////////////////////////////////////////
    # Positive Score
    # Initialize stemming object
    stemmer = PorterStemmer()
    positive_file_path = "MasterDictionary/positive-words.txt"
    positive_score = 0
    with open(positive_file_path, "r") as file:
        positive_words = file.read().splitlines()
        # print(positive_words)
        for words in clean_data:
            # Tokenize the text
            words = word_tokenize(words)
            positive_score += sum(1 if stemmer.stem(word) in positive_words else 0 for word in words)
    print("positive_score :",positive_score)
    # //////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # Negative score
    negative_file_path = "MasterDictionary/negative-words.txt"
    negative_score = 0
    with open(negative_file_path, "r") as file:
        negative_words = file.read().splitlines()
        # print(positive_words)
        for words in clean_data:
            # print(words)
            # Tokenize the text
            words = word_tokenize(words)
            negative_score += sum(-1 if stemmer.stem(word) in negative_words else 0 for word in words)

    print("negative_score :",negative_score)

    # //////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # Polarity Score
    Polarity_Score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    print("Polarity_Score :",Polarity_Score)


    # //////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # Subjective score
    Total_words = 0
    for sentence in clean_data:
        print(sentence)
        words = sentence.split(" ")
        Total_words += len(words)
    # print("Total_words :",Total_words)
    Subjectivity_Score = (positive_score + negative_score) / (Total_words + 0.000001)
    # print("Subjectivity_Score :",Subjectivity_Score)


    # ////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # Analysis of Readability
    # Average Sentence Length = the number of words / the number of sentences
    avg_sentence_length = Total_words / Total_sentence;
    avg_sentence_length = math.floor(avg_sentence_length)
    print("avg_sentence_length :",avg_sentence_length)


    # Percentage of Complex words = the number of complex words / the number of words
    cleaned_complex_words = []


    def get_complex_words(text):
        # Remove punctuation and convert text to lowercase
        text = re.sub(r'[^\w\s]', '', text[0].lower())

        # Split text into words
        words = text.split()
        #     print(words)
        complex_words = []
        for word in words:
            syllable_count = textstat.syllable_count(word)
            if syllable_count > 2:
                complex_words.append(word)
        return complex_words


    cleaned_complex_words = get_complex_words(clean_data)

    print("length of cleaned_complex_words :",len(cleaned_complex_words))
    percentage_of_Complex_words = math.floor((len(cleaned_complex_words) / Total_words) * 100)
    print("percentage_of_Complex_words :",percentage_of_Complex_words)

    # Fog Index
    Fog_Index = 0.4 * (avg_sentence_length + percentage_of_Complex_words)
    print("Fog_Index :",math.floor(Fog_Index))


    # /////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # Average Number of Words Per Sentence
    average_Number_of_Words_Per_Sentence=0
    if Total_sentence != 0:
        average_Number_of_Words_Per_Sentence = Total_words / Total_sentence
    print("average_Number_of_Words_Per_Sentence :",average_Number_of_Words_Per_Sentence)

    # /////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # Syllable Count Per Word
    import pyphen

    # Created a Pyphen instance for syllable counting
    dic = pyphen.Pyphen(lang='en')

    # Extracts syllable words
    syllable_words = []
    for word in cleaned_complex_words:
        syllables = dic.inserted(word).count('-') + 1

        # Handling exceptions for words ending with "es" or "ed"
        if word.endswith(('es', 'ed')):
            syllables -= 1

        if syllables > 1:
            syllable_words.append(word)

    # Print the syllable words
    print("Syllable words:", len(syllable_words))

    # /////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # Personal Pronouns
    space_separated_string = " ".join(cleaned_complex_words)
    # print(space_separated_string)

    # personal pronouns to search for
    personal_pronouns = ['i', 'we', 'my', 'ours', 'us']

    lowercase_text = space_separated_string.lower()

    count = 0

    # Counting personal pronouns in the text
    for pronoun in personal_pronouns:
        # print(pronoun)
        count += lowercase_text.count(pronoun)

    # Print the count
    print("Personal Pronouns Count:", count)


    # /////////////////////////////////////////////////////////////////////////////////////////////////////////////
    #Average Word Length
    # Average Word Length= total number of characters / Total_words
    split_clean_data=clean_data[0].split()

    count_all_characters = 0;
    for i in split_clean_data:
        count_all_characters += len(i)
    print("count_all_characters :",count_all_characters)

    average_Word_Length = count_all_characters / len(split_clean_data)
    print("average_Word_Length :",average_Word_Length)

    # ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

    # All Result
    URL=urls.URL[index]
    POSITIVE_SCORE = positive_score
    NEGATIVE_SCORE = negative_score
    POLARITY_SCORE = Polarity_Score
    SUBJECTIVITY_SCORE = Subjectivity_Score
    AVG_SENTENCE_LENGTH = avg_sentence_length
    PERCENTAGE_OF_COMPLEX_WORDS = math.floor(percentage_of_Complex_words)
    FOG_INDEX = math.floor(Fog_Index)
    AVG_NUMBER_OF_WORDS_PER_SENTENCE = math.floor(average_Number_of_Words_Per_Sentence)
    COMPLEX_WORD_COUNT = len(cleaned_complex_words)
    WORD_COUNT = Total_words
    SYLLABLE_PER_WORD = len(syllable_words)
    PERSONAL_PRONOUNS = count
    AVG_WORD_LENGTH = math.floor(average_Word_Length)

    # added all data in excel file
    sheet.append([URL, POSITIVE_SCORE, NEGATIVE_SCORE, POLARITY_SCORE, SUBJECTIVITY_SCORE, AVG_SENTENCE_LENGTH,
                  PERCENTAGE_OF_COMPLEX_WORDS, FOG_INDEX, AVG_NUMBER_OF_WORDS_PER_SENTENCE, COMPLEX_WORD_COUNT,
                  WORD_COUNT, SYLLABLE_PER_WORD, PERSONAL_PRONOUNS, AVG_WORD_LENGTH])

excel.save("Completed.xlsx");