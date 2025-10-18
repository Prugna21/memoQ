# Download the required NLTK data or download all commonly used NLTK data at once
>>> nltk.download()
>>> import nltk
>>> nltk.download('punkt_tab')
[nltk_data] Downloading package punkt_tab to
[nltk_data]     C:\Users\...
[nltk_data]   Unzipping tokenizers\punkt_tab.zip.
True

# Tokenization
>>> from nltk.tokenize import word_tokenize
>>> text = "Who would have thought that computer programs would be analyzing human sentiments"
>>> tokens = word_tokenize(text)
>>> print(tokens)

# Output
['Who', 'would', 'have', 'thought', 'that', 'computer', 'programs', 'would', 'be', 'analyzing', 'human', 'sentiments']

>>> nltk.download('stopwords')
[nltk_data] Downloading package stopwords to
[nltk_data]     C:\Users\...
[nltk_data]   Package stopwords is already up-to-date!
True

>>> import nltk
>>> stopwords = nltk.corpus.stopwords.words('english')
>>> print(stopwords)

# Output 
['a', 'about', 'above', 'after', 'again', 'against', 'ain', 'all', 'am', 'an', 'and', 'any', 'are', 'aren', "aren't", 'a
s', 'at', 'be', 'because', 'been', 'before', 'being', 'below', 'between', 'both', 'but', 'by', 'can', 'couldn', "couldn't", 'd', 'did', 'didn', "didn't", 'do', 'does', 'doesn', "doesn't", 'doing', 'don', "don't", 'down', 'during', 'each', 'few', 'for', 'from', 'further', 'had', 'hadn', "hadn't", 'has', 'hasn', "hasn't", 'have', 'haven', "haven't", 'having', 'he', "he'd", "he'll", 'her', 'here', 'hers', 'herself', "he's", 'him', 'himself', 'his', 'how', 'i', "i'd", 'if', "i'll", "i'm", 'in', 'into', 'is', 'isn', "isn't", 'it', "it'd", "it'll", "it's", 'its', 'itself', "i've", 'just', 'll', 'm', 'ma', 'me', 'mightn', "mightn't", 'more', 'most', 'mustn', "mustn't", 'my', 'myself', 'needn', "needn't", 'no', 'nor', 'not', 'now', 'o', 'of', 'off', 'on', 'once', 'only', 'or', 'other', 'our', 'ours', 'ourselves', 'out', 'over', 'own', 're', 's', 'same', 'shan', "shan't", 'she', "she'd", "she'll", "she's", 'should', 'shouldn', "shouldn't", "should've", 'so', 'some', 'such', 't', 'than', 'that', "that'll", 'the', 'their', 'theirs', 'them', 'themselves', 'then', 'there', 'these', 'they', "they'd", "they'll", "they're", "they've", 'this', 'those', 'through', 'to', 'too', 'under', 'until', 'up', 've', 'very', 'was', 'wasn', "wasn't", 'we', "we'd", "we'll", "we're", 'were', 'weren', "weren't", "we've", 'what', 'when', 'where', 'which', 'while', 'who', 'whom', 'why', 'will', 'with', 'won', "won't", 'wouldn', "wouldn't", 'y', 'you', "you'd", "you'll", 'your', "you're", 'yours', 'yourself', 'yourselves', "you've"]                                      >>>

>>> newtokens = [word for word in tokens if word not in stopwords]
>>> print(newtokens)

# Output
['Who', 'would', 'thought', 'computer', 'programs', 'would', 'analyzing', 'human', 'sentiments']


# Lemmatization
>>> from nltk.stem import WordNetLemmatizer                                                                             
>>> text = "Who would have thought that computer programs would be analyzing human sentiments"                          
>>> tokens = word_tokenize(text)                                                                                        
>>> lemmatizer = WordNetLemmatizer()                                                                                    
>>> tokens = [lemmatizer.lemmatize(word) for word in tokens]
>>> print(tokens)

# Output 
['Who', 'would', 'have', 'thought', 'that', 'computer', 'program', 'would', 'be', 'analyzing', 'human', 'sentiment']

# Stemming
>>> from nltk.stem import PorterStemmer                                                                                 
>>> from nltk.tokenize import word_tokenize                                                                             
                                                                                                                    
>>> text = "Who would have thought that computer programs would be analyzing human sentiments"                          
>>> tokens = word_tokenize(text.lower())                                                                                
>>> ps = PorterStemmer()
>>> tokes = [ps.stem(word) for word in tokens]
>>> print(tokens)

# Output 
['who', 'would', 'have', 'thought', 'that', 'computer', 'programs', 'would', 'be', 'analyzing', 'human', 'sentiments']
