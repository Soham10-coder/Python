from nltk.corpus import stopwords 
from nltk import word_tokenize

text = "This is a simple sentence demonstrating the removal of stop words."
tokens = word_tokenize(text)
stop_words = set(stopwords.words('english'))

filtered_words = [word for word in tokens if word.lower() not in stop_words]
print(filtered_words)
