from nltk.stem import PorterStemmer

stemmer = PorterStemmer()
words = ["running", "ran", "easily", "fairly"]
stems = [stemmer.stem(word) for word in words]
print(stems)
