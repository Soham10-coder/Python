import pandas as pd
data = pd.read_csv("08\hello2.csv",delimiter=';')
data_top = data.head()
print(data_top )