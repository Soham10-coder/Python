import pandas as pd
df = pd.read_csv(r"08/hello2.csv", delimiter=';')
ser_first_name = pd.Series(df['First name'])
print(ser_first_name)

ser = pd.DataFrame(df[['First name','Last name','Department']])
print(ser)


