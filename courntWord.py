import pandas as pd

# Load the Excel files
file1 = 'Suspected.xlsx'
file2 = 'Diagnosed.xlsx'
file3 = 'Normal1.xlsx'
file4 = 'Normal2.xlsx'

df1 = pd.read_excel(file1, engine='openpyxl')
df2 = pd.read_excel(file2, engine='openpyxl')
df3 = pd.read_excel(file3, engine='openpyxl')
df4 = pd.read_excel(file4, engine='openpyxl')

# Function to count words and letters
def count_words_and_letters(text):
    words = len(text.split())
    letters = len(text.replace(" ", ""))
    return words, letters

# Add the new columns
for df in [df1, df2,df3,df4]:
    df['num of character'] = df['Text'].apply(lambda x: count_words_and_letters(x)[1] if pd.notnull(x) else 0)
    df['num of word'] = df['Text'].apply(lambda x: count_words_and_letters(x)[0] if pd.notnull(x) else 0)

# Save the updated Excel files
df1.to_excel('SuspectedCounted.xlsx', index=False)
df2.to_excel('DiagnosedCounted.xlsx', index=False)
df3.to_excel('Normal1Counted.xlsx', index=False)
df4.to_excel('Normal2Counted.xlsx', index=False)
