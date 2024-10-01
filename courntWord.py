import pandas as pd

# Load the Excel files
file1 = 'Suspected.xlsx'
file2 = 'Diagnosed.xlsx'
file3 = 'normal1.xlsx'
file4 = 'normal2.xlsx'

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
    df['word'] = df['Text'].apply(lambda x: count_words_and_letters(x)[0] if pd.notnull(x) else 0)
    df['character'] = df['Text'].apply(lambda x: count_words_and_letters(x)[1] if pd.notnull(x) else 0)

# Save the updated Excel files
df1.to_excel('Suspected1.xlsx', index=False)
df2.to_excel('sureCounted2.xlsx', index=False)
df3.to_excel('normalCounted111.xlsx', index=False)
df4.to_excel('normalCounted222.xlsx', index=False)
