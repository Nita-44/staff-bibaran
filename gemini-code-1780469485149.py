import pandas as pd

# Load the Excel file to inspect its sheets and structure
file_path = "कर्मचारी.xlsx"
try:
    xls = pd.ExcelFile(file_path)
    print("Sheet names:", xls.sheet_names)
    
    # Read the first sheet to see what it looks like
    df = pd.read_excel(file_path, sheet_name=0)
    print("\nFirst 5 rows of the data:")
    print(df.head())
    print("\nData info:")
    print(df.info())
except Exception as e:
    print("Error:", e)

Code output
Sheet names: ['Sheet1', 'Sheet2']

First 5 rows of the data:
  स्थानीय तहको नामः रोहिणी गाउँपालिकाका धकधई, रुपन्देही Unnamed: 1              Unnamed: 2 Unnamed: 3              Unnamed: 4          Unnamed: 5         Unnamed: 6  Unnamed: 7
0                      स्थायी कर्मचारी (स्वास्थ्य तर्फ)        NaN                     NaN        NaN                     NaN                 NaN                NaN         NaN
1                                                क्र.स.  संकेत नं.                  नाम थर         तह                   सेवा                 समुह                पद   सम्पर्क न.
2                                                     1     161006  राम नेवास प्रसाद चौधरी       सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन   स्वास्थ्य अधिकृत   9847023055
3                                                     2     172257        तारा प्रसाद गैरे       सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन   स्वास्थ्य अधिकृत   9847054056
4                                                     3     158044      बुद्धि प्रसाद पन्थ       सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन   स्वास्थ्य अधिकृत   9847046633

Data info:
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 84 entries, 0 to 83
Data columns (total 8 columns):
 #   Column                                                 Non-Null Count  Dtype 
---  ------                                                 --------------  ----- 
 0   स्थानीय तहको नामः रोहिणी गाउँपालिकाका धकधई, रुपन्देही  83 non-null     object
 1   Unnamed: 1                                             77 non-null     object
 2   Unnamed: 2                                             77 non-null     object
 3   Unnamed: 3                                             77 non-null     object
 4   Unnamed: 4                                             77 non-null     object
 5   Unnamed: 5                                             77 non-null     object
 6   Unnamed: 6                                             77 non-null     object
 7   Unnamed: 7                                             77 non-null     object
dtypes: object(8)
memory usage: 5.4+ KB
None

Python
# Let's inspect Sheet 2 as well to understand the complete dataset
df2 = pd.read_excel("कर्मचारी.xlsx", sheet_name='Sheet2')
print("Sheet2 First 5 rows:")
print(df2.head(10))
print("\nSheet2 info:", df2.info())

Code output
Sheet2 First 5 rows:
  करार तथा ज्यालादारी कर्मचारी              Unnamed: 1            Unnamed: 2   Unnamed: 3 Unnamed: 4
0                        क्र.स.                 नाम थर                   पद   सम्पर्क नं.     कैफियत
1                             1             निता पौडेल         सूचना प्रविधि   9847324645       करार
2                             2    विशाल सिहं क्षेत्री             ईन्जिनियर   9867419784       करार
3                             3           धमेन्द्र पाल         रोजगार संयोजक   9806997600       करार
4                             4  सुरेन्द्र कुमार चौधरी      एम.आई.एस. अपरेटर   9867420735       करार
5                             5          धमेन्द्र यादव                  अमिन   9804435165       करार
6                             6        विजय कुमार यादव           फिल्ड सहायक   9804445984       करार
7                             7             कमला चौधरी  उद्यम विकास सहजकर्ता   9817555900       करार
8                             8      राजेश कुमार मिश्र  उद्यम विकास सहजकर्ता   9848022882       करार
9                             9        अनिल कुमार यादव         ना.प.से.प्रा.   9807436234       करार
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 65 entries, 0 to 64
Data columns (total 5 columns):
 #   Column                         Non-Null Count  Dtype 
---  ------                         --------------  ----- 
 0   करार तथा ज्यालादारी कर्मचारी   65 non-null     object
 1   Unnamed: 1                     64 non-null     object
 2   Unnamed: 2                     64 non-null     object
 3   Unnamed: 3                     60 non-null     object
 4   Unnamed: 4                     64 non-null     object
dtypes: object(5)
memory usage: 2.7+ KB

Sheet2 info: None

Python
# Let's clean up Sheet 1
df1_raw = pd.read_excel("कर्मचारी.xlsx", sheet_name='Sheet1')

# The main title seems to be at row 0 (or column name)
title_1 = df1_raw.columns[0]
subtitle_1 = df1_raw.iloc[0, 0]

# Row 1 contains the headers
headers_1 = df1_raw.iloc[1].tolist()
print("Headers 1:", headers_1)

# Slice data from row 2 onwards
df1_clean = df1_raw.iloc[2:].copy()
df1_clean.columns = headers_1
df1_clean = df1_clean.dropna(subset=['क्र.स.']) # remove any empty or category separator rows if they exist, let's verify what data is inside

print(df1_clean.head(10))
print("Total rows in cleaned Sheet 1:", len(df1_clean))

Code output
Headers 1: ['क्र.स.', 'संकेत नं.', 'नाम थर', 'तह', 'सेवा ', 'समुह', 'पद ', 'सम्पर्क न.']
   क्र.स. संकेत नं.                  नाम थर    तह                   सेवा                 समुह                         पद   सम्पर्क न.
2       1    161006  राम नेवास प्रसाद चौधरी  सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन            स्वास्थ्य अधिकृत   9847023055
3       2    172257        तारा प्रसाद गैरे  सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन            स्वास्थ्य अधिकृत   9847054056
4       3    158044      बुद्धि प्रसाद पन्थ  सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन            स्वास्थ्य अधिकृत   9847046633
5       4    162396  डाजेन्द्र प्रसाद नेपाल  सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन            स्वास्थ्य अधिकृत   9841980523
6       5    159595           नारायण पाण्डे  सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन            स्वास्थ्य अधिकृत   9847039841
7       6    200289         उदय राज अधिकारी  सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन            स्वास्थ्य अधिकृत   9857016612
8       7    190181          रेनुका पाण्डेय  सातौ  स्थानीय स्वास्थ्य सेवा             प.हे.न            स्वास्थ्य अधिकृत   9857034434
9       8    189707         बालाकृष्ण पौडेल  सातौ  स्थानीय स्वास्थ्य सेवा         मे.ल्या.टे            स्वास्थ्य अधिकृत   9847640628
10      9    182651   शान्ता कुमारी भट्टराई  सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन            स्वास्थ्य अधिकृत   9847054926
11     10    181841          चन्द्रीका यादव  सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन   ज.स्वास्थ्य अधिकृत (पदनाम)  9857015223
Total rows in cleaned Sheet 1: 81

Python
# Let's inspect unique entries in 'तह' or 'पद ' or if there are subheadings inside Sheet 1 data
print(df1_clean['तह'].value_counts(dropna=False))
print(df1_clean['पद '].value_counts(dropna=False).head(10))

Code output
तह
पाँचौ          36
छैटौं          18
सातौ           10
NaN             5
तह              5
आठौँ            4
चौथो            2
रा.प. तृतिय     1
Name: count, dtype: int64
पद 
सि.अ.हे.व.                10
सि.अ.न.मी                  9
स्वास्थ्य अधिकृत           9
बरिष्ठ प्रशासन सहायक       7
कार्यालय सहयोगी            6
NaN                        5
पद                         5
सि.अ.न.मी                  3
जन स्वास्थ निरीक्षक        2
सि.अ.न.मी निरीक्षक         2
Name: count, dtype: int64

Python
# Let's print the rows where 'तह' is NaN or 'तह' to see if there are subheaders inside the data rows
print(df1_clean[df1_clean['तह'].isna() | (df1_clean['तह'] == 'तह')])

Code output
            क्र.स.  संकेत नं.  नाम थर   तह  सेवा   समुह  पद   सम्पर्क न.
43    प्रशासन तर्फ        NaN     NaN  NaN    NaN   NaN  NaN         NaN
44          क्र.स.  संकेत नं.  नाम थर   तह  सेवा   समुह  पद   सम्पर्क न.
66     शिक्षा तर्फ        NaN     NaN  NaN    NaN   NaN  NaN         NaN
67          क्र.स.  संकेत नं.  नाम थर   तह  सेवा   समुह  पद   सम्पर्क न.
70  प्राविधिक तर्फ        NaN     NaN  NaN    NaN   NaN  NaN         NaN
71          क्र.स.  संकेत नं.  नाम थर   तह  सेवा   समुह  पद   सम्पर्क न.
77       कृषी तर्फ        NaN     NaN  NaN    NaN   NaN  NaN         NaN
78          क्र.स.  संकेत नं.  नाम थर   तह  सेवा   समुह  पद   सम्पर्क न.
81  पशु तथा मत्स्य        NaN     NaN  NaN    NaN   NaN  NaN         NaN
82          क्र.स.  संकेत नं.  नाम थर   तह  सेवा   समुह  पद   सम्पर्क न.

Python
# Ah, Sheet 1 has multiple service categories embedded as subheadings! 
# Let's write a script to properly categorize Sheet 1 rows based on these subheadings.

current_category = "स्वास्थ्य तर्फ" # Default starting category based on row 0
sheet1_data = []

for idx, row in df1_raw.iterrows():
    if idx == 0 or idx == 1:
        continue
    val_0 = str(row.iloc[0]).strip()
    val_1 = row.iloc[1]
    
    # If it's a category header row
    if pd.isna(val_1) and not pd.isna(row.iloc[0]):
        current_category = val_0
        continue
    # If it's a repeated column header row
    if val_0 == 'क्र.स.':
        continue
        
    if pd.isna(row.iloc[0]) and pd.isna(row.iloc[1]) and pd.isna(row.iloc[2]):
        continue
        
    row_dict = row.to_dict()
    # map to standard headers
    cleaned_row = {
        'क्र.स.': row.iloc[0],
        'संकेत नं.': row.iloc[1],
        'नाम थर': row.iloc[2],
        'तह': row.iloc[3],
        'सेवा': row.iloc[4],
        'समुह': row.iloc[5],
        'पद': row.iloc[6],
        'सम्पर्क नं.': row.iloc[7],
        'किसिम': 'स्थायी',
        'सेवा समूह': current_category
    }
    sheet1_data.append(cleaned_row)

df1_final = pd.DataFrame(sheet1_data)
print(df1_final['सेवा समूह'].value_counts())
print("\nSample records:")
print(df1_final.head(2))

Code output
सेवा समूह
स्वास्थ्य तर्फ    40
प्रशासन तर्फ      21
प्राविधिक तर्फ     5
शिक्षा तर्फ        2
कृषी तर्फ          2
पशु तथा मत्स्य     1
Name: count, dtype: int64

Sample records:
   क्र.स. संकेत नं.                  नाम थर    तह                    सेवा                समुह                 पद सम्पर्क नं.   किसिम       सेवा समूह
0       1    161006  राम नेवास प्रसाद चौधरी  सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन   स्वास्थ्य अधिकृत   9847023055  स्थायी  स्वास्थ्य तर्फ
1       2    172257        तारा प्रसाद गैरे  सातौ  स्थानीय स्वास्थ्य सेवा  हेल्थ ईन्स्पेक्सन   स्वास्थ्य अधिकृत   9847054056  स्थायी  स्वास्थ्य तर्फ

Python
# Let's clean up Sheet 2 data (करार तथा ज्यालादारी कर्मचारी)
# Let's print out all unique values in column 0 or see if there are subheaders
print(df2.head(15))
print(df2.iloc[:, 0].value_counts().head(10))

Code output
   करार तथा ज्यालादारी कर्मचारी              Unnamed: 1            Unnamed: 2   Unnamed: 3 Unnamed: 4
0                         क्र.स.                 नाम थर                   पद   सम्पर्क नं.     कैफियत
1                              1             निता पौडेल         सूचना प्रविधि   9847324645       करार
2                              2    विशाल सिहं क्षेत्री             ईन्जिनियर   9867419784       करार
3                              3           धमेन्द्र पाल         रोजगार संयोजक   9806997600       करार
4                              4  सुरेन्द्र कुमार चौधरी      एम.आई.एस. अपरेटर   9867420735       करार
5                              5          धमेन्द्र यादव                  अमिन   9804435165       करार
6                              6        विजय कुमार यादव           फिल्ड सहायक   9804445984       करार
7                              7             कमला चौधरी  उद्यम विकास सहजकर्ता   9817555900       करार
8                              8      राजेश कुमार मिश्र  उद्यम विकास सहजकर्ता   9848022882       करार
9                              9        अनिल कुमार यादव         ना.प.से.प्रा.   9807436234       करार
10                            10      बिनोद कुमार ठाकुर         ना.प.से.प्रा.   9847032110       करार
11                            11          माधव न्यौपाने         ना.प.से.प्रा.   9819456619       करार
12                            12        हरि नारायण थारु         ना.प.से.प्रा.   9807425448       करार
13                            13           सरस्वती खनाल    स.कम्प्यूटर अपरेटर   9867575060       करार
14                            14             मुना वि.क.    स.कम्प्यूटर अपरेटर   9814415830       करार
करार तथा ज्यालादारी कर्मचारी 
क्र.स.    2
2         2
3         2
4         2
5         2
6         2
1         2
38        1
42        1
41        1
Name: count, dtype: int64

Python
# Let's find out rows where "करार तथा ज्यालादारी कर्मचारी" has headers or weird text
print(df2[df2['करार तथा ज्यालादारी कर्मचारी'] == 'क्र.स.'])
print(df2[df2['करार तथा ज्यालादारी कर्मचारी'].isna() | df2['करार तथा ज्यालादारी कर्मचारी'].str.contains('[A-Za-z\u0900-\u097F]', na=False)].head(20))

Code output
Traceback (most recent call last):
  File "<xbox-string>", line 2, in <module>
    print(df2[df2['करार तथा ज्यालादारी कर्मचारी'] == 'क्र.स.'])
  File "frame.py", line 3760, in __getitem__
    indexer = self.columns.get_loc(key)
  File "base.py", line 3654, in get_loc
    raise KeyError(key) from err
KeyError: 'करार तथा ज्यालादारी कर्मचारी'

The above exception was the direct cause of the following exception:

KeyError: 'करार तथा ज्यालादारी कर्मचारी'

Python
# Let's see the column names of df2 again
print(df2.columns)
# Let's parse df2 safely
sheet2_data = []
current_cat_2 = "करार"

for idx, row in df2.iterrows():
    val_0 = str(row.iloc[0]).strip()
    val_1 = str(row.iloc[1]).strip()
    
    if idx == 0:
        continue # 'क्र.स.', 'नाम थर', etc.
        
    if val_0 == 'क्र.स.':
        continue
        
    # Check if this row is a text section boundary
    if pd.isna(row.iloc[1]) or val_1 == 'nan' or val_1 == '':
        if not pd.isna(row.iloc[0]) and val_0 != 'nan':
            current_cat_2 = val_0
        continue
        
    sheet2_data.append({
        'क्र.स.': row.iloc[0],
        'नाम थर': row.iloc[1],
        'पद': row.iloc[2],
        'सम्पर्क नं.': row.iloc[3],
        'किसिम': row.iloc[4] if not pd.isna(row.iloc[4]) else current_cat_2,
        'सेवा समूह': current_cat_2
    })

df2_final = pd.DataFrame(sheet2_data)
print(df2_final.head(10))
print("\nUnique types in Sheet 2:")
print(df2_final['किसिम'].value_counts())
print(df2_final['सेवा समूह'].value_counts())

Code output
Index(['करार तथा ज्यालादारी कर्मचारी ', 'Unnamed: 1', 'Unnamed: 2',
       'Unnamed: 3', 'Unnamed: 4'],
      dtype='object')
   क्र.स.                 नाम थर                    पद सम्पर्क नं. किसिम सेवा समूह
0       1             निता पौडेल         सूचना प्रविधि  9847324645  करार      करार
1       2    विशाल सिहं क्षेत्री             ईन्जिनियर  9867419784  करार      करार
2       3           धमेन्द्र पाल         रोजगार संयोजक  9806997600  करार      करार
3       4  सुरेन्द्र कुमार चौधरी      एम.आई.एस. अपरेटर  9867420735  करार      करार
4       5          धमेन्द्र यादव                  अमिन  9804435165  करार      करार
5       6        विजय कुमार यादव           फिल्ड सहायक  9804445984  करार      करार
6       7             कमला चौधरी  उद्यम विकास सहजकर्ता  9817555900  करार      करार
7       8      राजेश कुमार मिश्र  उद्यम विकास सहजकर्ता  9848022882  करार      करार
8       9        अनिल कुमार यादव         ना.प.से.प्रा.  9807436234  करार      करार
9      10      बिनोद कुमार ठाकुर         ना.प.से.प्रा.  9847032110  करार      करार

Unique types in Sheet 2:
किसिम
करार                51
दैनिक ज्यालादारी    11
Name: count, dtype: int64
सेवा समूह
करार                                               56
अन्य निकायबाट यस स्थानीय तहमा परिचालित कर्मचारी     6
Name: count, dtype: int64

Python
# Let's inspect rows from df2_final where service group is different to see what it is
print(df2_final[df2_final['सेवा समूह'] != 'करार'])

Code output
    क्र.स.              नाम थर                       पद  सम्पर्क नं. किसिम                                        सेवा समूह
56       1         राधा आर.सी.        सामाजिक परिचारिका   9866146925  करार  अन्य निकायबाट यस स्थानीय तहमा परिचालित कर्मचारी
57       2      शुसिला ज्ञवाली  अ.न.मी चौथो(रोमीग नर्स)   9867821209  करार  अन्य निकायबाट यस स्थानीय तहमा परिचालित कर्मचारी
58       3      शान्ता अधिकारी  अ.न.मी चौथो(रोमीग नर्स)   9867278463  करार  अन्य निकायबाट यस स्थानीय तहमा परिचालित कर्मचारी
59       4   डा.अजय कुमार पासी           मेडिकल अधिकृत   ९८१६४३८९९४   करार  अन्य निकायबाट यस स्थानीय तहमा परिचालित कर्मचारी
60       5        प्रकाश वि.क.          ना.प्रा.स. चौथो   9800696668  करार  अन्य निकायबाट यस स्थानीय तहमा परिचालित कर्मचारी
61       6  मेघनाथ प्रसाद यादव            पशुसेवा प्रा.   ९८२१२८३७१७  करार  अन्य निकायबाट यस स्थानीय तहमा परिचालित कर्मचारी

Python
# Let's count totals to present a summary dashboard
print("Total Permanent Staff:", len(df1_final))
print("Total Contract/Wage Staff:", len(df2_final))

import matplotlib.pyplot as plt
import seaborn as sns

# Set font that supports Nepali if possible, or use standard lookups
# Since we just want a simple and beautiful clean summary table and look, let's format it beautifully.
# Let's merge both dataframes into a master list for easy search/display/systemization.
df1_mod = df1_final[['नाम थर', 'पद', 'तह', 'सम्पर्क नं.', 'किसिम', 'सेवा समूह']].copy()
df2_mod = df2_final[['नाम थर', 'पद', 'किसिम', 'सेवा समूह']].copy()
df2_mod['तह'] = 'करार/ज्यालादारी'
df2_mod['सम्पर्क नं.'] = df2_final['सम्पर्क नं.']

master_df = pd.concat([df1_mod, df2_mod], ignore_index=True)
print("\nMaster dataset shape:", master_df.shape)
print(master_df['किसिम'].value_counts())

Code output
Total Permanent Staff: 71
Total Contract/Wage Staff: 62

Master dataset shape: (133, 6)
किसिम
स्थायी              71
करार                51
दैनिक ज्यालादारी    11
Name: count, dtype: int64

Python
# Let's see why Permanent Staff dropped from 81 to 71. Ah, let's recheck length of df1_final
print(len(df1_final))
# Let's print out the full distribution of types and summaries so we can make a beautiful dashboard text.
print(master_df.groupby(['किसिम', 'सेवा समूह']).size())

Code output
71
किसिम             सेवा समूह                                      
करार              अन्य निकायबाट यस स्थानीय तहमा परिचालित कर्मचारी     6
                  करार                                               45
दैनिक ज्यालादारी  करार                                               11
स्थायी            कृषी तर्फ                                           2
                  पशु तथा मत्स्य                                      1
                  प्रशासन तर्फ                                       21
                  प्राविधिक तर्फ                                      5
                  शिक्षा तर्फ                                         2
                  स्वास्थ्य तर्फ                                     40
dtype: int64
import pandas as pd

# (Assuming df is processed as shown in the script above)
# You can quickly save a simple view to HTML using:
master_df = pd.read_excel("कर्मचारी.xlsx") # Replace with cleaned dataframe logic

html_content = """
<html>
<head>
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
    <script type="text/javascript" src="https://code.jquery.com/jquery-3.7.0.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <style> body { font-family: Arial, sans-serif; margin: 30px; } </style>
</head>
<body>
    <h2>कर्मचारी विवरण प्रणाली</h2>
    """ + master_df.to_html(classes='display', id='empTable', index=False) + """
    <script>
        $(document).ready(function() {
            $('#empTable').DataTable({ "pageLength": 25 });
        });
    </script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
print("Frontend saved as index.html. Just double-click it to view!")
