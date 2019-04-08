import pandas as pd
import csv, re

# used to take an array of crieria to substitute the matches with a particular word
def multireplace(string, replacements):
    """
    Given a string and a replacement map, it returns the replaced string.
    :param str string: string to execute replacements on
    :param dict replacements: replacement dictionary {'value to find': 'value to replace'}
    :rtype: str
    """
    # Place longer ones first to keep shorter substrings from matching where the longer ones should take place
    # For instance given the replacements {'ab': 'AB', 'abc': 'ABC'} against the string 'hey abc', it should produce
    # 'hey ABC' and not 'hey ABc'
    substrs = sorted(replacements, key=len, reverse=True)

    # Create a big OR regex that matches any of the substrings to replace
    regexp = re.compile('|'.join(map(re.escape, substrs)))

    # For each match, look up the new string in the replacements
    return regexp.sub(lambda match: replacements[match.group(0)], string)

def removeMD(name):
    name.replace('-MD', '')
    return name

def removeKD(name):
    name.replace('-KD', '')
    return name

def removePopBrands(name):
    return name

def cleanseItmName(name):
    name = removeMD(name)
    name = removeKD(name)
    name = removePopBrands(name)

    return name



#### CONFIGURATION ####
of = 'out.csv'
inf = r'C:\Users\Richard\Documents\GitHub\arbysEodReportParser\Product Mix.xlsx'
data = pd.read_excel (inf, header=5)
df = pd.DataFrame(data, columns= ['Item Name', 'Quantity', 'Total'])
#print(df) 

# gets headers from inf
headers = df.dtypes.index
#print(headers)

# Loads log file
writer = csv.DictWriter(open(of, 'w',encoding='UTF-8', newline=''),headers)

# inits the dictionary
productData = dict()

for index, row in df.iterrows():
    print(row['Item Name'])

    # skips NaN rows
    if isinstance(row['Item Name'], str):
        print(row['Item Name'] + ' is not a nan')
        # checks if item already exists in dict
        if row['Item Name'] in productData:

            #otherwise, update existing rows
            print('a new one!')
            productData[(row['Item Name'], 'Quantity')] = productData[(row['Item Name'], 'Quantity')] + int(row['Quantity'])
            productData[(row['Item Name'], 'Total')] = productData[(row['Item Name'], 'Total')] + float(row['Total'])


        else:
            
            # if not, just add data to dictionary
            productData[(row['Item Name'], 'Item Name')] = row['Item Name']
            productData[(row['Item Name'], 'Quantity')] = int(row['Quantity'])
            productData[(row['Item Name'], 'Total')] = float(row['Total']) # float() not best for $, but since number of ops is small, error is small
            
            #writer.writerow({'Item Name': row['Item Name']})

print(productData)






            
