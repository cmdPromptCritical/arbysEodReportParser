import pandas as pd
import csv, re
import pyodbc, datetime    

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

def removeSuffixes(name):
    name = name.replace('-KD', '')
    name = name.replace('-HH', '')
    name = name.replace('-SM', '')
    name = name.replace('-MD', '')
    name = name.replace('-LG', '')
    return name

def renameSnackFries(name):
    name = name.replace('Snack Curly Fry', 'Small Curly Fry')
    return name

def removeApostrophes(name):
    name = name.replace("'","")
    return name

def cleanseItmName(name):
    name = removeSuffixes(name)
    name = renameSnackFries(name)
    name = removeApostrophes(name)
    name = name.strip()
    return name

def exporttoDb(connStr, data):
    conn = pyodbc.connect(connStr)
    cursor = conn.cursor()
    x = 0
    itemName = ''
    quantity = 0
    total = 0
    for key, value in data:
        if (x==0):
            itemName = data[(key, value)]
            x += 1
        elif (x==1):
            quantity = data[(key, value)]
            x += 1
        else:
            total = data[(key, value)]
            query = "INSERT INTO sales (itemName, quantity, grossRevenue) VALUES ('" + itemName + "', " + str(quantity) + ", " + str(total) + ");"
            print(query)
            cursor.execute(query)
            conn.commit()

            x = 0


    #query = "INSERT INTO sales (itemName) VALUES ('" + time + "');"
    #print(query)





#### CONFIGURATION ####
connStr = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\Richard\Documents\RegisterPrintout_be.accdb;'
of = 'out.csv'
inf = r'C:\Users\Richard\Documents\GitHub\arbysEodReportParser\Product Mix.xlsx'
data = pd.read_excel (inf, header=5)
df = pd.DataFrame(data, columns= ['Item Name', 'Quantity', 'Total'])
drinkList = {
    'Fruitopia' : '',
    'Gold Peak Rsbry' : '',
    'Fanta Orange' : '',
    'Coke Zero Sugar' : '',
    'Cherry Coke' : '',
    'Gold Peak Lemon' : '',
    'Diet Coke' : '',
    "Barq's Root Beer" : '',
    'Sprite' : '',
    'Coke' : ''
}

# gets headers from inf
headers = df.dtypes.index
#print(headers)

# Loads log file
writer = csv.DictWriter(open(of, 'w',encoding='UTF-8', newline=''),headers)

# inits the dictionary
productData = dict()

for index, row in df.iterrows():
    #print(row['Item Name'])

    # skips NaN rows
    if isinstance(row['Item Name'], str):
        #print(row['Item Name'] + ' is not a nan')
        cleanName = cleanseItmName(row['Item Name'])
        # checks if item already exists in dict
        if (cleanName, 'Item Name') in productData.keys():

            #If so, update existing rows
            productData[(cleanName, 'Quantity')] = productData[(cleanName, 'Quantity')] + int(row['Quantity'])
            productData[(cleanName, 'Total')] = productData[(cleanName, 'Total')] + float(row['Total'])


        else:
            
            # if not, just add data to dictionary
            productData[(cleanName, 'Item Name')] = cleanName
            productData[(cleanName, 'Quantity')] = int(row['Quantity'])
            productData[(cleanName, 'Total')] = float(row['Total']) # float() not best for $, but since number of ops is small, error is small
            
#print(productData)

# deletes summary row if it was accidentally added
try:
    del productData[('Summary', 'Item Name')]
    del productData[('Summary', 'Quantity')]
    del productData[('Summary', 'Total')]
    print('sum deleted')
except:
    pass

# exports data to Access database
exporttoDb(connStr, productData)