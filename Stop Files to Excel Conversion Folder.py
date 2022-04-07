# If this is your first time running this notebook, run "pip install docx2txt"
# pip install docx2txt

import docx2txt
import pandas as pd
import os
import os.path
import pathlib

os.chdir(r"C:\Users\azhang3\SFMTA\Transit Planning - Service Planning\4_Stops_Management\Stop File Replacement\Raw Stop Files")
cwd = os.getcwd()

stopfiles_dir = {}
basepath = pathlib.Path(cwd)
for item in [entry for entry in basepath.iterdir() if entry.is_file()]:
    stopfiles_dir[os.path.basename(item)] = os.path.join(cwd,item)

def readdocxfile(document): # reads docx file and returns list of lines
    # Import text from document
    doc_text = docx2txt.process(document)
    # Split by new line
    doc_lines = [l for l in doc_text.splitlines()]
    return doc_lines

def findlastof(value, l): # Look for the last instance of a value and return index 
    keeplooking = True
    for count, row in enumerate(reversed(l)):
        if value in row and keeplooking == True:
            reversecutoff = count
            keeplooking = False
    cutoff = len(l) - reversecutoff
    return cutoff

def findheadercutoff(doc_lines): # Return only header rows
    headercutoff = findlastof('-------------------------------',doc_lines)
    return doc_lines[:headercutoff]

# Deprecated code
flagwords = ['M. L. RUCUS ', 'IN EFFECT', 'LINE','C        Y','T\t\t\t STOP\t', 'STREETS  I LOC', '----------------------']

def returnineffect(txt): # Find effective date in a string
    if 'IN EFFECT' in txt:
        num = txt.find('IN EFFECT') + 9
        date = txt[num:]
        return date.strip()
    else:
        return None

def returndivisions(txt): # Find divisions in a string
    divisions = ['BEACH','CABLE','FLYNN','GREEN','ISLAIS','KIRKLAND','MME','POTRERO','PRESIDIO','WOODS']
    matcheddiv = []
    for div in divisions:
        if div in txt:
            matcheddiv.append(div)
    if len(matcheddiv) > 0:
        return matcheddiv
    else:
        return None

def processheaderrows(headderrows): # Iterate over a list of strings to return effective dates and divisions
    ineffectdates = []
    matcheddivisions = []
    for r in headderrows:    
        match = returnineffect(r)
        if match != None and match not in ineffectdates:
            ineffectdates.append(match)
        match2 = returndivisions(r)
        if match2 != None:
            for m in match2:
                if m not in matcheddivisions:
                    matcheddivisions.append(m)
    effectivedate = ','.join(ineffectdates)
    route_division = ','.join(matcheddivisions)
    return (effectivedate, route_division)

def findstoprows(doc_lines): # Return only data rows
    headercutoff = findlastof('-------------------------------',doc_lines)
    return doc_lines[headercutoff:]

def customranges(listoflengths): # Create a list of tuples from a list of lengths to serve as ranges
    tuplelist = []
    x = 0
    for n in listoflengths:
        w = x
        x += n
        tuplelist.append((w,x))
    return tuplelist
    
def delimitbycustomlength(x,customlengths,lengthknown = False): # Split up a chunk of text based on a list of lengths. If lengthknown is false, returns any characters not included as the final value
    ranges = customranges(customlengths)
    delimited = []
    for a,b in ranges:
        delimited.append(x[a:b])
    if lengthknown == True:
        return delimited
    else:
        last = ranges[-1][1]
        delimited.append(x[last:])
        return delimited

# Defining field names (this may change depending on the stop file) - global variable
fields = ['A','Trapeze Name','C-I','Location','Orientation','Type','Length','Timepoint','Distance','Sign','Stop ID','Notes']

# Defining the field lengths (this may change depending on the stop file) - global variable
fieldlengths01 = [5,10,1,3,3,3,5,1,7,5,7]

def delimitrows(stoprows): # Delimit rows and strip extra spaces
    delimitedrows = []
    for row in stoprows:
        d = delimitbycustomlength(row,fieldlengths01)
        delimitedrows.append([e.strip() for e in d])
    return delimitedrows

def allblanks(i): # Returns blank if all items in a list are blank
    if i.count('') == len(i):
        return 'Blank'
    else:
        return i

def repetitions(rows): # Finds irregular ranges of blank values
    transformed = [allblanks(r) for r in rows]
    indexed = [i for i, j in enumerate(transformed) if j == 'Blank']
    notinpattern = []
    for i,k in enumerate(indexed[:-1]):
        if k+2 != indexed[i+1]:
            notinpattern.append(k)
    return notinpattern

def cleanrows(delimitedrows): # Remove blank values, but preserve blanks if there are more than two in a row
    clean_rows = []
    kept_blanks = repetitions(delimitedrows)
    for count,row in enumerate(delimitedrows):
        if allblanks(row) != 'Blank':
            clean_rows.append(row)
        else:
            if count in kept_blanks:
                clean_rows.append(row)
    return clean_rows

# Run these to check that lines were properly delimited
# sample = doc_lines[64]
# print(sample)
# sample2 = delimitbycustomlength(sample,fieldlengths01)
# print(sample2)


# Run to check if blank rows were properly deleted
#for a in clean_rows:
#    print(a)

def findroutename(text):
    if 'stop file' in text.lower():
        num = text.lower().find('stop file')
        name = text[:num].strip()
        if name[-1] == '-':
            name = name[:-1]
        return name.strip()
    else:
        return text

def converttoexcel(clean_rows, document_name, effectivedate, route_division): # Create dataframe, add extra fields, and export to excel
    # Create dataframe
    stops_df = pd.DataFrame(clean_rows, columns = fields)
    # Add extra identifying fields
    stops_df['Effective Date'] = effectivedate
    stops_df['Division'] = route_division
    filename = document_name[document_name.find('Raw Stop Files') + 15:]
    stops_df['Route File Name'] = filename.replace('.docx','')
    stops_df['Route Name'] = findroutename(filename)
    # Check to see if DataFrame was created properly
    # Create excel file
    output_name = document_name.replace('.docx','.xlsx')
    stops_df.to_excel(output_name)
    print(output_name, ' created successfully')

def stopdocxtoexcel(location):
    doc_lines = readdocxfile(location)
    # Header rows
    headerrows = findheadercutoff(doc_lines)
    headerresult = processheaderrows(headerrows)
    effectiveon = headerresult[0]
    divisions = headerresult[1]
    # Stop rows
    stoprows = findstoprows(doc_lines)
    delimitedrows = delimitrows(stoprows)
    cleanedrows = cleanrows(delimitedrows)
    converttoexcel(cleanedrows, location, effectiveon, divisions)

for stopfile in list(stopfiles_dir.keys()):
    stopdocxtoexcel(stopfiles_dir[stopfile])
