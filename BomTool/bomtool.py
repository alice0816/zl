import csv
import os
import re
import sys

TABEL = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 
                'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
NEED_COLUMNS = ['Part Reference', 'Value', 'PCB Footprint']

def getNeedColumnNo(columnNames, totalColumn):
    needColumnNolist = []
    
    for needColumnName in NEED_COLUMNS:
        j = 0
        for i in range(totalColumn):
            if needColumnName == columnNames[i].replace('"', ''):
                if j == 0:
                    needColumnNolist.append(i)
                j += 1
    if j == 0 :
        raise Exception('No the %s column in the inputfile.' % needColumnName)
    elif j > 1:
        raise Exception('More than one column name of : %s.' % needColumnName)
    return needColumnNolist

def saveNeedColumnValues(columnNames, totalColumn, tempFile1) :
    try:
        csvfile = file(tempFile1,  'wb')
    except:
        pass
    writer = csv.writer(csvfile)
    writer.writerow(['Reference', 'Part', 'PCB Footprint'])
    needColumnNo = getNeedColumnNo(columnNames, totalColumn)
    for eachline in lines[2:] :
        needColumnValues = []
        thisColumnValues = eachline.split('\t')
        if len(thisColumnValues) != totalColumn:
            raise Exception('Less then %s columns in this row:.' % str(totalColumn))
        for columnNo in needColumnNo:
            needColumnValues.append(thisColumnValues[columnNo])
        data = [(needColumnValues[0].replace('"', '').strip()), (needColumnValues[1].replace('"', '').strip()), 
                (needColumnValues[2].replace('"', '').strip())]
        writer.writerow(data)
    try:
        csvfile.close()
    except:
        pass   
    
def mergeRows(tempFile1, tempFile2):
    try:
        mergeCsvfile = file(tempFile2, 'wb') 
        csvfile = file(tempFile1, 'r+') 
    except:
        pass
    writer = csv.writer(mergeCsvfile)
    writer.writerow(['Quantity', 'Reference', 'Part', 'PCB Footprint'])
    
    reader = csv.reader(csvfile)
    rows =  [row for row in reader][1:]
    
    '''Delete test poist and delete NC  value '''
    rowsLen = len(rows)
    for i in range(rowsLen):
        if len(rows[(rowsLen-i-1 )][0]) < 2:
            raise Exception('Reference value:"%s" is unvalid, in the %s row.' % ((rows[(rowsLen-i-1 )][0]), str(i+2)))
        elif  not re.match('^[0-9a-zA-Z]+$',(rows[(rowsLen-i-1 )][0])):
            raise Exception('Reference value:"%s" is unvalid, in the %s row.' % ((rows[(rowsLen-i-1 )][0]), str(i+2)))
        elif  not rows[(rowsLen-i-1 )][0][0].upper() in TABEL:
            raise Exception('Reference value:"%s" is unvalid, in the %s row.' % ((rows[(rowsLen-i-1 )][0]), str(i+2)))
        elif rows[(rowsLen-i-1 )][0].lower().startswith('t'):
            if not (('test' in rows[(rowsLen-i-1 )][1].lower()) or('point' in rows[(rowsLen-i-1)][1].lower())):
                raise Exception('Reference with T,but not the "Test poist" in the %s row.' % str(i+1))
            del rows[(rowsLen-i-1 )]  
        elif 'NC'  in rows[(rowsLen-i-1 )][1].upper():
            del rows[(rowsLen-i-1 )]

    '''Merge data '''
    item = 0
    while len(rows):    
        i = 0
        needDeleteRows = []
        countQuantity = 0
        row = rows[0]
        referenceValue = row[0]
        for thisrow in rows:
            if (thisrow[1] == row[1] and thisrow[2] == row[2]):
                needDeleteRows.append(i)
                countQuantity += 1
                if countQuantity > 1:
                    referenceValue = referenceValue + ',' + thisrow[0]
            i += 1
            
        item += 1
        writer.writerow([(countQuantity), (referenceValue), (row[1]), (row[2])])
        
        _leng = len(needDeleteRows)
        for i in range(_leng):
            del rows[needDeleteRows[(_leng-i-1 )]]      
    try:
        mergeCsvfile.close()
        csvfile.close()
    except:
        pass   
    
def sortAndOrderComponents(outputFile, tempFile2):
    try:
        sortfile = file(outputFile, 'wb')  # output file
        csvfile = file(tempFile2, 'r+')
    except:
        pass
    writer = csv.writer(sortfile)
    writer.writerow(['Quantity', 'Reference', 'Part', 'PCB Footprint'])

    reader = csv.reader(csvfile)
    rows =  [row for row in reader][1:]
    
    '''  first sort '''
    uFirstResult = []
    cFirstResult = [] # C
    rFirstResult = [] # R
    dFirstResult = [] # D
    otherFirstResult = [] #other

    for row in rows:
        if row[1].upper().startswith('U')  and (not row[1][1].upper() in TABEL):
            uFirstResult.append(row)
        elif row[1].upper().startswith('C') and (not row[1][1].upper() in TABEL):
            cFirstResult.append(row)
        elif row[1].upper().startswith('R') and (not row[1][1].upper() in TABEL):
            rFirstResult.append(row)
        elif row[1].upper().startswith('D') and (not row[1][1].upper() in TABEL):
            dFirstResult.append(row)
        else:
            otherFirstResult.append(row)
            
        '''order the C  according to the size of part value.'''
        for i in range(len(cFirstResult)):
            for j in range(0, len(cFirstResult)-i-1):
                a = cFirstResult[j][2].upper()
                b = cFirstResult[j+1][2].upper()
                if  'PF'  in a:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*PF.*',a, re.M|re.I)
                    a = float(matchResult.group(1))
                elif 'NF'  in a:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*NF.*',a, re.M|re.I)
                    a = float(matchResult.group(1))*1000
                elif 'UF'  in a:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*UF.*',a, re.M|re.I)
                    a = float(matchResult.group(1))*1000000
                elif re.match(r'^[0-9]+$',a) and (len(a)==3):
                    a = int(a[:2])*10**int(a[2])
                    a = float(a)
                else:
                    raise Exception('Part  vaule:%s is wrong'  % a)
                
                if  'PF'  in b:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*PF.*',b, re.M|re.I)
                    b = float(matchResult.group(1))
                elif 'NF'  in b:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*NF.*',b, re.M|re.I)
                    b = float(matchResult.group(1))*1000
                elif 'UF'  in b:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*UF.*',b, re.M|re.I)
                    b = float(matchResult.group(1))*1000000
                elif re.match(r'^[0-9]+$',b) and (len(b)==3):
                    b = int(b[:2])*10**int(b[2])
                    b = float(b)
                else:
                    raise Exception('Part  vaule:%s is wrong'  % b)

                if a > b:
                    cFirstResult[j], cFirstResult[j + 1] = cFirstResult[j + 1], cFirstResult[j]

        '''order the R  according to the size of part value.'''
        for i in range(len(rFirstResult)):
            for j in range(0, len(rFirstResult)-i-1):
                a = rFirstResult[j][2].upper()
                b = rFirstResult[j+1][2].upper()
                if  a =='0':
                    a = float(a)
                elif 'M' in a:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*M.*',a, re.M|re.I)
                    a = float(matchResult.group(1))*10**6
                elif 'K'  in a:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*K.*',a, re.M|re.I)
                    a = float(matchResult.group(1))*1000
                elif 'R'  in a:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*R.*',a, re.M|re.I)
                    a = float(matchResult.group(1))
                elif re.match(r'^[0-9]+$',a) and (len(a)==3):
                    a = int(a[:2])*10**int(a[2])
                    a = float(a)
                elif '%'  in a:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*%.*',a, re.M|re.I)
                    a = float(matchResult.group(1))
                else:
                    raise Exception('Part  vaule:%s is wrong'  % a)
                 
                if  b =='0':
                    b = float(b)
                elif 'M' in b:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*M.*',b, re.M|re.I)
                    b = float(matchResult.group(1))*10**6
                elif 'K'  in b:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*K.*',b, re.M|re.I)
                    b = float(matchResult.group(1))*1000
                elif 'R'  in b:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*R.*',b, re.M|re.I)
                    b = float(matchResult.group(1))
                elif re.match(r'^[0-9]+$',b) and (len(b)==3):
                    b = int(b[:2])*10**int(b[2])
                    b = float(b)
                elif '%'  in b:
                    matchResult = re.match(r'(\d+\.?[0-9]*).*%.*',b, re.M|re.I)
                    b = float(matchResult.group(1))
                else:
                    raise Exception('Part  vaule:%s is wrong'  % b)
 
                if a > b:
                    rFirstResult[j], rFirstResult[j + 1] = rFirstResult[j + 1], rFirstResult[j]
                    
    '''order other '''
    for i in range(len(otherFirstResult)):
        for j in range(0, len(otherFirstResult)-i-1):
            a = otherFirstResult[j][1][0].upper()
            b = otherFirstResult[j+1][1][0].upper()
            c = ord(a)
            d = ord(b)
            if c > d:
                otherFirstResult[j], otherFirstResult[j + 1] = otherFirstResult[j + 1], otherFirstResult[j]
                               
    allResult = uFirstResult
    for result in cFirstResult:
        allResult.append(result)
    for result in rFirstResult:
        allResult.append(result)
    for result in dFirstResult:
        allResult.append(result)
    for result in otherFirstResult:
        allResult.append(result)

    for result in allResult:
        writer.writerow(result)

    try:
        sortfile.close()
        csvfile.close()
    except:
        pass   

if __name__ == "__main__" :
    length = len(sys.argv) 
    if length < 2 or length > 3:
        raise Exception, 'The length of argvs is improper.'
    
    elif os.path.isfile(sys.argv[1]):

        inputFile = sys.argv[1]
        print 'Input File : '  + inputFile

        if length ==2:
            print os.path.dirname(sys.argv[1])
            outputFile = os.path.join(os.path.dirname(sys.argv[1]), 'outputfile.csv' )

        elif length == 3:
            if os.path.isfile(sys.argv[2]):
                outputFile = sys.argv[2]
            
            elif os.path.isdir(sys.argv[2]):
                outputFile = os.path.join(sys.argv[2], 'outputfile.csv' )
            else:
                raise Exception, 'The third argv is invalid.'
        print 'Output File : ' + outputFile
        tempFile1 = os.path.join(os.path.dirname(outputFile), 'temp.csv')
        tempFile2 = os.path.join(os.path.dirname(outputFile), 'temp2.csv')
                                    
    else:
        print 'Usage: bomtool .py <input file>  [output file]'   +  '\n'

    try:
        f = open(inputFile, 'r') # input file
        lines = f.readlines()
    finally:
        f.close()
    columnline = lines[1]
    columnNames = columnline.split('\t')
    totalColumn = len(columnNames)
    saveNeedColumnValues(columnNames, totalColumn, tempFile1)    
    mergeRows(tempFile1, tempFile2)  
    sortAndOrderComponents(outputFile, tempFile2)

