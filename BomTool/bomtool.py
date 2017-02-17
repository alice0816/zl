import csv

needColumns = ['Part Reference', 'Value', 'PCB Footprint']

try:
    f = open(r'D:\Alice\Project\Python\AnalysisData\x1000.EXP','r') # input file
    lines = f.readlines()
finally:
    f.close()
columnline = lines[1]
columnNames = columnline.split('\t')
totalColumn = len(columnNames)


def getNeedColumnNo():
    needColumnNolist = []
    
    for needColumnName in needColumns:
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

def saveNeedColumnValues():
    
    csvfile = file(r'D:\Alice\Project\Python\AnalysisData\tmp.csv', 'wb')
    writer = csv.writer(csvfile)
    writer.writerow(['Reference', 'Part', 'PCB Footprint'])
    needColumnNo = getNeedColumnNo()
    for eachline in lines[2:] :
        needColumnValues = []
        thisColumnValues = eachline.split('\t')
        if len(thisColumnValues) != totalColumn:
            raise Exception('Less then %s columns in this row:.' % str(totalColumn))
        for columnNo in needColumnNo:
            needColumnValues.append(thisColumnValues[columnNo])
        data = [(needColumnValues[0].replace('"', '')), (needColumnValues[1].replace('"', '')), (needColumnValues[2].replace('"', ''))]
        writer.writerow(data)
    csvfile.close()   
    
def mergeRows():
    
    mergeCsvfile = file(r'D:\Alice\Project\Python\AnalysisData\x1000.csv', 'wb')  
    writer = csv.writer(mergeCsvfile)
    writer.writerow(['Quantity', 'Reference', 'Part', 'PCB Footprint'])
    
    csvfile = file(r'D:\Alice\Project\Python\AnalysisData\tmp.csv', 'r+')
    reader = csv.reader(csvfile)
    rows =  [row for row in reader][1:]
    
    '''Delete test poist '''
    rowsLen = len(rows)
    for i in range(rowsLen):

        if rows[(rowsLen-i-1 )][0].lower().startswith('t'):
            if not (('test' in rows[(rowsLen-i-1 )][1].lower()) or('point' in rows[(rowsLen-i-1)][1].lower())):
                print rows[(rowsLen-i-1 )][0]
                raise Exception('Reference with T,but not the "Test poist"!!!!')
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
            
    mergeCsvfile.close()
    csvfile.close()
    
def sortComponents():
    sortfile = file(r'D:\Alice\Project\Python\AnalysisData\sort.csv', 'wb')  # output file
    writer = csv.writer(sortfile)
    writer.writerow(['Quantity', 'Reference', 'Part', 'PCB Footprint'])
    
    csvfile = file(r'D:\Alice\Project\Python\AnalysisData\x1000.csv', 'r+')
    reader = csv.reader(csvfile)
    rows =  [row for row in reader][1:]
    
    orderBasis = ['U', 'C', 'R', 'D', 'A', 'B', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 
                'N', 'O', 'P', 'Q', 'S', 'T', 'V', 'W', 'X', 'Y', 'Z']
    for order in orderBasis:
        firstOrderResult = []
        for row in rows:
            if row[1].upper().startswith(order):
                firstOrderResult.append(row)
                
        for i in range(len(firstOrderResult)):
            for j in range(i):
                a = firstOrderResult[j][1][1]
                b = firstOrderResult[j+1][1][1]
                if (a.upper() in orderBasis):
                    a = ord(a.upper())
                else:
                    a = int(a)
                if (b.upper() in orderBasis):
                    b = ord(b.upper())
                else:
                    b = int(b)
                    
                if a > b:
                    firstOrderResult[j], firstOrderResult[j + 1] = firstOrderResult[j + 1], firstOrderResult[j]
        
        for result in firstOrderResult:
            writer.writerow(result)
                  
saveNeedColumnValues()    
mergeRows()  
sortComponents() 
