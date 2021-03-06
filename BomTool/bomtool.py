# -*- coding: utf-8 -*-

import csv
import os
import re
import sys
import StringIO

TABEL = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 
                'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
NEED_COLUMNS = ['part reference', 'value', 'pcb footprint']

def getNeedColumnNo(columnNames, totalColumn):
    needColumnNolist = []
    
    for needColumnName in NEED_COLUMNS:
        j = 0
        for i in range(totalColumn):
            if needColumnName == columnNames[i].strip().replace('"', '').replace('\n', '').replace('\r', '').lower():
                if j == 0:
                    needColumnNolist.append(i)
                j += 1
    if j == 0 :
        raise Exception('No the %s column in the inputfile.' % needColumnName)
    elif j > 1:
        raise Exception('More than one column name of : %s.' % needColumnName)
    return needColumnNolist

def saveNeedColumnValues() :
    s1 = StringIO.StringIO()
    needColumnNo = getNeedColumnNo(columnNames, totalColumn)
    for eachline in lines[2:] :
        needColumnValues = []
        thisColumnValues = eachline.split('\t')
        if len(thisColumnValues) != totalColumn:
            raise Exception('Less then %s columns in this row:.' % str(totalColumn))
        for columnNo in needColumnNo:
            needColumnValues.append(thisColumnValues[columnNo])
        data = (needColumnValues[0].replace('"', '').strip().replace('\n', '').replace('\r', '')) +'&&'  + (needColumnValues[1].replace('"', '').strip().replace('\n', '').replace('\r', '')) +'&&'  + (needColumnValues[2].replace('"', '').strip().replace('\n', '').replace('\r', ''))
        s1.write(data)
        s1.write('\n')
    return s1.getvalue()

def mergeRows():
    s2 = StringIO.StringIO()
    needValues = saveNeedColumnValues()
    _rows =  needValues.split('\n')[:-1]
    
    rows = []
    
    for r in _rows:
        tempRow = r.split('&&' )
        rows.append(tempRow)
        
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
            if (thisrow[1].upper() == row[1].upper() and thisrow[2].upper() == row[2].upper()):
                needDeleteRows.append(i)
                countQuantity += 1
                if countQuantity > 1:
                    referenceValue = referenceValue + ',' + thisrow[0]
            i += 1
            
        item += 1
        data = str(countQuantity) + '&&' + (referenceValue)+'&&' +(row[1]) +'&&' + (row[2])
        s2.write(data)
        s2.write('\n')
        _leng = len(needDeleteRows)
        for i in range(_leng):
            del rows[needDeleteRows[(_leng-i-1 )]]     
    return s2.getvalue() 

def sortAndOrderComponents():
    try:
        sortfile = file(outputFile, 'wb')  # output file
    except:
        pass
    writer = csv.writer(sortfile)
    Item = '类型'
    c_unicode = Item.decode( "utf -8")
    Item = c_unicode.encode( "gbk ")
    
    Quantity = '数量'
    c_unicode = Quantity.decode( "utf -8")
    Quantity = c_unicode.encode( "gbk ")

    Reference = '元件位号'
    c_unicode = Reference.decode( "utf -8")
    Reference = c_unicode.encode( "gbk ")
    
    Part = '规格型号'
    c_unicode = Part.decode( "utf -8")
    Part = c_unicode.encode( "gbk ")
    
    PCB_Footprint = '元件封装'
    c_unicode = PCB_Footprint.decode( "utf -8")
    PCB_Footprint = c_unicode.encode( "gbk ")
    writer.writerow([Item, Quantity, Reference, Part, PCB_Footprint])

    mergeValues = mergeRows().split('\n')[:-1]
    rows = []
    for r in mergeValues:
        tempRow = r.split('&&' )
        rows.append(tempRow)

    '''  first sort '''
    uFirstResult = []
    cFirstResult = [] # C
    rFirstResult = [] # R
    dFirstResult = [] # D
    qFirstResult = [] # Q
    lFirstResult = [] # L
    yFirstResult = [] # Y
    jFirstResult = [] # J
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
        elif row[1].upper().startswith('Q') and (not row[1][1].upper() in TABEL):
            qFirstResult.append(row)
        elif row[1].upper().startswith('L') and (not row[1][1].upper() in TABEL):
            lFirstResult.append(row)
        elif row[1].upper().startswith('Y') and (not row[1][1].upper() in TABEL):
            yFirstResult.append(row)
        elif row[1].upper().startswith('J') and (not row[1][1].upper() in TABEL):
            jFirstResult.append(row)
        else:
            otherFirstResult.append(row)
            
        '''order the C  according to the size of part value.'''
        for i in range(len(cFirstResult)):
            for j in range(0, len(cFirstResult)-i-1):
                a = cFirstResult[j][2].upper()
                b = cFirstResult[j+1][2].upper()
                if  'PF'  in a:
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*PF.*',a, re.M|re.I)
                        a = float(matchResult.group(1))
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % a) 
                    
                elif 'NF'  in a:
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*NF.*',a, re.M|re.I)
                        a = float(matchResult.group(1))*1000
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % a) 
                    
                elif 'UF'  in a:
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*UF.*',a, re.M|re.I)
                        a = float(matchResult.group(1))*1000000
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % a) 
                    
                elif re.match(r'^[0-9]+$',a) and (len(a)==3):
                    a = int(a[:2])*10**int(a[2])
                    a = float(a)
                else:
                    raise Exception('Part  vaule:%s is wrong'  % a)
                
                if  'PF'  in b:
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*PF.*',b, re.M|re.I)
                        b = float(matchResult.group(1))
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % b) 
                    
                elif 'NF'  in b:
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*NF.*',b, re.M|re.I)
                        b = float(matchResult.group(1))*1000
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % b) 
                    
                elif 'UF'  in b:
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*UF.*',b, re.M|re.I)
                        b = float(matchResult.group(1))*1000000
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % b) 
                    
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
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*M.*',a, re.M|re.I)
                        a = float(matchResult.group(1))*10**6
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % a) 
                    
                elif 'K'  in a:
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*K.*',a, re.M|re.I)
                        a = float(matchResult.group(1))*1000
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % a) 
                    
                elif 'R'  in a:
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*R.*',a, re.M|re.I)
                        a = float(matchResult.group(1))
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % a) 
                    
                elif re.match(r'^[0-9]+$',a) and (len(a)==3):
                    a = int(a[:2])*10**int(a[2])
                    a = float(a)
                elif '%'  in a:
                    try: 
                        matchResult = re.match(r'(\d+\.?[0-9]*).*%.*',a, re.M|re.I)
                        a = float(matchResult.group(1))
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % a) 
                else:
                    raise Exception('Part  vaule:%s is wrong'  % a)
                 
                if  b =='0':
                    b = float(b)
                elif 'M' in b:
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*M.*',b, re.M|re.I)
                        b = float(matchResult.group(1))*10**6
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % b) 
                    
                elif 'K'  in b:
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*K.*',b, re.M|re.I)
                        b = float(matchResult.group(1))*1000
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % b) 
                    
                elif 'R'  in b:
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*R.*',b, re.M|re.I)
                        b = float(matchResult.group(1))
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % b) 
                    
                elif re.match(r'^[0-9]+$',b) and (len(b)==3):
                    b = int(b[:2])*10**int(b[2])
                    b = float(b)
                elif '%'  in b:
                    try:
                        matchResult = re.match(r'(\d+\.?[0-9]*).*%.*',b, re.M|re.I)
                        b = float(matchResult.group(1))
                    except:
                        raise Exception('Part  vaule:%s is wrong'  % b) 
                    
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
    i = 0 
    if len(uFirstResult):
        content = '芯片'
        c_unicode = content.decode( "utf -8")
        c_gbk = c_unicode.encode( "gbk ")
        writer.writerow([c_gbk])       
    for result in uFirstResult:
        i += 1
        result.insert(0, i)
        writer.writerow(result)
        
    i = 0
    if len(cFirstResult):
        content = '电容'
        c_unicode = content.decode( "utf -8")
        c_gbk = c_unicode.encode( "gbk ")
        writer.writerow([c_gbk])
    for result in cFirstResult:
        i += 1
        result.insert(0, i)
        writer.writerow(result)
        
    i = 0
    if len(rFirstResult):
        content = '电阻'
        c_unicode = content.decode( "utf -8")
        c_gbk = c_unicode.encode( "gbk ")
        writer.writerow([c_gbk])  
    for result in rFirstResult:
        i += 1
        result.insert(0, i)
        writer.writerow(result)
    
    i = 0
    if len(dFirstResult):
        content = '二极管'
        c_unicode = content.decode( "utf -8")
        c_gbk = c_unicode.encode( "gbk ")
        writer.writerow([c_gbk])      
    for result in dFirstResult:
        i += 1
        result.insert(0, i)
        writer.writerow(result)
        
    i = 0
    if len(qFirstResult):
        content = '晶体管'
        c_unicode = content.decode( "utf -8")
        c_gbk = c_unicode.encode( "gbk ")
        writer.writerow([c_gbk])    
    for result in qFirstResult:
        i += 1
        result.insert(0, i)
        writer.writerow(result)
        
    i = 0
    if len(lFirstResult):
        content = '电感'
        c_unicode = content.decode( "utf -8")
        c_gbk = c_unicode.encode( "gbk ")
        writer.writerow([c_gbk])    
    for result in lFirstResult:
        i += 1
        result.insert(0, i)
        writer.writerow(result)
        
    i = 0
    if len(yFirstResult):
        content = '晶振'
        c_unicode = content.decode( "utf -8")
        c_gbk = c_unicode.encode( "gbk ")
        writer.writerow([c_gbk])    
    for result in yFirstResult:
        i += 1
        result.insert(0, i)
        writer.writerow(result)
        
    i = 0
    if len(jFirstResult):
        content = '接插件'
        c_unicode = content.decode( "utf -8")
        c_gbk = c_unicode.encode( "gbk ")
        writer.writerow([c_gbk])    
    for result in jFirstResult:
        i += 1
        result.insert(0, i)
        writer.writerow(result)
    
    i = 0
    if len(otherFirstResult):
        content = '其他'
        c_unicode = content.decode( "utf -8")
        c_gbk = c_unicode.encode( "gbk ")
        writer.writerow([c_gbk])
    for result in otherFirstResult:
        i += 1
        result.insert(0, i)
        writer.writerow(result)

    try:
        sortfile.close()
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
            outputFile = os.path.join(os.path.dirname(sys.argv[1]), 'outputfile.csv' )

        elif length == 3:
            if os.path.isfile(sys.argv[2]):
                outputFile = sys.argv[2]
            
            elif os.path.isdir(sys.argv[2]):
                outputFile = os.path.join(sys.argv[2], 'outputfile.csv' )
            else:
                raise Exception, 'The third argv is invalid.'
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
    sortAndOrderComponents()
    print 'Output File : ' + outputFile

