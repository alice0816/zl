# -*- coding: utf-8 -*-
import os
import re
import xlwt
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
#     try:
#         sortfile = file(outputFile, 'wb')  # output file
#     except:
#         pass
#     writer = csv.writer(sortfile)
#     writer.writerow([Item, Quantity, Reference, Part, PCB_Footprint])

    mergeValues = mergeRows().split('\n')[:-1]
    rows = []
    for r in mergeValues:
        tempRow = r.split('&&' )
        q = int(tempRow[0])  # swicth quantity from str to int
        tempRow[0] = q
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
    
    '''set the cell style'''
    def set_style(borders=True, bold=False, alignment=False):
        style = xlwt.XFStyle()
        if borders:
            borders = xlwt.Borders()
            borders.left = xlwt.Borders.THIN
            borders.right = xlwt.Borders.THIN
            borders.top = xlwt.Borders.THIN
            borders.bottom = xlwt.Borders.THIN
            style.borders = borders
        if bold:
            font = xlwt.Font()
            font.bold = bold
            style.font = font
        if alignment:
            alignment_center  = xlwt.Alignment()
            alignment_center.horz  = xlwt.Alignment.HORZ_CENTER
            style.alignment = alignment_center
        return style
        
    '''write date to xls file'''
    f = xlwt.Workbook()
    bom1 = f.add_sheet('bom 1', cell_overwrite_ok=True)
    row0 = [u'序号', u'数量', u'元件位号', u'规格型号', u'元件封装']
    for i in range(0,len(row0)): 
        alignment_status = False
        if i < 2:
            alignment_status = True
        bom1.write(0,i,row0[i], set_style(bold=True, alignment=alignment_status))

    i ,j = 1 , 0
    if len(uFirstResult):
        bom1.write_merge(i, i+0 , j, j+ 4, u'芯片', set_style(bold=True, alignment=True))
        i += 1
        n = 0
        for result in uFirstResult: 
            n += 1
            bom1.write(i, 0, n, set_style(alignment=True))  # add Item
            bom1.write(i, 1, result[0], set_style(alignment=True)) #add data
            for k in range(0, len(result)-1): # add datas
                bom1.write(i, k + 2, result[k+1], set_style())
            i += 1
                
    if len(cFirstResult):
        bom1.write_merge(i, i+0 , j, j+ 4, u'电容', set_style(bold=True, alignment=True))
        i += 1
        n = 0
        for result in cFirstResult:
            n += 1
            bom1.write(i, 0, n, set_style(alignment=True))  # add Item
            bom1.write(i, 1, result[0], set_style(alignment=True)) #add data
            for k in range(0, len(result)-1): # add datas
                bom1.write(i, k + 2, result[k+1], set_style())
            i += 1
                
    if len(rFirstResult):
        bom1.write_merge(i, i+0 , j, j+ 4, u'电阻', set_style(bold=True, alignment=True))
        i += 1
        n = 0
        for result in rFirstResult:
            n += 1
            bom1.write(i, 0, n, set_style(alignment=True))  # add Item
            bom1.write(i, 1, result[0], set_style(alignment=True)) #add data
            for k in range(0, len(result)-1): # add datas
                bom1.write(i, k + 2, result[k+1], set_style())
            i += 1

    if len(dFirstResult):
        bom1.write_merge(i, i+0 , j, j+ 4, u'二极管', set_style(bold=True, alignment=True))
        i += 1
        n = 0
        for result in dFirstResult:
            n += 1
            bom1.write(i, 0, n, set_style(alignment=True))  # add Item
            bom1.write(i, 1, result[0], set_style(alignment=True)) #add data
            for k in range(0, len(result)-1): # add datas
                bom1.write(i, k + 2, result[k+1], set_style())
            i += 1

    if len(qFirstResult):
        bom1.write_merge(i, i+0 , j, j+ 4, u'晶体管', set_style(bold=True, alignment=True))
        i += 1
        n = 0
        for result in qFirstResult:
            n += 1
            bom1.write(i, 0, n, set_style(alignment=True))  # add Item
            bom1.write(i, 1, result[0], set_style(alignment=True)) #add data
            for k in range(0, len(result)-1): # add datas
                bom1.write(i, k + 2, result[k+1], set_style())
            i += 1

    if len(lFirstResult):
        bom1.write_merge(i, i+0 , j, j+ 4, u'电感', set_style(bold=True, alignment=True))
        i += 1
        n = 0
        for result in lFirstResult:
            n += 1
            bom1.write(i, 0, n, set_style(alignment=True))  # add Item
            bom1.write(i, 1, result[0], set_style(alignment=True)) #add data
            for k in range(0, len(result)-1): # add datas
                bom1.write(i, k + 2, result[k+1], set_style())
            i += 1

    if len(yFirstResult):
        bom1.write_merge(i, i+0 , j, j+ 4, u'晶振', set_style(bold=True, alignment=True))
        i += 1
        n = 0
        for result in yFirstResult:
            n += 1
            bom1.write(i, 0, n, set_style(alignment=True))  # add Item
            bom1.write(i, 1, result[0], set_style(alignment=True)) #add data
            for k in range(0, len(result)-1): # add datas
                bom1.write(i, k + 2, result[k+1], set_style())
            i += 1

    if len(jFirstResult):
        bom1.write_merge(i, i+0 , j, j+ 4, u'接插件', set_style(bold=True, alignment=True))
        i += 1
        n = 0
        for result in jFirstResult:
            n += 1
            bom1.write(i, 0, n, set_style(alignment=True))  # add Item
            bom1.write(i, 1, result[0], set_style(alignment=True)) #add data
            for k in range(0, len(result)-1): # add datas
                bom1.write(i, k + 2, result[k+1], set_style())
            i += 1

    if len(otherFirstResult):
        bom1.write_merge(i, i+0 , j, j+ 4, u'其它', set_style(bold=True, alignment=True))
        i += 1
        n = 0
        for result in otherFirstResult:
            n += 1
            bom1.write(i, 0, n, set_style(alignment=True))  # add Item
            bom1.write(i, 1, result[0], set_style(alignment=True)) #add data
            for k in range(0, len(result)-1): # add datas
                bom1.write(i, k + 2, result[k+1], set_style())
            i += 1
            
    '''set the cell width'''
    bom1.col(0).width =3333 /2
    bom1.col(1).width =3333 /2
    bom1.col(2).width =3333 *2
    bom1.col(3).width =3333 *2
    bom1.col(4).width =3333 *2
            
    try:
        f.save(outputFile)
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
            outputFile = os.path.join(os.path.dirname(sys.argv[1]), 'outputfile.xls' )

        elif length == 3:
            if os.path.isfile(sys.argv[2]):
                outputFile = sys.argv[2]
            
            elif os.path.isdir(sys.argv[2]):
                outputFile = os.path.join(sys.argv[2], 'outputfile.xls' )
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

