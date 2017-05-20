# -*- coding:utf-8 -*-
'''
Created on 2017/4/12

@author: alice
'''

import sqlite3
import xlwt

class SqlOperater(object):
    def __init__(self, db):
        self._db = db

    def connect_datebase(self):
        conn = sqlite3.connect( self._db)
        conn.row_factory = sqlite3.Row
        return conn

# cursor.execute('select * from result where TestID=? and ScriptID=? and Loop=?' , ('44','1', '1'))
    def execute_select_sql(self, sql):
        conn = self.connect_datebase()
        cursor = conn.cursor()
        cursor.execute(sql)
        values= cursor.fetchall()
        cursor.close()
        conn.close()
        return values

class ScriptModule(SqlOperater):
    def __init__(self, db):
        self._db = db
        self.nameIDDict = {}
        self.module = []  # use for order.
       
        sql = 'select * from script'
        self._script_values = self.execute_select_sql(sql)

        self.getModuleIdAndName()

    def getModuleIdAndName(self):
        sql = 'select * from module'
        values = self.execute_select_sql(sql)
        for row in values:
            self.nameIDDict[row[0]] = row[1]  # The key is moduleId, value is moduleName
            self.module.append(row[1])

    def getScriptIdAndModule(self):
        moduleIDDict = {}
        for key, value in self.nameIDDict.items():
            moduleIDDict[value] = []  # must init dict firstly.
        for row in self._script_values:
           
            for key , value in self.nameIDDict.items():
                if row[ 12] == key:
                    moduleIDDict[value].append(row[ 0])
        return moduleIDDict # The key is moduleName, value is ScriptId list
   
    def getScriptIDAndScriptName(self):
        scriptIDAndScriptName = {} # The key is scriptID, value is ScriptName
        for row in self._script_values:
            scriptIDAndScriptName[row[ 0]] = row[ 1]
        return scriptIDAndScriptName

class DevicesModule(SqlOperater):
    def __init__(self, db):
        self._db = db
        self._tagIdAndDevicesSeries = {}
        self._stageIdAndDevicesSeries = {}

        self.getTagIdAndDevicesSeries()
        self.getStageIdAndDevicesSeries()

    def getTagIdAndDevicesSeries(self):
        sql = 'select * from tag'
        values = self.execute_select_sql(sql)
        for row in values:
            self._tagIdAndDevicesSeries[row[0]] = row[1]  # The key is tagId, value is devicesSeries

    def getStageIdAndDevicesSeries(self):
        sql = 'select * from stage'
        values = self.execute_select_sql(sql)

        for key, value in self._tagIdAndDevicesSeries.items():
            self._stageIdAndDevicesSeries[value] = []  # must init dict firstly.

        for row in values:
            for key , value in self._tagIdAndDevicesSeries.items():
                if row[ 2] == key:
                    self._stageIdAndDevicesSeries[value].append(row[ 0])

    def getDeviceIdAndDevicesSeries(self):
        sql = 'select * from device'
        values = self.execute_select_sql(sql)
        deviceIdAndDevicesSeries = {}

        for row in values:
            for key , value in self._stageIdAndDevicesSeries.items():
                if row[ 2] in value:
                    deviceIdAndDevicesSeries[row[ 0]] = key
        return deviceIdAndDevicesSeries

class Result(SqlOperater):
    def __init__(self, db, testId):
        self._db = db
        self._testId = testId
        self._devicesId = []
        self._scriptId = []
        self._loopId = []
        self._scriptAndModule = {} # the key is module, value is scriptsName.
        self._sm = ScriptModule(self._db)
        self._moduleScriptIdDict = self._sm.getScriptIdAndModule()
        self._scriptIDAndScriptName = self._sm.getScriptIDAndScriptName()
       
        self._module = self._sm.module # user for order
       
        self._devices_module = DevicesModule(self ._db)
        self._deviceIdAndDevicesSeriesDict = self._devices_module.getDeviceIdAndDevicesSeries()
       
        self._existDevices = []
        self._existModule = []

        self.getbasicinfo()
        self.sortscripts()
       
        self.getExistDevices()
        self.getExistModule()

    def getbasicinfo(self):
        sql = 'select * from result where TestID=' + str(self._testId)
        values = self.execute_select_sql(sql)
        for row in values:
            if row[ 2] not in self._scriptId:
                self._scriptId.append(row[2])

            if row[ 4] not in self._devicesId:
                self._devicesId.append(row[4])

            if row[ 9] not in self._loopId:
                self._loopId.append(row[9])

    def sortscripts(self):
        allModuleId = self._sm.getScriptIdAndModule()
        for key, value in allModuleId.items():
            i = 0
            for scriptId in self._scriptId:
                if scriptId in value:
                    if i == 0:
                        self._scriptAndModule[key] = []
                    self._scriptAndModule[key].append(self._scriptIDAndScriptName[scriptId])
                    i += 1
                   
    def getExistDevices(self):
        for d in self._devicesId:
            self._existDevices.append(self._deviceIdAndDevicesSeriesDict[int(d)])
                   
    def getExistModule(self):
        keys = []
        for key in self._scriptAndModule:
            keys.append(key)
       
        for m in self._module:
            if m in keys:
                self._existModule.append(m)

    def getValueByDevicesAndScript(self):
        self._valueByDevicesAndScript = []
        for device in self._devicesId:
            for script in self._scriptId:
                for loop in self._loopId:

                    '''device, script, loop, count_success, count_failed, count_kill, total_success_time,
                    total_failed_time, total_kill_time, sucess_failed_time, total_all_time.'''
                    temp = []

                    sql = 'select * from result where TestID=' +self._testId + ' and ScriptID=' +str(script) +' and Devices=' + str(device) + ' and Loop='  + str(loop)
                    values = self.execute_select_sql(sql)
#                     for i in values:
#                         print i

                    if values:
                        count_success = 0 #
                        count_failed = 0
                        count_kill = 0
                        total_success_time = 0
                        total_failed_time = 0
                        total_kill_time = 0

                        for v in values:
                            if v[ 7] == 1:
                                count_success += 1
                                total_success_time += v[ 10]
                            elif v[ 7] == 2:
                                count_failed += 1
                                total_failed_time += v[ 10]
                            elif v[ 7] == 3:
                                count_kill += 1
                                total_kill_time += v[ 10]

                        sucess_failed_time = total_success_time + total_failed_time
                        total_all_time = sucess_failed_time + total_kill_time
                       
                        device_name = self._deviceIdAndDevicesSeriesDict[int(device)]
                        script_name = self._scriptIDAndScriptName[script]

                        temp.append(device_name)
                        temp.append(script_name)
                        temp.append(loop)
                        temp.append(count_success)
                        temp.append(count_failed)
                        temp.append(count_kill)
                       
                        temp.append((float(total_success_time)/ 3600/ 24))
                        temp.append((float(total_failed_time)/ 3600/ 24))
                        temp.append((float(total_kill_time)/ 3600/ 24))
                        temp.append((float(sucess_failed_time)/ 3600/ 24))
                        temp.append((float(total_all_time)/ 3600/ 24))

                        self._valueByDevicesAndScript.append(temp)
        return self._valueByDevicesAndScript

    def getValueByDevicesAndModule(self):
        self._ValueByDevicesAndModule = []
        _valueByDevicesAndScript = self.getValueByDevicesAndScript()
       
        for device in self._devicesId:
            for loop in self._loopId:
#                 for key, value in self._scriptAndModule.items():
                for module in self._module:
                    if not (module in self._existModule):
                        continue
                       
                    '''device, module, loop, count_success, count_failed, count_kill, total_success_time,
                    total_failed_time, total_kill_time, sucess_failed_time, total_all_time.'''
                    temp = []
                    count_success = 0 #
                    count_failed = 0
                    count_kill = 0
                    total_success_time = 0
                    total_failed_time = 0
                    total_kill_time = 0
                    sucess_failed_time = 0
                    total_all_time = 0
                   
                    for tempvalue in _valueByDevicesAndScript:
                        if (tempvalue[ 0] == self._deviceIdAndDevicesSeriesDict[int(device)] and tempvalue[ 2]==loop and (tempvalue[1] in self._scriptAndModule[module])):
                            count_success += tempvalue[ 3]
                            count_failed += tempvalue[ 4]
                            count_kill += tempvalue[ 5]
                            total_success_time += tempvalue[ 6]
                            total_failed_time += tempvalue[ 7]
                            total_kill_time += tempvalue[ 8]
                            sucess_failed_time += tempvalue[ 9]
                            total_all_time += tempvalue[ 10]
                   
                    if count_success + count_failed + count_kill != 0:
                        device_name = self._deviceIdAndDevicesSeriesDict[int(device)]
                       
                        temp.append(device_name)
                        temp.append(module)
                        temp.append(loop)
                        temp.append(count_success)
                        temp.append(count_failed)
                        temp.append(count_kill)
                        temp.append(total_success_time)
                        temp.append(total_failed_time)
                        temp.append(total_kill_time)
                        temp.append(sucess_failed_time)
                        temp.append(total_all_time)
                       
                        self._ValueByDevicesAndModule.append(temp)

        return self._ValueByDevicesAndModule
   
class ExportReport(object):
    def __init__(self, _valueByDevicesAndScript, _ValueByDevicesAndModule, _existDevices, _existModule, _loopId):
        self._valueByDevicesAndScript = _valueByDevicesAndScript
        self._ValueByDevicesAndModule = _ValueByDevicesAndModule
       
        self._existDevices = _existDevices
        self._existModule = _existModule
        self._loopId = _loopId

    def exportResult(self):   
        style = xlwt.XFStyle()
        style.num_format_str = 'h:mm:ss'
        
        '''write date1 to xls file'''
        f = xlwt.Workbook()
        result1 = f.add_sheet( 'result1', cell_overwrite_ok=True )
        row0 = [ 'device', 'script' , 'loop', 'count_success', 'count_failed', 'count_kill', 'total_success_time' ,
                'total_failed_time', 'total_kill_time' , 'sucess_failed_time', 'total_all_time']

        for i in range( 0,len(row0)):
            result1.write( 0,i,row0[i])
           
        if self._valueByDevicesAndScript:
           
            i =  0
            for value in self._valueByDevicesAndScript:
                i += 1
                for j in range(len(value)):
                    result1.write(i, j,value[j])
                   
                   
        '''write date2 to xls file'''
        result2 = f.add_sheet( 'result2', cell_overwrite_ok=True )
        result3 = f.add_sheet( 'result3', cell_overwrite_ok=True )

        rownum = 0
        for d in self._existDevices:
            result2.write(rownum, 0, d)
            result3.write(rownum, 0, d)
           
            rownum += 2
            loop_row = rownum - 1
           
            for m in self._existModule:
                result2.write(rownum, 0, m)
                result3.write(rownum, 0, m)
               
                loop_colum = 0
                for loop in self._loopId:
                    loop_colum += 1
                    result2.write(loop_row, loop_colum, loop)
                    result3.write(loop_row, loop_colum, loop)

                    for value in self._ValueByDevicesAndModule:
                        if value[ 0] == d and value[ 1] == m and value[ 2] == loop:
                                result2.write(rownum, loop_colum, value[ 9], style)
                                result3.write(rownum, loop_colum, value[ 3])

                rownum += 1
            rownum += 3
                           
        f.save(r'D:\Alice\Project\Python\output_report\22_4.18.xls')

if __name__ == '__main__':

    db = r'D:\Alice\Project\Python\output_report\AutoMan.atm'
    s = Result(db, '44')
   
    getValueByDevicesAndScript = s.getValueByDevicesAndScript()
    getValueByDevicesAndModule = s.getValueByDevicesAndModule()
   
    report = ExportReport(getValueByDevicesAndScript, getValueByDevicesAndModule, s._existDevices, s._existModule, s._loopId)
    report.exportResult()
    
    