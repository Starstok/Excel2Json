#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xlrd
import json

def read_excel(inPutFile, sysType):
    workbook = xlrd.open_workbook(inPutFile)    # 打开文件
    sheetCount = len(workbook.sheets())         #sheet数量
    sheet = workbook.sheet_by_index(0)          # 获取第一个sheet
    # sheet的名称，行数，列数
    print("表格名:", sheet.name, "行数:", sheet.nrows, "列数:", sheet.ncols)
    
    if sheet.ncols < 5: # 如果表格少于5列，返回1
        return 1

    if sysType == "通用下发":
        roomInfo = {}
        room = []
        ads = {}
        devInfo = {}
        device = []
        array = {}
        dataInfo = {}
        array["dic"] = dataInfo				    # array字典嵌套dic并赋值dataInfo
        for i in range( 1, sheet.nrows ): 		# 从第二行开始循环
            typeInfo = conve(sheet.cell(i, 0).value)	# 获取第1列的值
            typeInfo = str(typeInfo)
            dataInfo[typeInfo] = {}				# 在dataInfo下嵌套typeInfo并赋值{}
  
            name = conve(sheet.cell(i, 1).value)	    # 获取第2列的值
            name = str(name)

            mark = conve(sheet.cell(i, 2).value)	    # 获取第3列的值
            mark = str(mark)
            dataInfo[typeInfo][name] = mark     # 在dataInfo/typeInfo下嵌套name并赋值mark
        device.append(array)
        devInfo['device'] = device
        devInfo['permission'] = 1
        ads['analog_digital_string'] = devInfo
        ads['room_id'] = 0
        ads['room_name'] = "通用下发"
        
        room.append(ads)
        roomInfo['room'] = room
        roomInfo['security'] = []
        return roomInfo

    elif sysType == "通用上报":
        roomInfo = {}
        room = []
        ads = {}
        devInfo = {}
        device = []
        for rown in range(1, sheet.nrows):
            array = {}
            array['device_id'] = conve(sheet.cell_value(rown, 0))
            array['alias'] = conve(sheet.cell_value(rown, 1))
            array['name'] = conve(sheet.cell_value(rown, 2))
            array['pid'] = conve(sheet.cell_value(rown, 3))
            array['modelId'] = conve(sheet.cell_value(rown, 4))
            array['poiCode'] = conve(sheet.cell_value(rown, 5))
            array['datapoint_report'] = conve(sheet.cell_value(rown, 6))
            array['sn'] = conve(sheet.cell_value(rown, 7))
            device.append(array)
            devInfo['device'] = device
            devInfo['permission'] = 1
        ads['analog_digital_string'] = devInfo
        ads['room_id'] = 0
        ads['room_name'] = "通用上报"
        
        room.append(ads)
        roomInfo['room'] = room
        roomInfo['security'] = []
        return roomInfo

    elif sysType == "通用消防":
        roomInfo = {}
        room = []
        fire = {}
        devInfo = {}
        device = []
        for rown in range(1, sheet.nrows):
            array = {}
            array['device_id'] = conve(sheet.cell_value(rown, 0))
            array['card_id'] = conve(sheet.cell_value(rown, 1))
            array['point_id'] = conve(sheet.cell_value(rown, 2))
            array['alias'] = conve(sheet.cell_value(rown, 3))
            array['name'] = conve(sheet.cell_value(rown, 4))
            array['pid'] = conve(sheet.cell_value(rown, 5))
            array['modelId'] = conve(sheet.cell_value(rown, 6))
            array['poiCode'] = conve(sheet.cell_value(rown, 7))
            array['appType'] = conve(sheet.cell_value(rown, 8))
            array['datapoint_report'] = conve(sheet.cell_value(rown, 9))
            array['address'] = conve(sheet.cell_value(rown, 10))
            array['sn'] = conve(sheet.cell_value(rown, 11))
            device.append(array)
            devInfo['device'] = device
            devInfo['permission'] = 1
        fire['analog_digital_string'] = devInfo
        fire['room_id'] = 0
        fire['room_name'] = "通用消防系统"
        
        room.append(fire)
        roomInfo['room'] = room
        roomInfo['security'] = []
        return roomInfo

    elif sysType == "标准消防":
        roomInfo = {}
        room = []
        fire = {}
        devInfo = {}
        device = []
        for rown in range(1, sheet.nrows):
            array = {}
            array['device_id'] = conve(sheet.cell_value(rown, 0))
            array['card_id'] = conve(sheet.cell_value(rown, 1))
            array['point_id'] = conve(sheet.cell_value(rown, 2))
            array['name'] = conve(sheet.cell_value(rown, 3))
            array['address'] = conve(sheet.cell_value(rown, 4))
            array['pid'] = conve(sheet.cell_value(rown, 5))
            array['modelId'] = conve(sheet.cell_value(rown, 6))
            array['poiCode'] = conve(sheet.cell_value(rown, 7))
            array['datapoint_report'] = conve(sheet.cell_value(rown, 8))
            array['prefix'] = conve(sheet.cell_value(rown, 9))
            device.append(array)
            devInfo['device'] = device
            devInfo['permission'] = 1
        fire['fire_control'] = devInfo
        fire['room_id'] = 0
        fire['room_name'] = "消防系统"
        
        room.append(fire)
        roomInfo['room'] = room
        roomInfo['security'] = []
        return roomInfo

    elif sysType == "标准照明":
        roomInfo = {}
        room = []
        lighting = {}
        devInfo = {}
        device = []
        for i in range(sheetCount):
            sheet = workbook.sheet_by_index(i)          # 获取全部sheet
            chnlInfo = {}
            channel = []
            for rown in range(1, sheet.nrows):
                array = {}
                array['device_id'] = conve(sheet.cell_value(rown, 0))
                array['channel'] = conve(sheet.cell_value(rown, 1))
                array['light_id'] = conve(sheet.cell_value(rown, 2))
                array['name'] = conve(sheet.cell_value(rown, 3))
                array['type'] = conve(sheet.cell_value(rown, 4))
                array['sn'] = conve(sheet.cell_value(rown, 5))
                array['icon'] = conve(sheet.cell_value(rown, 6))
                array['index'] = conve(sheet.cell_value(rown, 7))
                channel.append(array)
            chnlInfo['channel'] = channel
            chnlInfo['device_id'] = conve(sheet.cell_value(rown, 0))
            chnlInfo['name'] = sheet.name
            chnlInfo['sync_type'] = conve(sheet.cell_value(rown, 8))
            chnlInfo['type'] = conve(sheet.cell_value(rown, 4))
            device.append(chnlInfo)
            devInfo['device'] = device
            devInfo['permission'] = 1
        lighting['lighting'] = devInfo
        lighting['room_id'] = 0
        lighting['room_name'] = "照明系统"

        room.append(lighting)
        roomInfo['room'] = room
        roomInfo['security'] = []
        return roomInfo

    elif sysType == "标准场景":
        roomInfo = {}
        room = []
        sceneInfo = {}
        scene = []
        for rown in range(1, sheet.nrows):
            array = {}
            array['scene_id'] = conve(sheet.cell_value(rown, 0))
            array['name'] = conve(sheet.cell_value(rown, 1))
            array['sn'] = conve(sheet.cell_value(rown, 2))
            array['permission'] = conve(sheet.cell_value(rown, 3))
            array['sync_type'] = conve(sheet.cell_value(rown, 4))
            array['icon'] = conve(sheet.cell_value(rown, 5))
            scene.append(array)
        sceneInfo['scene_setting'] = scene
        sceneInfo['room_id'] = 0
        sceneInfo['room_name'] = "场景系统"
        
        room.append(sceneInfo)
        roomInfo['room'] = room
        roomInfo['security'] = []
        return roomInfo

    elif sysType == "标准空调":
        roomInfo = {}
        room = []
        hvac = {}
        devInfo = {}
        device = []
        for rown in range(1, sheet.nrows):
            array = {}
            array['PB'] = conve(sheet.cell_value(rown, 0))
            array['ac_id'] = conve(sheet.cell_value(rown, 1))
            array['default_temper'] = conve(sheet.cell_value(rown, 2))
            array['device_id'] = conve(sheet.cell_value(rown, 3))
            array['device_name'] = conve(sheet.cell_value(rown, 4))
            array['device_type'] = conve(sheet.cell_value(rown, 5))
            array['index'] = conve(sheet.cell_value(rown, 6))
            array['max_temper'] = conve(sheet.cell_value(rown, 7))
            array['min_temper'] = conve(sheet.cell_value(rown, 8))
            array['name'] = conve(sheet.cell_value(rown, 9))
            array['sync_type'] = conve(sheet.cell_value(rown, 10))
            array['timeout_interval'] = conve(sheet.cell_value(rown, 11))
            array['sn'] = sheet.cell_value(rown, 12)
            device.append(array)
            devInfo['device'] = device
            devInfo['permission'] = 1
        hvac['conditioner'] = devInfo
        hvac['room_id'] = 0
        hvac['room_name'] = "空调系统"
        
        room.append(hvac)
        roomInfo['room'] = room
        roomInfo['security'] = []
        return roomInfo

    else:
        return

def conve(data):
    if type(data) == float:
        intData = int(data)
    else:
        intData = data
    return intData

def list2json(inPutFile, outPutFile, sysType):
    tempData = read_excel(inPutFile, sysType)
    if tempData == 1:
        return 1
    jsonData = json.dumps(tempData)
    jsonData = jsonData.encode('utf-8').decode('unicode_escape')
    fhandle = open(outPutFile, 'w', encoding='utf-8')                 # 写入文件
    fhandle.write(jsonData)
    fhandle.close()
    return 0
