# coding=utf-8
import os
import logging
import xlwt

logging.basicConfig(filename='mylog.txt', format="%(asctime)s : %(message)s",
                    level=logging.DEBUG)

# 网元映射关系(十六进;制)
NFTypeDict = {'00': 'reserved', '01': '5G-EIR', '02': 'AF', '03': 'AMF', '04': 'AUSF', '05': 'BSF',
              '06': 'CHF', '07': 'GMLC', '08': 'LMF', '09': 'N3IWF', '0A': 'NEF', '0B': 'NRF',
              '0C': 'NSSF', '0D': 'NWDAF', '0E': 'PCF', '0F': 'SEPP', '10': 'SMF', '11': 'SMSF',
              '12': 'UDM', '13': 'UDR', '14': 'UDSF', '15': 'UPF'}

# 大区映射关系
regionDict = {'01': '中部大区', '02': '西北大区', '03': '南部大区', '04': '西南大区',
              '05': '东部大区', '06': '北部大区', '07': '上海大区', '08': '北京大区'}

# 省份映射关系
provinceDict = {'80': '广东', '40': '广西', '20': '海南', '10': '湖南', '08': '福建'}

# 网络类型
NetwokTypeDict = {'00': '人网', '01': '物网'}

'''
人网AMF查询的SMF NF cache信息，不能超过如下范围
广东:     750000~7AFFFF       810000~86FFFF

物网AMF查询的SMF NF cache信息，不能超过如下范围：
广东:     750000~7AFFFF       810000~86FFFF
湖南:     870000~89FFFF       8D0000~8DFFFF
广西:     B30000~B5FFFF       BA0000~BAFFFF
福建:     5E0000~60FFFF       680000~68FFFF
海南:     C00000~C0FFFF       C20000~C2FFFF
'''
NFCacheRange2C = {'750000': '7AFFFF', '810000': '86FFFF'}

# 物网AMF查询的SMF NF cache信息，不能超过如下范围
NFCacheRange2B = {'750000': '7AFFFF', '810000': '86FFFF',
                  '870000': '89FFFF', '8D0000': '8DFFFF',
                  'B30000': 'B5FFFF', 'BA0000': 'BAFFFF',
                  '5E0000': '60FFFF', '680000': '68FFFF',
                  'C00000': 'C0FFFF', 'C20000': 'C2FFFF', }


# 读取文件夹下txt
def TXTFileList():
    filelist = []

    for root, dirs, files in os.walk(".", topdown=False):
        for name in files:
            str = os.path.join(root, name)
            if str.split('.')[-1] == 'txt' and 'MML' in str:
                filelist.append(str)
    return filelist


# 非法号段规则
def illegalRule(d1, d2):
    res = ''
    # d1格式：'start: 460017023100000,'     d2格式：'end: 460017023199999'
    # IMSI号，共15位，判断前十位不相等认为是非法号段。
    if len(d1.split(':')[1].rstrip(',').strip()) == 15 and len(d2.split(':')[1].strip()) == 15:
        if d1.split(':')[1].strip()[0:10] != d2.split(':')[1].strip()[0:10]:
            res = d1.strip() + ' ' + d2.strip()
    # PCF号段类型为MSISDN(共13位，判断前九位是否一致认为时非法号段)
    if len(d1.split(':')[1].rstrip(',').strip()) == 13 and len(d2.split(':')[1].strip()) == 13:
        if d1.split(':')[1].strip()[0:9] != d2.split(':')[1].strip()[0:9]:
            res = d1.strip() + ' ' + d2.strip()
    if len(d1.split(':')[1].rstrip(',').strip()) != len(d2.split(':')[1].strip()):
        res = d1.strip() + ' ' + d2.strip()

    # SMF TAC号段： {"start": "750000","end": "7affff"}
    if len(d1.split(':')[1].rstrip(',').strip()) == 6 and len(d2.split(':')[1].strip()) == 6:
        iStart = int(d1.split(':')[1].rstrip(',').strip(), 16)
        iEnd = int(d2.split(':')[1].rstrip(',').strip(), 16)
        flag = 0
        for key in NFCacheRange2B:
            iStartRange = int(key.upper(), 16)
            iEndRange = int(NFCacheRange2B[key].upper(), 16)
            if iStart >= iStartRange and iEnd <= iEndRange:
                flag = 1
                break
        if flag == 0:
            res = d1.strip() + ' ' + d2.strip()
    return res


# 解析文件，将数据放入到dict中，key表示所有非法号段，value表示fqdn
def txtAnalysis(filePath):
    dataListTmp = []
    dataList = []
    dataLists = []
    file = open(filePath, "rb")
    for line in file.readlines():
        dataLists.append(line.decode().strip('\n\t\r').replace('"', ''))

    # 将数据按照‘DSP NFCACHE: QUERYTYPE=NFID’分割，相同的放到一个list中
    NFID = ''
    for dts in dataLists:
        if 'nfInstanceId:' in dts:
            dataList.append(dataListTmp)
            dataListTmp = []
        dataListTmp.append(dts)
    dataList.append(dataListTmp)

    # 循环遍历各数组
    illegaldic = {}
    for dList in dataList:
        flag = 0
        illegalList = []
        segNum = 0
        if 'nfInstanceId' in dList[0]:
            nfInstanceId = str(dList[0].split(':')[1].rstrip(',').strip().split('-')[-1][0:6])

        else:
            continue
        if 'nfType' in dList[1]:
            nfType = dList[1].split(':')[1].rstrip(',').strip()
        for dt in range(1, len(dList)):
            # # 计算号段个数
            # if 'start' in dList[dt]:
            #     segNum += 1
            dt1 = dList[dt - 1]
            dt2 = dList[dt]
            if (dt + 1) < len(dList):
                dt3 = dList[dt + 1]
            if (dt + 15) < len(dList):
                dt4 = dList[dt + 15]
            # 规则1:取连续相邻的start和end为一组
            if 'start' in dt1 and 'end' in dt2:
                if len(illegalRule(dt1, dt2)):
                    illegalList.append(illegalRule(dt1, dt2))
            # 规则2:不连续数据，有些数据中间有一空行，需要取后一行比较
            if 'start' in dt1 and 'end' not in dt2 and 'end' in dt3:
                if len(illegalRule(dt1, dt3)):
                    illegalList.append(illegalRule(dt1, dt3))
            # 规则3:不连续数据，dt1=start，dt2 不等于end，dt3取后16行的数据
            if 'start' in dt1 and 'end' not in dt2 and 'end' in dt4:
                if len(illegalRule(dt1, dt4)):
                    illegalList.append(illegalRule(dt1, dt4))
            if 'MENAME:' in dList[dt] and flag == 0:
                # UNC / * MEID: 0   MENAME: GD_GD_GZ_AMF800_C_HW */  解析出网络类型；_C_
                netwokType = dList[dt].split('MENAME:')[1].split('*')[0].split('_')[-2]
                flag = 1

        keyflag = 0
        # key： [nfInstanceId=nfType]     value： [非法号段]
        for dickey in illegaldic:
            if nfInstanceId in dickey:
                # key = dickey.split('=')[0] + '=' + str(int(dickey.split('=')[1]) + segNum)
                # illegaldic[key] = illegaldic.pop(dickey)
                illegaldic[key].extend(illegalList)
                keyflag = 1
                break
        if keyflag == 0:
            key = str(nfInstanceId) + '=' + str(nfType) + '=' + str(netwokType)
            illegaldic[key] = illegalList
    file.close()
    for key in illegaldic:
        print(key, ':', illegaldic[key])
    return illegaldic


# 设置单元格式 入参type (1:表头第一列样式  2:某一单元格样式)
def SetFont(type):
    style = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    borders = xlwt.Borders()
    al = xlwt.Alignment()
    # 设置边框
    # borders.left = 1
    # borders.right = 1
    # borders.top = 1
    # borders.bottom = 1
    # borders.bottom_colour = 0x3A
    # style.borders = borders

    if type == 1:
        # 设置字体格式
        Font = xlwt.Font()
        Font.name = "Times New Roman"
        Font.bold = True  # 加粗
        style.font = Font

        al.horz = 0x02  # 设置水平居中
        al.vert = 0x01  # 设置垂直居中
        style.alignment = al
    elif type == 2:
        # 设置单元格背景色
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map['yellow']
        style.pattern = pattern
    elif type == 3:
        # 设置垂直居中
        al.vert = 0x01  # 设置垂直居中
        style.alignment = al
    elif type == 4:
        # 水平垂直居中
        al.horz = 0x02  # 设置水平居中
        al.vert = 0x01  # 设置垂直居中
        style.alignment = al
    return style


def MatchData(nfID):
    NFType = str(nfID[0:2])
    MDataDict = {}
    region = str(nfID[2:4])
    province = str(nfID[4:6])
    NetwokType = str(nfID[6:8])
    flag = 0
    for NFTyeKey in NFTypeDict:
        if NFType == NFTyeKey:
            MDataDict['NFType'] = NFTypeDict[NFTyeKey]
            flag = 1
            break
    if flag == 0:
        MDataDict['NFType'] = ''
    flag = 0
    for regionKey in regionDict:
        if region == regionKey:
            MDataDict['region'] = regionDict[regionKey]
            flag = 1
            break
    if flag == 0:
        MDataDict['region'] = ''
    flag = 0
    for provinceKey in provinceDict:
        if province == provinceKey:
            MDataDict['province'] = provinceDict[provinceKey]
            flag = 1
            break
    if flag == 0:
        MDataDict['province'] = ''
    flag = 0
    for NetwokTypeKey in NetwokType:
        if NetwokType == NetwokTypeKey:
            MDataDict['NetwokType'] = NetwokType[NetwokTypeKey]
            flag = 1
            break
    if flag == 0:
        MDataDict['NetwokType'] = ''
    return MDataDict


def XLSWrite(XLSPath, illegalData):
    # 写数据时，行计数器
    logging.info('xls write begin')
    # 实例化一个execl对象xls=工作薄
    xls = xlwt.Workbook()
    # 实例化一个工作表，名叫Sheet1
    sht1 = xls.add_sheet('非法号段信息')
    # 第一个参数是行，第二个参数是列，第三个参数是值,第四个参数是格式
    headFont = SetFont(1)
    bodyFont1 = SetFont(3)  # 垂直居中
    bodyFont2 = SetFont(4)  # 水平垂直居中

    sht1.write(0, 0, '网元类型', headFont)
    sht1.write(0, 1, '大区', headFont)
    sht1.write(0, 2, '省份', headFont)
    sht1.write(0, 3, '网络类型', headFont)
    sht1.write(0, 4, '冲突号段', headFont)

    shtNum1 = 1

    # 数据写入
    # sheet1
    rowBegin = 1
    for illData in illegalData:
        MdataDict = MatchData(illData.split('=')[0])
        NFType = illData.split('=')[1]
        networkType = illData.split('=')[2]
        nType = ''
        if networkType == 'C':
            nType = '人网'
        if networkType == 'B':
            nType = '物网'

        if len(illegalData[illData]):
            sht1.write_merge(rowBegin, rowBegin + len(illegalData[illData]) - 1, 0, 0, MdataDict['NFType'], bodyFont1)
            sht1.write_merge(rowBegin, rowBegin + len(illegalData[illData]) - 1, 1, 1, MdataDict['region'], bodyFont2)
            sht1.write_merge(rowBegin, rowBegin + len(illegalData[illData]) - 1, 2, 2, MdataDict['province'], bodyFont2)
            sht1.write_merge(rowBegin, rowBegin + len(illegalData[illData]) - 1, 3, 3, nType, bodyFont2)
            shtNum1 = rowBegin
            for ld in illegalData[illData]:
                sht1.write(shtNum1, 4, ld, bodyFont1)
                shtNum1 = shtNum1 + 1
            rowBegin += len(illegalData[illData])
        else:
            sht1.write(rowBegin, 0, MdataDict['NFType'], bodyFont1)
            sht1.write(rowBegin, 1, MdataDict['region'], bodyFont2)
            sht1.write(rowBegin, 2, MdataDict['province'], bodyFont2)
            sht1.write(rowBegin, 3, nType, bodyFont2)
            sht1.write(rowBegin, 4, '无', bodyFont1)
            rowBegin += 1

    xls.save(XLSPath)
    logging.info('xls write end')


def txtWrite(illegalData):
    file = open(os.getcwd() + '\\result.txt', 'w')
    for illData in illegalData:
        file.write(illData.split('=')[0] + '\n')
        file.write('    号段总数：' + illData.split('=')[1] + '\n')
        file.write('    非法号段：\n')
        for ld in illegalData[illData]:
            file.write('        ' + ld + '\n')
        file.write('\n')
    file.close()


def main():
    # 解析xls文件到list，用于后续数据处理数据源
    logging.info('welcome to txt world.')
    rNum = 1
    try:
        for f in TXTFileList():
            # 文件分析，提取所需数据
            illegaldic = txtAnalysis(f)
            # 数据输出，写入到txt
            # txtWrite(illegaldic)

            # 数据输出写入xls
            XLSWrite(os.getcwd() + '\\illegalSegment' + str(rNum) + '.xls', illegaldic)
            rNum += 1


    except Exception as err:
        logging.error(err)

    logging.info("end txt world")


if __name__ == '__main__':
    # # 750000~7AFFFF
    # print(int('750001', 16))
    # print(int('7AFFFF', 16))
    main()
