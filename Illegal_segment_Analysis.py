# coding=utf-8
import os
import logging
import xlwt

logging.basicConfig(filename='mylog.txt', format="%(asctime)s : %(message)s",
                    level=logging.DEBUG)


# 读取文件夹下txt
def TXTFileList():
    filelist = []

    for root, dirs, files in os.walk(".", topdown=False):
        for name in files:
            str = os.path.join(root, name)
            if str.split('.')[-1] == 'txt' and 'MML' in str.upper():
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
        if d1.split(':')[1].strip()[0:8] != d2.split(':')[1].strip()[0:8]:
            res = d1.strip() + ' ' + d2.strip()
    if len(d1.split(':')[1].rstrip(',').strip()) != len(d2.split(':')[1].strip()):
        res = d1.strip() + ' ' + d2.strip()
    return res


# 解析文件，将数据放入到dict中，key表示所有非法号段，value表示fqdn
def txtAnalysis(filePath):
    logging.info('begin to analysis mml:%s', filePath)
    dataListTmp = []
    dataList = []
    dataLists = []
    file = open(filePath, "rb")
    for line in file.readlines():
        dataLists.append(line.decode().strip('\n\t\r').replace('"', ''))

    # 将数据按照‘DSP NFCACHE: QUERYTYPE=NFID’分割，相同的放到一个list中
    NFID = ''
    for dts in dataLists:
        if 'fqdn' in dts:
            if len(dataListTmp):
                dataList.append(dataListTmp)
                dataListTmp = []
        dataListTmp.append(dts)
    dataList.append(dataListTmp)
    # for i in dataList:
    #     print(i)

    # 循环遍历各数组
    illegaldic = {}
    for dList in dataList:
        illegalList = []
        segNum = 0
        fqdn = ''
        if 'fqdn' in dList[0]:
            fqdn = dList[0].rstrip(',').split(':')[1].strip(' ')
        else:
            continue
        for dt in range(1, len(dList)):
            # 计算号段个数
            if 'start' in dList[dt]:
                segNum += 1
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

        keyflag = 0
        # key： [fqdn:号段个数]     value： [非法号段]
        for dickey in illegaldic:
            if fqdn in dickey:
                key = dickey.split('=')[0] + '=' + str(int(dickey.split('=')[1]) + segNum)
                illegaldic[key] = illegaldic.pop(dickey)
                illegaldic[key].extend(illegalList)
                keyflag = 1
                break
        if keyflag == 0:
            key = fqdn + '=' + str(segNum)
            illegaldic[key] = illegalList
    file.close()
    logging.info('end analysis mml:%s', filePath)
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


def XLSWrite(XLSPath, illegalData):
    # 写数据时，行计数器
    logging.info('xls write begin：%s', XLSPath)
    # 实例化一个execl对象xls=工作薄
    xls = xlwt.Workbook()
    # 实例化一个工作表，名叫Sheet1
    sht1 = xls.add_sheet('非法号段信息')
    # 第一个参数是行，第二个参数是列，第三个参数是值,第四个参数是格式
    headFont = SetFont(1)
    bodyFont1 = SetFont(3)  # 垂直居中
    bodyFont2 = SetFont(4)  # 水平垂直居中

    sht1.write(0, 0, 'fqdn', headFont)
    sht1.write(0, 1, '号段总数', headFont)
    sht1.write(0, 2, '非法号段', headFont)

    for i in illegalData:
        print(i, ' ', illegalData[i])

    shtNum1 = 1

    # 数据写入
    # sheet1
    rowBegin = 1
    for illData in illegalData:
        fqdn = illData.split('=')[0]
        segNum = illData.split('=')[1]

        if len(illegalData[illData]):
            sht1.write_merge(rowBegin, rowBegin + len(illegalData[illData]) - 1, 0, 0, fqdn, bodyFont1)
            sht1.write_merge(rowBegin, rowBegin + len(illegalData[illData]) - 1, 1, 1, segNum, bodyFont2)
            shtNum1 = rowBegin
            for ld in illegalData[illData]:
                sht1.write(shtNum1, 2, ld, bodyFont2)
                shtNum1 = shtNum1 + 1
            rowBegin += len(illegalData[illData])
        else:
            sht1.write(rowBegin, 0, fqdn, bodyFont1)
            sht1.write(rowBegin, 1, segNum, bodyFont2)
            sht1.write(rowBegin, 2, 'NULL', bodyFont2)
            rowBegin += 1

    xls.save(XLSPath)
    logging.info('xls write end:%s', XLSPath)


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
    logging.info('welcome to illegal_segment_analysis world.')
    try:
        for f in TXTFileList():
            # 文件分析，提取所需数据
            illegaldic = txtAnalysis(f)
            # 数据输出，写入到txt
            # txtWrite(illegaldic)

            # 数据输出写入xls
            XLSWrite(os.getcwd() + '\\illegalSegment.xls', illegaldic)


    except Exception as err:
        logging.error(err)

    logging.info("end illegal_segment_analysis world")


if __name__ == '__main__':
    main()
