"""
    本代码封装从excel中读取变量取值范围的功能
"""

import time
import openpyxl as op


def GetValueRange(FilePath, SheetName, SimuCount):
    """输入指标文件路径、表格名称、各参数模拟的次数，输出指标名称和对应范围列表的字典"""
    wb = op.load_workbook(filename=FilePath, data_only=True)
    ws = wb[SheetName]

    GDPRanges = []  # Range为该指标多个Range的集合（3×9），该指标为人均GDP增长率
    for i in range(SimuCount[0]):  # range(3)指这个指标分了3个层次，如果细分更多层次需要修改
        row = 2 + i  # 指标从哪行开始就填入哪个数字
        GDPRanges.append([])  # 每个层次创建独立的列表进行储存
        for j in range(9):
            GDPRanges[i].append(ws[chr(ord("E") + j) + str(row)].value)  # 从2020年开始填入列表值

    IndPorpRanges = []  # Range为该指标多个Range的集合（3×9），该指标为工业占比
    for i in range(SimuCount[1]):  # range(3)指这个指标分了3个层次，如果细分更多层次需要修改
        row = 5 + i  # 指标从哪行开始就填入哪个数字
        IndPorpRanges.append([])  # 每个层次创建独立的列表进行储存
        for j in range(9):
            IndPorpRanges[i].append(ws[chr(ord("E") + j) + str(row)].value)  # 从2020年开始填入列表值

    PopGrowRanges = []  # Range为该指标多个Range的集合（3×9），该指标为人口增长率
    for i in range(SimuCount[2]):  # range(3)指这个指标分了3个层次，如果细分更多层次需要修改
        row = 8 + i  # 指标从哪行开始就填入哪个数字
        PopGrowRanges.append([])  # 每个层次创建独立的列表进行储存
        for j in range(9):
            PopGrowRanges[i].append(ws[chr(ord("E") + j) + str(row)].value)  # 从2020年开始填入列表值

    ElecSaveRanges = []  # Range为该指标多个Range的集合（3×9），该指标为用电节能速度
    for i in range(SimuCount[3]):  # range(3)指这个指标分了3个层次，如果细分更多层次需要修改
        row = 11 + i  # 指标从哪行开始就填入哪个数字
        ElecSaveRanges.append([])  # 每个层次创建独立的列表进行储存
        for j in range(9):
            ElecSaveRanges[i].append(ws[chr(ord("E") + j) + str(row)].value)  # 从2020年开始填入列表值

    EVPropRanges = []  # Range为该指标多个Range的集合（3×9），该指标为电动车占比
    for i in range(SimuCount[4]):  # range(3)指这个指标分了3个层次，如果细分更多层次需要修改
        row = 14 + i  # 指标从哪行开始就填入哪个数字
        EVPropRanges.append([])  # 每个层次创建独立的列表进行储存
        for j in range(9):
            EVPropRanges[i].append(ws[chr(ord("E") + j) + str(row)].value)  # 从2020年开始填入列表值

    PVGrowRanges = []  # Range为该指标多个Range的集合（3×9），该指标为光伏发电增长率
    for i in range(SimuCount[5]):  # range(3)指这个指标分了3个层次，如果细分更多层次需要修改
        row = 17 + i  # 指标从哪行开始就填入哪个数字
        PVGrowRanges.append([])  # 每个层次创建独立的列表进行储存
        for j in range(9):
            PVGrowRanges[i].append(ws[chr(ord("E") + j) + str(row)].value)  # 从2020年开始填入列表值

    ImpElecGrowRanges = []  # Range为该指标多个Range的集合（3×9），该指标为调入电力量
    for i in range(SimuCount[6]):  # range(3)指这个指标分了3个层次，如果细分更多层次需要修改
        row = 20 + i  # 指标从哪行开始就填入哪个数字
        ImpElecGrowRanges.append([])  # 每个层次创建独立的列表进行储存
        for j in range(9):
            ImpElecGrowRanges[i].append(ws[chr(ord("E") + j) + str(row)].value)  # 从2020年开始填入列表值

    HydrogenRepRanges = []  # Range为该指标多个Range的集合（3×9），该指标为氢能替代速率
    for i in range(SimuCount[7]):  # range(3)指这个指标分了3个层次，如果细分更多层次需要修改
        row = 23 + i  # 指标从哪行开始就填入哪个数字
        HydrogenRepRanges.append([])  # 每个层次创建独立的列表进行储存
        for j in range(9):
            HydrogenRepRanges[i].append(ws[chr(ord("E") + j) + str(row)].value)  # 从2020年开始填入列表值

    RangeNames = ["GDPRanges", "IndPorpRanges", "PopGrowRanges", "ElecSaveRanges",
                  "EVPropRanges", "PVGrowRanges", "ImpElecGrowRanges", "HydrogenRepRanges"]  # 指标名称列表
    Ranges = [GDPRanges, IndPorpRanges, PopGrowRanges, ElecSaveRanges,
              EVPropRanges, PVGrowRanges, ImpElecGrowRanges, HydrogenRepRanges]  # 指标对应范围列表
    AllRanges = dict(zip(RangeNames, Ranges))  # 创建指标名称和指标对应范围的字典（AllRanges)
    # print(AllRanges)

    return AllRanges



