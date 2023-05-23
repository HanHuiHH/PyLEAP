"""
    本代码封装LEAP模型参数修改功能
    无法单独使用
"""

import time
import shutil
import win32com.client as client
from CalculateAndCheck import CALandCHECK

start_time = time.time()


def simulation(leap, AllRanges, AllRecords, params, saved_origin_parameter):
    """
    根据读取的参数对LEAP模型中的参数赋值，
    :param leap: LEAP对象
    :param AllRanges: 所有参数的取值范围
    :param AllRecords: 所有结果参数的记录
    :param params: 蒙特卡洛的参数
    :param saved_origin_parameter: 用电节能 和 氢能替代 在LEAP中较难直接更改，需要提前保存好数值
    :return:
    """
    """（1）输入参数取值的范围和准备好的情景，进行模拟"""
    NewAreaName = "Simulation0"  # 定义模拟Area的名称为Simulation

    [IndIndexToChange, IndOriginExpression6th, SerIndexToChange, SerOrrginExpression6th, HyIndexToChange,
     HyOriginExpression5th] = saved_origin_parameter

    """（2）具体模拟开始"""
    """    1）不同人均GDP增长率"""
    GDPRanges = AllRanges.get("GDPRanges")  # 读取自己设置的该参数取值范围
    GDPSample = params[0]  # 读取Monte Carlo参数的取值
    IndexExpression = "Interp(2020, 3.68%, 2021, 8.03%, 2022, 6.5%, "  # 表达式的开头，不同数据可能需要修改
    for Grj in range(1, 9):
        IndexExpression = IndexExpression + str(2020 + 5 * Grj) + ", " \
                          + str('%f' % ((GDPRanges[1][Grj] - GDPRanges[2][Grj]) * GDPSample + GDPRanges[2][Grj])) + "%, "
    IndexExpression = IndexExpression[:-2] + ")"  # 表达式最后不能为“,”，需要替换为“)”
    leap.Branch(r"Key\地区生产总值\人均GDP增长率").Variable("Activity Level").Expression = IndexExpression  # 在LEAP中赋值

    """    2）不同工业GDP占比"""
    IndPorpRanges = AllRanges.get("IndPorpRanges")  # 读取自己设置的该参数取值范围
    IndSample = params[1]  # 读取Monte Carlo参数的取值
    IndexExpression = "Growth(Interp("  # 表达式的开头，不同数据可能需要修改
    for Ipj in range(9):
        IndexExpression = IndexExpression + str(2020 + 5 * Ipj) + ", " \
                          + str('%f' % ((IndPorpRanges[1][Ipj] - IndPorpRanges[2][Ipj]) * IndSample + IndPorpRanges[2][Ipj])) + "%, "
    IndexExpression = IndexExpression[:-2] + "))"  # 表达式最后不能为“,”，需要替换为“)”
    leap.Branch(r"Key\地区生产总值\工业GDP比重").Variable("Activity Level").Expression = IndexExpression  # 在LEAP中赋值

    """    3）不同人口增长率"""
    PopGrowRanges = AllRanges.get("PopGrowRanges")  # 读取自己设置的该参数取值范围
    PopSample = params[2]  # 读取Monte Carlo参数的取值
    IndexExpression = "Growth(Interp("  # 表达式的开头，不同数据可能需要修改
    for Poj in range(9):
        IndexExpression = IndexExpression + str(2020 + 5 * Poj) + ", " \
                          + str('%f' % ((PopGrowRanges[1][Poj] - PopGrowRanges[2][Poj]) * PopSample + PopGrowRanges[2][Poj])) + "%, "
    IndexExpression = IndexExpression[:-2] + "))"  # 表达式最后不能为“,”，需要替换为“)”
    leap.Branch(r"Key\人口\常住人口").Variable("Activity Level").Expression = IndexExpression  # 在LEAP中赋值

    """    4）不同用电节能速度"""
    ElecSaveRange = AllRanges.get("ElecSaveRanges")  # 读取自己设置的该参数取值范围
    EsSample = params[3]  # 读取Monte Carlo参数的取值
    for Indi in range(len(IndIndexToChange)):
        IndIndex = r"Demand\Industry\ " + str(IndIndexToChange[Indi]) + "\Electricity"
        leap.Branch(IndIndex).Variable("Final Energy Intensity").Expression = IndOriginExpression6th[Indi][:-1] +\
            " + " + str('%f' % ((ElecSaveRange[1][-1] - ElecSaveRange[2][-1]) * EsSample + ElecSaveRange[2][-1])) + "%)"  # 在LEAP中赋值
    for Seri in range(len(SerIndexToChange)):
        SerIndex = r"Demand\OtherService\ " + str(SerIndexToChange[Seri]) + "\Electricity"
        leap.Branch(SerIndex).Variable("Final Energy Intensity").Expression = SerOrrginExpression6th[Seri][:-1] +\
            " + " + str('%f' % ((ElecSaveRange[1][-1] - ElecSaveRange[2][-1]) * EsSample + ElecSaveRange[2][-1])) + "%)"  # 在LEAP中赋值

    """    5）不同电动车替代率"""
    EVPropRanges = AllRanges.get("EVPropRanges")  # 读取自己设置的该参数取值范围
    EpSample = params[4]  # 读取Monte Carlo参数的取值
    IndexExpression = "Interp("  # 表达式的开头，不同数据可能需要修改
    for Epi in range(9):
        EVprop = (EVPropRanges[0][Epi] - EVPropRanges[2][Epi]) * EpSample * 2 + EVPropRanges[2][Epi]
        if EVprop > 100:  # 电动车占有率不可能超过100%和小于0%
            EVprop = 100
        if EVprop < 0:
            EVprop = 0
        IndexExpression = IndexExpression + str(2020 + 5 * Epi) + ", " \
                          + str('%f' % EVprop) + ", "  # 注意这个不是增长率，后面没有%号
        #  因为电动车占比超过100%就没办法上升，所以要在这里限制
    IndexExpression = IndexExpression[:-2] + ")"  # 表达式最后不能为“,”，需要替换为“)”
    leap.Branch(r"Demand\Transport\轿车\电动汽车") \
        .Variable("Activity Level").Expression = IndexExpression  # 在LEAP中赋值

    """    6）不同光伏发电增长速度"""
    PVGrowRanges = AllRanges.get("PVGrowRanges")  # 读取自己设置的该参数取值范围
    PgSample = params[5]  # 读取Monte Carlo参数的取值
    IndexExpression = "Growth(Interp("  # 表达式的开头，不同数据可能需要修改
    for Pgj in range(9):
        IndexExpression = IndexExpression + str(2020 + 5 * Pgj) + ", " \
                          + str('%f' % ((PVGrowRanges[1][Pgj] - PVGrowRanges[2][Pgj]) * PgSample + PVGrowRanges[2][Pgj])) + "%, "
    IndexExpression = IndexExpression[:-2] + "))"  # 表达式最后不能为“,”，需要替换为“)”
    leap.Branch(r"Transformation\Electricity Generation\Processes\Solar") \
        .Variable("Exogenous Capacity").Expression = IndexExpression  # 在LEAP中赋值

    """    7）不同调入电力量"""
    ImpElecGrowRanges = AllRanges.get("ImpElecGrowRanges")  # 读取自己设置的该参数取值范围
    IgSample = params[6]  # 读取Monte Carlo参数的取值
    IndexExpression = "Interp("  # 表达式的开头，不同数据可能需要修改
    for Igj in range(9):
        IndexExpression = IndexExpression + str(2020 + 5 * Igj) + ", " \
                          + str('%f' % ((ImpElecGrowRanges[1][Igj] - ImpElecGrowRanges[2][Igj]) * IgSample + ImpElecGrowRanges[2][Igj])) + ", "  # 注意这个不是增长率，后面没有%号
    IndexExpression = IndexExpression[:-2] + ")"  # 表达式最后不能为“,”，需要替换为“)”
    leap.Branch(r"Key\其他指标\调入电力功率").Variable(
        "Activity Level").Expression = IndexExpression  # 在LEAP中赋值

    """    8）不同氢能替代速度"""
    HydrogenRepRange = AllRanges.get("HydrogenRepRanges")  # 读取自己设置的该参数取值范围
    HrSample = params[7]  # 读取Monte Carlo参数的取值
    for Indi in range(len(HyIndexToChange)):
        Index = r"Demand\Industry\ " + str(HyIndexToChange[Indi]) + "\Hydrogen"
        leap.Branch(Index).Variable("Final Energy Intensity").Expression = HyOriginExpression5th[Indi] + " * " +\
            str('%f' % ((HydrogenRepRange[1][-1] - HydrogenRepRange[2][-1]) * HrSample + HydrogenRepRange[2][-1])) + "%"  # 在LEAP中赋值

    """（3）最终 计算步骤"""
    CALandCHECK(leap, NewAreaName, AllRecords)

    return AllRecords

