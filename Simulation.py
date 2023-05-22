
import time
import shutil
import win32com.client as client
from CalculateAndCheck import CALandCHECK

start_time = time.time()


def ReplaceName(ScenarioName, NameToReplace, i):
    NameLocation = ScenarioName.index(NameToReplace)
    ScenarioName = list(ScenarioName)
    ScenarioName[NameLocation + 3] = str(i)
    ScenarioName = (''.join(ScenarioName))
    return ScenarioName


def InitialAllRecords():
    EnergyDataRecord = []
    EmissionDataRecord = []
    NetEmissionRecord = []
    ElecConsDataRecord = []
    CleanPropRecord = []
    EnergyIntensityRecord = []
    CarbonIntensityRecord = []
    DataRecordNamess = ["EnergyDataRecord", "EmissionDataRecord", "NetEmissionRecord", "ElecConsDataRecord",
                        "CleanPropRecord", "EnergyIntensityRecord", "CarbonIntensityRecord"]  # 指标名称列表
    Records = [EnergyDataRecord, EmissionDataRecord, NetEmissionRecord, ElecConsDataRecord, CleanPropRecord,
               EnergyIntensityRecord, CarbonIntensityRecord]  # 指标对应记录列表
    AllRecords = dict(zip(DataRecordNamess, Records))  # 创建指标名称和指标对应范围的字典（AllRecords)
    return AllRecords


def test(leap, AllRanges, AllRecords, params, saved_origin_parameter):
    """LEAP模型循环运行"""
    """（1）输入参数取值的范围和准备好的情景，进行模拟"""
    NewAreaName = "Simulation0"  # 定义模拟Area的名称为Simulation
    # leap.Areas.Add(NewAreaName, OriginAreaName)  # 创建一个新的LEAP Area，如果报错，打开任务管理器把所有的LEAP进程都关掉重新试一试
    # leap.Areas(NewAreaName).Open()
    # leap.Scenarios("Comprehensive Scenario").Active = True  # 定位到所需要的情景
    # print("目前打开的Area：", leap.ActiveArea.Name)
    # print("目前打开的Scenario：", leap.ActiveScenario.Name)  # 如果scenario变成current account，在LEAP里面选择综合情景再直接模拟
    # ScenarioName = "Gr 0 Ip 0 Es 0 Ep 0 Pg 0 Ig 0 Hr 0 "  # 初始化情景组合名称
    # TotalSimNum = 1

    [IndIndexToChange, IndOriginExpression6th, SerIndexToChange, SerOrrginExpression6th, HyIndexToChange,
     HyOriginExpression5th] = saved_origin_parameter

    """（2）具体模拟开始"""
    """    1）不同人均GDP增长率"""
    GDPRanges = AllRanges.get("GDPRanges")
    GDPSample = params[0]
    IndexExpression = "Interp(2020, 3.68%, 2021, 8.03%, 2022, 6.5%, "  # 表达式的开头，不同数据可能需要修改
    for Grj in range(1, 9):
        IndexExpression = IndexExpression + str(2020 + 5 * Grj) + ", " \
                          + str('%f' % ((GDPRanges[1][Grj] - GDPRanges[2][Grj]) * GDPSample + GDPRanges[2][Grj])) + "%, "
    IndexExpression = IndexExpression[:-2] + ")"  # 表达式最后不能为“,”，需要替换为“)”
    leap.Branch(r"Key\地区生产总值\人均GDP增长率").Variable("Activity Level").Expression = IndexExpression  # 需要LEAP中赋值

    """    2）不同工业GDP占比"""
    IndPorpRanges = AllRanges.get("IndPorpRanges")
    IndSample = params[1]
    IndexExpression = "Growth(Interp("  # 表达式的开头，不同数据可能需要修改
    for Ipj in range(9):
        IndexExpression = IndexExpression + str(2020 + 5 * Ipj) + ", " \
                          + str('%f' % ((IndPorpRanges[1][Ipj] - IndPorpRanges[2][Ipj]) * IndSample + IndPorpRanges[2][Ipj])) + "%, "
    IndexExpression = IndexExpression[:-2] + "))"  # 表达式最后不能为“,”，需要替换为“)”
    leap.Branch(r"Key\地区生产总值\工业GDP比重").Variable("Activity Level").Expression = IndexExpression  # 需要LEAP中赋值

    """    3）不同人口增长率"""
    PopGrowRanges = AllRanges.get("PopGrowRanges")
    PopSample = params[2]
    IndexExpression = "Growth(Interp("  # 表达式的开头，不同数据可能需要修改
    for Poj in range(9):
        IndexExpression = IndexExpression + str(2020 + 5 * Poj) + ", " \
                          + str('%f' % ((PopGrowRanges[1][Poj] - PopGrowRanges[2][Poj]) * PopSample + PopGrowRanges[2][Poj])) + "%, "
    IndexExpression = IndexExpression[:-2] + "))"  # 表达式最后不能为“,”，需要替换为“)”
    leap.Branch(r"Key\人口\常住人口").Variable("Activity Level").Expression = IndexExpression  # 需要LEAP中赋值

    """    4）不同用电节能速度"""
    ElecSaveRange = AllRanges.get("ElecSaveRanges")
    EsSample = params[3]
    for Indi in range(len(IndIndexToChange)):
        IndIndex = r"Demand\Industry\ " + str(IndIndexToChange[Indi]) + "\Electricity"
        leap.Branch(IndIndex).Variable("Final Energy Intensity").Expression = IndOriginExpression6th[Indi][:-1] +\
            " + " + str('%f' % ((ElecSaveRange[1][-1] - ElecSaveRange[2][-1]) * EsSample + ElecSaveRange[2][-1])) + "%)"  # 需要LEAP中赋值
    for Seri in range(len(SerIndexToChange)):
        SerIndex = r"Demand\OtherService\ " + str(SerIndexToChange[Seri]) + "\Electricity"
        leap.Branch(SerIndex).Variable("Final Energy Intensity").Expression = SerOrrginExpression6th[Seri][:-1] +\
            " + " + str('%f' % ((ElecSaveRange[1][-1] - ElecSaveRange[2][-1]) * EsSample + ElecSaveRange[2][-1])) + "%)"  # 需要LEAP中赋值

    """    5）不同电动车替代率"""
    EVPropRanges = AllRanges.get("EVPropRanges")
    EpSample = params[4]
    IndexExpression = "Interp("  # 表达式的开头，不同数据可能需要修改
    for Epi in range(9):
        EVprop = (EVPropRanges[0][Epi] - EVPropRanges[2][Epi]) * EpSample * 2 + EVPropRanges[2][Epi]
        if EVprop > 100:
            EVprop = 100
        if EVprop < 0:
            EVprop = 0
        IndexExpression = IndexExpression + str(2020 + 5 * Epi) + ", " \
                          + str('%f' % EVprop) + ", "  # 注意这个不是增长率，后面没有%号
        #  因为电动车占比超过100%就没办法上升，所以要在这里限制
    IndexExpression = IndexExpression[:-2] + ")"  # 表达式最后不能为“,”，需要替换为“)”
    leap.Branch(r"Demand\Transport\轿车\电动汽车") \
        .Variable("Activity Level").Expression = IndexExpression  # 需要LEAP中赋值

    """    6）不同光伏发电增长速度"""
    PVGrowRanges = AllRanges.get("PVGrowRanges")
    PgSample = params[5]
    IndexExpression = "Growth(Interp("  # 表达式的开头，不同数据可能需要修改
    for Pgj in range(9):
        IndexExpression = IndexExpression + str(2020 + 5 * Pgj) + ", " \
                          + str('%f' % ((PVGrowRanges[1][Pgj] - PVGrowRanges[2][Pgj]) * PgSample + PVGrowRanges[2][Pgj])) + "%, "
    IndexExpression = IndexExpression[:-2] + "))"  # 表达式最后不能为“,”，需要替换为“)”
    leap.Branch(r"Transformation\Electricity Generation\Processes\Solar") \
        .Variable("Exogenous Capacity").Expression = IndexExpression  # 需要LEAP中赋值

    """    7）不同调入电力量"""
    ImpElecGrowRanges = AllRanges.get("ImpElecGrowRanges")
    IgSample = params[6]
    IndexExpression = "Interp("  # 表达式的开头，不同数据可能需要修改
    for Igj in range(9):
        IndexExpression = IndexExpression + str(2020 + 5 * Igj) + ", " \
                          + str('%f' % ((ImpElecGrowRanges[1][Igj] - ImpElecGrowRanges[2][Igj]) * IgSample + ImpElecGrowRanges[2][Igj])) + ", "  # 注意这个不是增长率，后面没有%号
    IndexExpression = IndexExpression[:-2] + ")"  # 表达式最后不能为“,”，需要替换为“)”
    leap.Branch(r"Key\其他指标\调入电力功率").Variable(
        "Activity Level").Expression = IndexExpression  # 需要LEAP中赋值

    """    8）不同氢能替代速度"""
    HydrogenRepRange = AllRanges.get("HydrogenRepRanges")
    HrSample = params[7]
    for Indi in range(len(HyIndexToChange)):
        Index = r"Demand\Industry\ " + str(HyIndexToChange[Indi]) + "\Hydrogen"
        leap.Branch(Index).Variable("Final Energy Intensity").Expression = HyOriginExpression5th[Indi] + " * " +\
            str('%f' % ((HydrogenRepRange[1][-1] - HydrogenRepRange[2][-1]) * HrSample + HydrogenRepRange[2][-1])) + "%"  # 需要LEAP中赋值
    """最终 计算步骤，包含在最后一层的for循环之中"""
    CALandCHECK(leap, NewAreaName, AllRecords)

    return AllRecords


if __name__ == '__main__':
    leap = client.DispatchEx('leap.LEAPApplication')  # 启动独立的进程
    leap.Visible = 0  # 0表示在后台以进程方式运行，不显示软件界面，1表示显示软件界面并可能需要操作
    MySimuCount = [3, 3, 3, 3, 3, 3, 3, 3]  # TODO:设置每个parameter的变化情况,不变就设置成1，变化的设置成变化维度
    # MySimuCount = [1, 1, 1, 1, 1, 1, 3]  # TODO:设置每个parameter的变化情况,不变就设置成1，变化的设置成变化维度

    ResultPath = "E:/1安徽碳中和/3模型计算/LEAP情景组合/模型记录/模型记录" + \
                 time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime()) + ".xlsx"  # TODO：修改结果保存路径
    shutil.copyfile(r"E:\1安徽碳中和\3模型计算\LEAP情景组合\空白参数记录.xlsx", ResultPath)  # TODO：空白参数记录文件地址
