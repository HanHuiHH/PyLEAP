"""
    本代码封装蒙特卡洛模拟
    可以运行
"""


import time
from SALib.sample import sobol as sobol_sample

from Simulation import simulation
import shutil
import win32com.client as client
from CalculateAndCheck import InitialAllRecords
from ImportFromExcel import GetValueRange
from SaveToExcel import save_all_records
from tqdm import *

"""
    模拟前准备
"""
start_time = time.time()
leap = client.DispatchEx('leap.LEAPApplication')  # 启动独立的进程

"""
    设置模拟参数
"""
OriginAreaName = "20231108安徽省碳排放总模型"
N_value = 2 ** 0  # 敏感性分析的N
leap.Visible = 0  # 0表示在后台以进程方式运行，不显示软件界面，1表示显示软件界面并可能需要操作
ImportPath = r"./ImportExcel/1 Parameters.xlsx"  # 参数读取路径
ResultPath = "./Results/Simulation/Simulation" + time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime()) + ".xlsx"  # 结果保存路径
shutil.copyfile(r"./ImportExcel/2 Simulation results (blank).xlsx", ResultPath)  # 空白参数记录文件地址

# 定义需要进行敏感性分析的参数的信息
problem = {
    'num_vars': 8,  # 需要测试敏感性的参数个数，可能和模型输入个数不同
    'names': ["GDP", "IndPorp", "PopGrow", "ElecSave", "EVProp", "PVGrow", "ImpElecGrow", "HydrogenRep"],
    # 需要测试敏感性的参数名称
    'dists': ['norm', 'norm', 'norm', 'norm', 'norm', 'norm', 'norm', 'norm'],
    'bounds': [[0.5, 0.2], [0.5, 0.2], [0.5, 0.2], [0.5, 0.2], [0.5, 0.2], [0.5, 0.2], [0.5, 0.2], [0.5, 0.2]]
    # 定义每个参数的均值和方差
}

if __name__ == '__main__':
    """
        开始进行敏感性分析
    """
    # sample
    param_values = sobol_sample.sample(problem, N_value)  # 后一个是敏感性分析N，sobol分析次数为N*(D+1)，D为分析维度，N推荐为500~1000

    # evaluate
    NewAreaName = "Simulation0"  # 定义模拟Area的名称为Simulation
    leap.Areas.Add(NewAreaName, OriginAreaName)  # 创建一个新的LEAP Area，如果报错，打开任务管理器把所有的LEAP进程都关掉重新试一试
    leap.Areas(NewAreaName).Open()
    leap.Scenarios("Comprehensive Scenario").Active = True  # 定位到所需要的情景
    print("目前打开的Area：", leap.ActiveArea.Name)
    print("目前打开的Scenario：", leap.ActiveScenario.Name)  # 如果scenario变成current account，在LEAP里面选择综合情景再直接模拟
    AllRecords = InitialAllRecords()

    """（1）提前进行 氢能替代 参数设置，否则会重复累加修改leap中参数"""
    HyIndexToChange = ["钢铁", "有色", "化工", "建材", "其他"]
    HyOriginExpression5th = []
    for Indi in range(len(HyIndexToChange)):  # 先保存原来的表达式，以免堆积修改
        Index = r"Demand\Industry\ " + str(HyIndexToChange[Indi]) + "\Hydrogen"
        HyOriginExpression5th.append(leap.Branch(Index).Variable("Final Energy Intensity").Expression)
    """（2）提前进行 用电节能 参数设置，否则会重复累加修改leap中参数"""
    IndIndexToChange = ["钢铁", "有色", "化工", "建材", "其他"]
    SerIndexToChange = ["Wholesale", "Public Service", "IT Service", "Real Estate", "Financial Service"]
    IndOriginExpression6th = []
    SerOrrginExpression6th = []
    for Indi in range(len(IndIndexToChange)):  # 先保存原来的表达式，以免堆积修改
        IndIndex = r"Demand\Industry\ " + str(IndIndexToChange[Indi]) + "\Electricity"
        IndOriginExpression6th.append(
            leap.Branch(IndIndex).Variable("Final Energy Intensity").Expression
        )
    for Seri in range(len(SerIndexToChange)):  # 先保存原来的表达式，以免堆积修改
        SerIndex = r"Demand\OtherService\ " + str(SerIndexToChange[Seri]) + "\Electricity"
        SerOrrginExpression6th.append(
            leap.Branch(SerIndex).Variable("Final Energy Intensity").Expression
        )
    saved_origin_parameter = [IndIndexToChange, IndOriginExpression6th, SerIndexToChange, SerOrrginExpression6th,
                              HyIndexToChange, HyOriginExpression5th]

    """
        主要模拟
    """
    for params in tqdm(param_values):
        AllRecords = simulation(
            leap=leap,
            AllRanges=GetValueRange(FilePath=ImportPath,
                                    SheetName="Sheet2"),
            AllRecords=AllRecords,
            params=params,
            saved_origin_parameter=saved_origin_parameter
        )
        if len(AllRecords["EnergyDataRecord"]) % (8 * 2 * (8 + 1)) == 0:  # 保存次数应该是N *（2D + 2），否则sobol分析会报错
            save_all_records(AllRecords, ResultPath)
            print("===================保存文件，已完成{}次模拟===================\n".format(len(AllRecords["EnergyDataRecord"])))

    save_all_records(AllRecords, ResultPath)  # 所有模拟完成后最后保存一次文件

    """⑥模拟结束，关闭并删除模拟使用的Area，否则重新模拟会报错"""
    # leap.Areas(NewAreaName).Save()
    leap.Areas(OriginAreaName).Open()
    leap.Areas.Delete(NewAreaName)
    leap.Visible = 1  # 显示LEAP界面，便于关闭
