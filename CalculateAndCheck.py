"""
    本代码封装了单次leap模型计算和查看结果的所有操作
    可以单独使用
"""

import time
import win32com.client as client


def InitialAllRecords():
    """
    初始化所有需要保存的参数
    :return: 参数保存列表
    """
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


def CalCleanProp(leap, Year):
    """
    用于计算清洁能源占比，计算将花费较长时间
    :param leap: leap对象
    :param Year: 查询年份
    :return: 清洁能源占比
    """
    CleanTypes = ["Biogas", "Wind", "Solar", "Hydro", "Nuclear", "Hydrogen", "Municipal Solid Waste", "Biomass",
                  "Import Clean Electricity"]  # 列出所有种类的清洁能源
    CleanConsumption = 0  # 用于累计清洁能源总量
    for CleanType in CleanTypes:
        CleanConsumption += leap.Branch("Demand").Variable(
            "Primary Requirements: Allocated to Demands"
        ).Value(Year, "Tonnes of Coal Equivalent", "fuel=" + CleanType) / (10 ** 8)
    CleanProp = CleanConsumption / leap.Branch("Demand").Variable(
        "Primary Requirements: Allocated to Demands"
    ).Value(Year, "Tonnes of Coal Equivalent") * (10 ** 8)

    return CleanProp


def CALandCHECK(leap, NewAreaName, AllRecords,
                check_energy_intensity_only=True,
                calculate_every_year=True,
                save_more_indicators=True,
                print_to_console=False):
    """
    计算LEAP，查看数据
    :param leap: LEAP模型
    :param NewAreaName: 新Area名称
    :param AllRecords: 参数保存列表
    :param check_energy_intensity_only: 是否只看能源强度下降率
    :param calculate_every_year: 是否计算每年的值
    :param save_more_indicators: 保存所有参数
    :param print_to_console: 在console中显示结果
    :return: Nothing
    """
    time1 = time.time()
    leap.Calculate()
    leap.Areas(NewAreaName).Open()  # 把LEAP软件窗口定位到Analysis界面，若定位在Results则可能会报错
    time2 = time.time()
    if print_to_console:
        print("计算耗时：", round(time2 - time1, 2), "s")

    EnergyConsumeFiveYear = []  # 创建空列表用于保存2020年及之后每5年能源消费量（万吨标准煤）
    CarbonEmissionFiveYear = []  # 创建空列表用于保存2020年及之后每5年碳排放量（kg）
    GDPFiveYear = []  # 创建空列表用于保存2020年及之后每5年GDP（元人民币）

    for i in range(9):  # 从2020至2060共8个5年，9个时间点
        EnergyConsumeFiveYear.append(
            leap.Branch("Demand").Variable(
                "Primary Requirements: Allocated to Demands"
            ).Value(2020 + 5 * i, "Tonnes of Coal Equivalent") / (10 ** 4)
        )  # 保存2020年及之后每5年能源消费量（万吨标准煤）
        CarbonEmissionFiveYear.append(
            leap.Branch("Demand").Variable(
                "One_Hundred Year GWP Direct and Indirect Allocated to Demands"
            ).Value(2020 + 5 * i) / (10 ** 7)
        )  # 保存2020年及之后每5年碳排放量（万吨）
        GDPFiveYear.append(
            leap.Branch(r"Key\地区生产总值\GDP").Variable("Activity Level").Value(2020 + 5 * i) / (10 ** 8)
        )  # 保存2020年及之后每5年GDP（亿元人民币）

    """
    查看模型计算单位GDP能耗下降率等关键数据，验证模型是否有误，如果不想看，请在函数输入中设置 
    """
    if print_to_console:
        # 计算并输出单位GDP能耗与下降率
        print(
            "单位GDP能耗下降率为：\n" +
            "第13个五年单位GDP能耗下降率为：" +
            "{:.2%}".format((EnergyConsumeFiveYear[0] / GDPFiveYear[0]) / 0.5162 - 1)  # 0.5162是2015年安徽省的单位GDP能耗
        )
        for i in range(8):
            print(
                "第" + str(i + 14) + "个五年单位GDP能耗下降率为：" +
                "{:.2%}".format((EnergyConsumeFiveYear[i + 1] / GDPFiveYear[i + 1])
                                / (EnergyConsumeFiveYear[i] / GDPFiveYear[i]) - 1)
            )
        print("")
    if not check_energy_intensity_only:
        # 计算并输出单位GDP碳排放与下降率
        print(
            "\n单位GDP碳排放下降率为：\n" +
            "第13个五年单位GDP碳排放下降率为：" +
            "{:.2%}".format((CarbonEmissionFiveYear[0] / GDPFiveYear[0]) / 1.2328 - 1)  # 1.2328是2015年安徽省单位GDP碳排放
        )
        for i in range(8):
            print(
                "第" + str(i + 14) + "个五年单位GDP碳排放下降率为：" +
                "{:.2%}".format((CarbonEmissionFiveYear[i + 1] / GDPFiveYear[i + 1])
                                / (CarbonEmissionFiveYear[i] / GDPFiveYear[i]) - 1)
            )

    time3 = time.time()
    if print_to_console:
        print("查看结果耗时：", round(time3 - time2, 2), "s")

    """开始保存模型数据，根据需要选择"""
    EnergyConsData = []  # 创建空列表用于保存能源消费量（万吨标准煤）
    EmissionData = []  # 创建空列表用于保存二氧化碳排放量（万吨）
    NetEmissionData = []  # 创建空列表用于保存净碳排放量（万吨二氧化碳当量）
    ElecConsData = []  # 创建空列表用于保存电力消费量（亿千瓦时）
    CleanPropData = []  # 创建空列表用于保存清洁能源比例（无单位，非%）
    EnergyIntensityData = []  # 创建空列表用于保存能源强度（吨标准煤/万元）
    CarbonIntensityData = []  # 创建空列表用于保存碳排放强度（吨二氧化碳/万元）

    """如果需要统计 每5年 的数据的话，请把 calculate_every_year 设成 False ，加快计算速度"""
    if calculate_every_year:
        for i in range(42):  # 从2019到2021年共42年
            EnergyConsData.append(
                leap.Branch("Demand").Variable(
                    "Primary Requirements: Allocated to Demands"
                ).Value(2019 + i, "Tonnes of Coal Equivalent") / (10 ** 4)
            )  # 保存每年能源消费量（万吨标准煤）
            EmissionData.append(
                leap.Branch("Demand").Variable(
                    "One_Hundred Year GWP Direct and Indirect Allocated to Demands"
                ).Value(2019 + i) / (10 ** 7)
            )  # 保存每年二氧化碳排放量（万吨）
            NetEmissionData.append(
                leap.Branch("安徽省碳达峰规划").Variable(
                    "One_Hundred Year GWP Direct At Point of Emissions"
                ).Value(2019 + i) / (10 ** 7)
            )  # 保存每年净碳排放量（万吨二氧化碳当量）
    else:  # 计算每5年的数据
        for i in range(9):  # 从2020至2060共8个5年，9个时间点
            EnergyConsData.append(
                leap.Branch("Demand").Variable(
                    "Primary Requirements: Allocated to Demands"
                ).Value(2020 + i * 5, "Tonnes of Coal Equivalent") / (10 ** 4)
            )  # 保存每年能源消费量（万吨标准煤）
            EmissionData.append(
                leap.Branch("Demand").Variable(
                    "One_Hundred Year GWP Direct and Indirect Allocated to Demands"
                ).Value(2020 + i * 5) / (10 ** 7)
            )  # 保存每年二氧化碳排放量（万吨）
            NetEmissionData.append(
                leap.Branch("安徽省碳达峰规划").Variable(
                    "One_Hundred Year GWP Direct At Point of Emissions"
                ).Value(2020 + i * 5) / (10 ** 7)
            )  # 保存每年净碳排放量（万吨二氧化碳当量）

    """如果需要统计 电力消费、清洁能源比例、五年下降率 的数据的话，请把 save_more_indicators 设为True ，否则不统计，加快计算速度"""
    if save_more_indicators:
        for i in range(9):  # 电力消费和清洁能源比重计算时间较久(主要是因为LEAP读取时间久），所以不计算每年的，从2020至2060共8个5年，9个时间点
            ElecConsData.append(
                leap.Branch("Demand").Variable(
                    "Energy Demand Final Units"
                ).Value(2020 + i * 5, "Kilowatt-Hour", "fuel=Electricity") / (10 ** 8)
            )  # 保存每年电力消费量（亿千瓦时）
            CleanPropData.append(
                CalCleanProp(leap, 2020 + i * 5)
            )

        for i in range(8):  # 计算五年下降率，从2025至2060共7个5年，8个时间点
            EnergyIntensityData.append(
                (EnergyConsumeFiveYear[i + 1] / GDPFiveYear[i + 1])
                / (EnergyConsumeFiveYear[i] / GDPFiveYear[i]) - 1
            )
            CarbonIntensityData.append(
                (CarbonEmissionFiveYear[i + 1] / GDPFiveYear[i + 1])
                / (CarbonEmissionFiveYear[i] / GDPFiveYear[i]) - 1
            )

    AllRecords["EnergyDataRecord"].append(EnergyConsData)
    AllRecords["EmissionDataRecord"].append(EmissionData)
    AllRecords["NetEmissionRecord"].append(NetEmissionData)
    AllRecords["ElecConsDataRecord"].append(ElecConsData)
    AllRecords["CleanPropRecord"].append(CleanPropData)
    AllRecords["EnergyIntensityRecord"].append(EnergyIntensityData)
    AllRecords["CarbonIntensityRecord"].append(CarbonIntensityData)

    time4 = time.time()
    if print_to_console:
        print("保存结果耗时：", round(time4 - time3, 2), "s" + "\n"
                                                        "小计一轮用时：", round(time4 - time1, 2), "s")


if __name__ == '__main__':
    leap = client.DispatchEx('leap.LEAPApplication')  # 启动独立的进程
    print("现在打开的Area为：", leap.ActiveArea())
    print("现在打开的Scenario为：", leap.ActiveScenario())
    CALandCHECK(leap, leap.ActiveArea(), InitialAllRecords(),
                print_to_console=True)

