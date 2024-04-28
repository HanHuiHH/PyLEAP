"""
    如果其他代码出错，请运行这个程序，以初始化LEAP软件
"""

import time
import win32com.client as client

start_time = time.time()

leap = client.DispatchEx('leap.LEAPApplication')  # 启动独立的进程
leap.Visible = 1  # 0表示在后台以进程方式运行，不显示软件界面，1表示显示软件界面并可能需要操作

leap.Areas("Freedonia").Open()
leap.Areas.Delete("Simulation0")

leap.Areas("20231108安徽省碳排放总模型").Open()
print("目前打开的Area：", leap.ActiveArea.Name)

leap.Visible = 1  # 0表示在后台以进程方式运行，不显示软件界面，1表示显示软件界面并可能需要操作

print("可以继续运行模拟代码")
# 打开LEAP并定位到需要的情景再运行！！！
