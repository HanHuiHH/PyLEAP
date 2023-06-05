"""
    本代码封装保存模拟结果功能
    在保存结果后使用
"""

import os
import shutil
import time

import openpyxl as op
import pandas as pd
from SALib.analyze import sobol
import numpy as np
from matplotlib import ticker
import matplotlib.pyplot as plt
from scipy.stats import norm
import seaborn as sns

config = {
    "font.family": 'serif',
    "font.size": 16,  # 相当于小四字体
    "font.serif": ['SongNTR'],  # 宋体
    'axes.unicode_minus': False,  # 处理负号，即-号
    "figure.figsize": (20, 12)
}
plt.rcParams.update(config)
start_time = time.time()

# TODO:定义需要进行敏感性分析的参数的信息
problem = {
    'num_vars': 8,  # 需要测试敏感性的参数个数，可能和模型输入个数不同
    'names': ["GDP", "IndPorp", "PopGrow", "ElecSave", "EVProp", "PVGrow", "ImpElecGrow", "HydrogenRep"],
    # 需要测试敏感性的参数名称
    'dists': ['norm', 'norm', 'norm', 'norm', 'norm', 'norm', 'norm', 'norm'],
    'bounds': [[0.5, 0.2], [0.5, 0.2], [0.5, 0.2], [0.5, 0.2], [0.5, 0.2], [0.5, 0.2], [0.5, 0.2], [0.5, 0.2]]
    # 定义每个参数的上下界
}

# TODO:导入LEAP蒙特卡洛模拟的结果
files = os.listdir("./Results/Simulation")
# print(files)
file_name = "./Results/Simulation/" + files[-1]
print('Reading file:', file_name)
df_dict = pd.read_excel(file_name, sheet_name=None, header=0)
data_list = []
data_name = []
for df in df_dict:
    data = pd.read_excel(file_name, sheet_name=df, header=0)
    data_list.append(data)
    data_name.append(df)
# print(data_list)


'''
    Select the data
'''
df = data_list[1]  # Get CO2 emission data
df_peak_emission = pd.Series(df.max(axis=1)) / 100
df_peak_year = pd.Series(df.idxmax(axis=1))

df = data_list[2]  # Get net GHG emission data
df_2060_emission = pd.Series(df[2060]) / 100
neutral_year = []
for index, row in df.iterrows():
    if row.iloc[-1] > 0:
        neutral_year.append(2061)
    else:
        neutral_year.append(row[row < 0].index[0])
df_neutral_year = pd.DataFrame(neutral_year)
# print(neutral_time)

# plt.hist(df_peak_year)
# plt.show()

'''
    Start drawing pictures.
'''
fig = plt.figure()
gs = fig.add_gridspec(2, 2)
ax_list = []  # Initialize graph list
for i in range(2):  # number in range() must be integer
    ax_list.append(fig.add_subplot(gs[0, i]))  # Add a plot in upper part
    ax_list.append(fig.add_subplot(gs[1, i]))  # Add a plot in lower part

'''
    1. Draw the first picture, peak value and fitting normal curve
'''
# Draw hist
ax_list[0].hist(df_peak_emission, density=True, color='#607c8e', edgecolor='black', range=(365, 415), bins=20,
                hatch='//', rwidth=0.9, label='Histogram of result distribution')  # Draw histogram
# Draw normal distribution
mu = np.mean(df_peak_emission)  # Mean value
sigma = np.std(df_peak_emission)  # Standard Variation
bins = np.arange(365, 415, 1)  # Used to draw line. The smaller 3rd value, the more accurate the curve
norm_line = norm.pdf(bins, mu, sigma)  # Draw line
ax_list[0].plot(bins, norm_line, 'r--', label="Fitted normal distribution curve")  # Draw line
# Set other parameters
ax_list[0].set_xticks(range(365, 416, 10))  # Set X-axis points, range(first, last, interval)
ax_list[0].set_title('(a) Energy-related CO${_2}$ emission in carbon peak')
ax_list[0].set_xlabel('CO${_2}$ emission (Mt)')
ax_list[0].set_ylabel('Proportion')
ax_list[0].yaxis.set_major_formatter(ticker.PercentFormatter(xmax=1, decimals=2))

'''
    2. Draw the second picture, emission value in 2060 and fitting normal curve
'''
# Draw hist
ax_list[1].hist(df_2060_emission, density=True, color='#607c8e', edgecolor='black', range=(-100, 150), bins=20,
                hatch='//', rwidth=0.9)  # Draw histogram
# Draw normal distribution
mu = np.mean(df_2060_emission)  # Mean value
sigma = np.std(df_2060_emission)  # Standard Variation
bins = np.arange(-100, 150, 1)  # Used to draw line. The smaller 3rd value, the more accurate the curve
norm_line = norm.pdf(bins, mu, sigma)  # Draw line
ax_list[1].plot(bins, norm_line, 'r--')  # Draw line
# Set other parameters
ax_list[1].set_xticks(range(-100, 151, 50))  # Set X-axis points, range(first, last, interval)
ax_list[1].set_title('(b) Net GHG emission in 2060')
ax_list[1].set_xlabel('Net GHG emission (Mt CO${_2}$e)')
ax_list[1].set_ylabel('Proportion')
ax_list[1].yaxis.set_major_formatter(ticker.PercentFormatter(xmax=1, decimals=3))

'''
    3. Draw the third picture, peak year and fitting normal curve
'''
# Draw hist
ax_list[2].hist(df_peak_year, density=True, color='#607c8e', range=(2027.5, 2034.5), bins=7, edgecolor='black',
                hatch='//',
                rwidth=0.9)  # Draw histogram
# Draw normal distribution
mu = np.mean(df_peak_year)  # Mean value
sigma = np.std(df_peak_year)  # Standard Variation
bins = np.arange(2027, 2035, 0.1)  # Used to draw line. The smaller 3rd value, the more accurate the curve
norm_line = norm.pdf(bins, mu, sigma)  # Draw line
ax_list[2].plot(bins, norm_line, 'r--')  # Draw line
# Set other parameters
ax_list[2].set_title('(c) Carbon peak year')
ax_list[2].set_xlabel('Year')
ax_list[2].set_ylabel('Proportion')
# ax_list[2].set_ylim(0, 0.4)
ax_list[2].yaxis.set_major_formatter(ticker.PercentFormatter(xmax=1, decimals=0))

'''
    4. Draw the fourth picture, neutral year 
'''
# Draw hist
ax_list[3].hist(df_neutral_year, density=True, color='#607c8e', range=(2054.5, 2061.5), bins=7, edgecolor='black',
                hatch='//',
                rwidth=0.9)  # Draw histogram
ax_list[3].set_xticks(range(2055, 2062, 1), [str(i) for i in range(2055, 2061)] + [r'>2060'])
ax_list[3].set_title('(d) Carbon neutral year')
ax_list[3].set_xlabel('Year')
ax_list[3].set_ylabel('Proportion')
ax_list[3].set_ylim(0, 0.5)
ax_list[3].yaxis.set_major_formatter(ticker.PercentFormatter(xmax=1, decimals=0))

for ax in ax_list:
    ax.grid(axis='y', color='gray', linestyle='--')  # Set x-axis grid lines in all graph
    ax.set_axisbelow(True)  # Set x-axis grid lines on the bottom layer
fig.legend(loc='upper center', ncol=4, bbox_to_anchor=(0.5, 0.96))  # Add legend
fig.subplots_adjust(hspace=0.3)  # Adjust location
plt.savefig('./Results/Figures/Distribution'
            + time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime()) + '.jpg', dpi=600)
plt.show()

'''
    Draw curves
'''
fig = plt.figure(figsize=(12, 8))
ax = fig.add_subplot()
history_data = [359.23, 391.6881, 424.15, 439.88795, 455.63, 458.6277496, 461.63, 464.4368663, 467.24, 473.34,  480]
history_year = [i for i in range(2010, 2021)]
ax.plot(history_year, history_data, 'o-', label="History", color='green')
array_curve = df.drop(df.columns[0], axis=1).to_numpy() / 100
years = list(range(2020, 2020 + array_curve.shape[1]))
ax.plot(years, np.mean(array_curve, axis=0), label="Mean", color='black')
std_dev = np.std(array_curve, axis=0)  # Calculate std
# 绘制1, 2, 3个标准差的曲线并填充
colors = ['green', 'blue', 'red']
for i, color in enumerate(colors, 1):
    lower_bound = np.mean(array_curve, axis=0) - i * std_dev
    upper_bound = np.mean(array_curve, axis=0) + i * std_dev

    # 绘制边界虚线
    ax.plot(years, lower_bound, linestyle='--', color=color, alpha=0.5,
            label=f"Mean \u00B1 {i}\u03C3")
    ax.plot(years, upper_bound, linestyle='--', color=color, alpha=0.5)

# Draw all data
ax.fill_between(years,
                np.percentile(array_curve, 50 - 100 / 2., axis=0),
                # np.percentile函数获取2.5%和97.5%的大小上下界
                np.percentile(array_curve, 50 + 100 / 2., axis=0),
                alpha=0.5, color='gray',
                label=f"All results")

ax.set_xlabel("Year")
ax.set_ylabel("Net GHG emission (Mt CO${_2}$e)")
ax.legend(loc='upper right')._legend_box.align = "left"
ax.grid(axis='y', color='gray', linestyle='--')  # Set x-axis grid lines in all graph
ax.set_axisbelow(True)  # Set x-axis grid lines on the bottom layer

plt.savefig('./Results/Figures/Curve'
            + time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime()) + '.jpg', dpi=600)
plt.show()
