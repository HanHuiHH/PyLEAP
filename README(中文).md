# PyLEAP - LEAP API 的 Python 编程
欢迎来到 PyLEAP - LEAP API 的 Python 编程。这个项目是为了进一步利用 LEAP 软件和 Python 进行研究，包括自动分析和多目标优化。

在示例代码中，我们使用 [SALib](https://github.com/SALib/SALib) 进行蒙特卡罗模拟和 Sobol 全局敏感性分析。

目前代码中的注释都是中文的。

## LEAP 及其 API

### LEAP 软件
LEAP（低排放分析平台）是斯德哥尔摩环境研究所（SEI）开发的广泛使用的能源政策、气候变化减缓和空气污染防治规划软件工具。在 [website of LEAP](https://leap.sei.org) 上可以找到更多信息。你还可以通过他们的 [Open Youtube Courses](https://www.youtube.com/watch?v=y4b2KCIxOJU&list=PLX-Kjcc7K01EOTxozEEBu2aerJmZ6ZfRT&ab_channel=LEAPPlatform) 快速入门。

### LEAP 的 API
LEAP 可以作为标准的 "COM 自动化服务器"，这意味着其他 Windows 程序可以直接控制 LEAP：更改数据值、计算结果并将其导出到 Excel 或其他应用程序中。API 甚至提供了用于检查或更改 LEAP 数据结构的功能。这种编程 LEAP 的能力非常强大。

例如，你可以编写一个简短的脚本，可以多次运行 LEAP 计算，每次使用不同的输入假设。然后 LEAP 的结果可以输出到 Excel 中，或者在脚本中进行处理，并用于计算后续 LEAP 计算的修订假设。通过这种方式，LEAP 的基本账务计算可以与更复杂的算法（如目标寻求或优化算法）相结合。

要获取更多信息，请打开 LEAP 软件中的内容，并查看高级主题/自动化 LEAP（API）。

## 性能
### 迭代
在一个包含 i7 8700 处理器、16GB 内存和 Python 3.9.16 的系统上，每次迭代大约需要5秒。对于大样本量来说，这相对较慢。为确保代码的功能性，建议你最初使用几个小样本进行测试。由于 LEAP 软件的限制，PyLEAP 中不可用多进程运行方法。我们正在尝试实现这个功能。

### 结果处理
迭代结束后，结果将被保存。如果需要一个大样本量（运行时间超过10小时），请尝试将其分成几个部分，并合并这些部分的结果。结果处理的时间不会太长。

##代码结构

### 计算和检查（CalculateAndCheck.py）
这段代码调用 LEAP 软件中的计算函数，并在控制台中打印出几个关键值（例如5年内的能源强度降低率）。你可以使用这段代码来检查 LEAP 区域中的关键值。

### 从 Excel 导入（ImportFromExcel.py）
这段代码用于从ImportExcel文件夹中导入Excel文件中的值。模拟需要定义关键参数的分布，以生成样本。均值和标准差是正态分布的关键参数，可以在这个Excel文件中进行编辑。

### 蒙特卡罗模拟 (MonteCarlo-Simulation.py)
这段代码将前两段代码集成到 SALib 库的蒙特卡罗模拟中。结果保存在 Result 文件夹中。

### Sobol 全局敏感性分析 (Sobol'GSA.py)
这段代码使用蒙特卡罗的结果进行 Sobol 全局敏感性分析，并绘制一些图形。