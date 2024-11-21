# 订单信息分析工具

## 简介

这是一个用于分析订单信息的Python脚本，它读取Excel文件中的订单数据，并计算每个月新增、延续和超期终止的用户数量。该脚本使用Pandas库来处理数据，并最终将结果输出为一个新的Excel文件。

## 功能

- 读取Excel文件中的订单信息。
- 将日期列转换为datetime类型，并提取月份和年份。
- 遍历指定日期范围内的每个月，计算每个月的新增终端用户、延续终端用户和超期终止终端用户的数量。
- 将统计结果保存到字典中，并最终输出为Excel文件。

## 使用方法

1. 确保Python环境已安装，并且已安装Pandas库。
2. 将脚本保存为`.py`文件，例如`order_analysis.py`。
3. 确保Excel文件`订单信息20241115.xlsx`位于脚本同一目录下，或者修改脚本中的文件路径以匹配Excel文件的实际位置。
