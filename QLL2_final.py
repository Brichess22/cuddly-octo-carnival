#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd

# 读取 Excel 文件
order_info = pd.read_excel('订单信息20241115.xlsx')

# 转换日期列为 datetime 类型
order_info['下单时间'] = pd.to_datetime(order_info['下单时间'])
order_info['授权时间'] = pd.to_datetime(order_info['授权时间'])

# 提取月份和年份
table_months = order_info['下单时间'].dt.month
table_years = order_info['下单时间'].dt.year
nopriday = order_info['下单时间'].dt.day
priday = order_info['授权时间'].dt.day

# 获取分析范围的月份列表
months = pd.date_range(start='2022-04-01', end='2024-05-01', freq='MS')

# 创建一个字典来存储每个月的用户数量
monthly_user_counts = {}

# 循环每个月进行统计
for month in months:
    # 获取上个月的开始和结束日期
    previous_month_start = month - pd.DateOffset(months=1)
    previous_month_end = month - pd.DateOffset(days=1)
    previous_month_end = previous_month_end.replace(hour=23, minute=59, second=59)

    # 1. 新增终端用户：上月开通的用户
    new_users = order_info[
        (order_info['下单时间'] >= previous_month_start) &
        (order_info['下单时间'] <= previous_month_end) &
        # 此处的0为文本类型，数字类型需要用单引号括起来
        (order_info['类型'] == 0)
        ]
    # 2. 延续终端用户：开通早于上月，且截至日期晚于上月
    continued_users = order_info[
        (order_info['下单时间'] < previous_month_start) &
        (order_info['授权时间'] > previous_month_end) &
     #   ~(order_info['userID'].isin(new_users['userID'])) &
        (order_info['类型'] == 0)
        ]
    # 3. 超期终止终端用户：开通早于上月，且截至日期在上月
    overage_users = order_info[
        (order_info['下单时间'] < previous_month_start) &
        (order_info['授权时间'] >= previous_month_start) &
        (order_info['授权时间'] <= previous_month_end) &
        (priday >= nopriday) &
  #     ~(order_info['userID'].isin(new_users['userID'])) &
        (order_info['类型'] == 0)
        ]
    
    
# userid字符串转换+去重
    users = pd.concat([new_users, continued_users, overage_users])
    users['userID'] = users['userID'].str.lower()
    unique_userid_count = users['userID'].nunique()
    #users_unique = users.drop_duplicates(subset=['userID])'
    
    # 计算三类用户数量的总和
    #total_users = len(overage_users)
    #total_users = len(new_users) + len(continued_users) + len(overage_users)

    # 将每个月的用户数量保存到字典中
    monthly_user_counts[month.strftime('%Y-%m')] = unique_userid_count

# 打印每个月的统计结果
for month, count in monthly_user_counts.items():
    print(f"{month}: {count}")
# 将统计结果转换为 DataFrame
results_df = pd.DataFrame(list(monthly_user_counts.items()), columns=['月份', '用户数量'])
# 将结果输出为 Excel 文件
output_file = 'payment_user_counts.xlsx'
results_df.to_excel(output_file, index=False)


# In[ ]:




