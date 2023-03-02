import pandas as pd
import numpy as np
from scipy.optimize import minimize, LinearConstraint, Bounds
import xlwings as xw
import string
import pandas as pd

# Read data
# 打開一個新的Excel文件
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=True
wb = xw.Book('book.xlsx')

sht=wb.sheets('工作表1')
last_cell = sht.range('A2').current_region.rows.count
# 從當前工作表中取得第一個儲存格（A2）
def num_to_col(num):
    """將數字轉換為Excel欄位的英文命名法"""
    col = ''
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        col = string.ascii_uppercase[remainder] + col
    return col
cell = sht.range('A1')
#以下是針對關連性做排序
last_column=sht.range('A2').current_region.columns.count
rng_all = sht.range('A1:{}'.format(num_to_col(last_column)+str(last_cell)))
dfa =pd.DataFrame(rng_all.value, columns=rng_all.value[0])
dfa = dfa.drop([0])
dfa = dfa.drop(dfa.columns[0], axis=1)


import numpy as np
import pandas as pd
from scipy.optimize import minimize

# 读取数据
df=dfa
# 计算每个策略的平均每笔收益和标准差
avg_returns = df.mean()
std_returns = df.std()

# 定义SQN函数
def sqn(weights, avg_returns, std_returns):
    combined_returns = np.dot(weights, avg_returns)
    combined_std = np.sqrt(np.dot(weights.T, np.dot(np.cov(df.T), weights)))
    sqn = np.sqrt(len(df)) * combined_returns / combined_std
    return -sqn  # 目标函数是最小化负的SQN值

# 定义约束条件
n_assets = len(df.columns)
constraints = [{'type': 'eq', 'fun': lambda x: np.sum(x) - 1}]
bounds = [(0, 1) for i in range(n_assets)]

# 初始权重
weights = np.ones(n_assets) / n_assets

# 最小化负的SQN函数，求得最优权重
result = minimize(sqn, weights, args=(avg_returns, std_returns), method='SLSQP',
                  bounds=bounds, constraints=constraints)

# 打印结果
print('Optimal weights:', result.x)
print('SQN value:', -sqn(result.x, avg_returns, std_returns))

import seaborn as sns
import matplotlib.pyplot as plt

w=result.x
dfx = pd.DataFrame({'Strategy': np.arange(1,w.shape[0]+1),
                   'Weight': w})

# 使用seaborn.barplot()繪製條形圖
plt.figure(figsize=(10, 8))
ax = sns.barplot(x='Strategy', y='Weight', data=dfx)
