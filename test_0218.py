# 引入相關套件
from pypfopt.efficient_frontier import EfficientFrontier
from pypfopt import risk_models
from pypfopt import expected_returns
import pandas as pd
import os
import xlwings
import  pandas as pd
import warnings
warnings.filterwarnings("ignore")
import xlwings as xw
import re
import string
from FinMind.data import DataLoader
from pypfopt.risk_models import CovarianceShrinkage

dl = DataLoader()
future_data = dl.taiwan_futures_daily(futures_id='TX', start_date='2020-01-01')
future_data = future_data[(future_data.trading_session == "position")]
future_data = future_data[(future_data.settlement_price > 0)]
future_data = future_data[future_data['contract_date'] == future_data.groupby('date')['contract_date'].transform('min')]
TXF1=future_data

def num_to_col(num):
    """將數字轉換為Excel欄位的英文命名法"""
    col = ''
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        col = string.ascii_uppercase[remainder] + col
    return col

# 打開一個新的Excel文件
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=True
wb = xw.Book('book.xlsx')

sht=wb.sheets('工作表1')
last_cell = sht.range('A2').current_region.rows.count
# 從當前工作表中取得第一個儲存格（A2）
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
returns=dfa
pct_df = pd.DataFrame(index=dfa.index, columns=dfa.columns)
for col in dfa.columns:
    pct_df[col] = dfa[col] / 100000 * 100

# 计算均值和协方差矩阵
mu = returns.mean()
S = returns.cov()

# 定义目标函数
def objective(x):
    return np.dot(x, mu)

# 定义约束条件
def constraint(x):
    return np.sum(x) - 1.0

# 定义初始解
x0 = np.ones(len(mu)) / len(mu)

# 定义边界
bnds = [(0, None) for i in range(len(mu))]

# 运行优化器
cons = {'type': 'eq', 'fun': constraint}
result = minimize(objective, x0, method='SLSQP', bounds=bnds, constraints=cons)

# 输出最优解
weights = pd.Series(result.x, index=returns.columns)
print(weights)

import seaborn as sns
import matplotlib.pyplot as plt

df = pd.DataFrame({'Strategy': np.arange(1,16),
                   'Weight': weights})
# 使用seaborn.barplot()繪製條形圖
ax = sns.barplot(x='Strategy', y='Weight', data=df)

# 在條形上添加標籤
for p in ax.patches:
    ax.annotate(f'{p.get_height():.2f}', (p.get_x() + p.get_width() / 2, p.get_height() + 0.01),
                ha='center')

# 設置圖表標題和坐標軸標籤
ax.set_title('Optimal weights')
ax.set_xlabel('Strategy index')
ax.set_ylabel('Weight')
# 顯示圖形
plt.show()
plt.close()


