import os
import xlwings
import  pandas as pd
import warnings
warnings.filterwarnings("ignore")
import xlwings as xw
import re
import string

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


'''
在上面的代码中，我们使用 drop 函数删除了 dfa 的第一列，方法是将 dfa.columns[0] 作为第一个参数传递给 drop 函数，这是第一列的列名。通过将 axis 参数设置为 1，我们告诉 drop 函数删除列而不是行。
请注意，上面的代码将使用 dfa.columns[0] 作为第一列的列名。如果您的 DataFrame 没有列名，则您需要使用数字索引来指定要删除的列，例如 dfa = dfa.drop(dfa.columns[0], axis=1) 可以替换为 dfa = dfa.drop(dfa.columns[0], axis=1)
'''
import numpy as np
# 讀取交易明細資料
# 轉換為numpy array
returns=np.array(dfa)
mean_returns = np.mean(returns, axis=0)
# 計算各個策略的波動率
volatility = np.std(returns, axis=0)
# 計算各個策略的最大回撤
cum_returns = np.cumsum(returns, axis=0)
max_drawdowns = np.zeros(15)
for i in range(15):
    j = np.argmax(cum_returns[:, i] - np.maximum.accumulate(cum_returns[:, i]))
    if j == 0:
        max_drawdowns[i] = 0
    else:
        max_drawdowns[i] = cum_returns[j, i] - cum_returns[j - 1, i]
# 印出各個策略的平均回報率、波動率和最大回撤
print('Mean returns:', mean_returns)
print('Volatility:', volatility)
print('Max drawdowns:', max_drawdowns)
# 計算最優權重
cov = np.cov(returns, rowvar=False)
inv_cov = np.linalg.inv(cov)
ones = np.ones(15)
w = inv_cov @ ones / (ones.T @ inv_cov @ ones)
# 印出最優權重
print('Optimal weights:', w)
import seaborn as sns
import matplotlib.pyplot as plt

df = pd.DataFrame({'Strategy': np.arange(1,16),
                   'Weight': w})
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










import pandas as pd
import numpy as np
from scipy.optimize import minimize

# 資料讀取
strategies = dfa.columns
dfa = dfa.astype(float)
returns = dfa.values


daily_returns = dfa.pct_change()
# 将 NaN 填充为 0
daily_returns =daily_returns.fillna(0)
# 将 Inf 替换为前一天的值
daily_returns = daily_returns.replace([np.inf, -np.inf], np.nan)
daily_returns = daily_returns.fillna(method='ffill')

# 定义目标函数和约束条件
def max_drawdown(weights):
    portfolio_return = daily_returns.dot(weights)
    cum_returns = np.cumprod(portfolio_return + 1) - 1
    rolling_max = np.maximum.accumulate(cum_returns)
    drawdown = (cum_returns - rolling_max) / (rolling_max + 1)
    return -np.max(drawdown)

cons = ({'type': 'eq', 'fun': lambda x: np.sum(x) - 1})

bounds = [(0, 1) for i in range(daily_returns.shape[1])]

# 使用优化算法求解
result = minimize(max_drawdown, [1/daily_returns.shape[1]]*daily_returns.shape[1], method='SLSQP', bounds=bounds, constraints=cons)

# 输出最佳组合策略的权重向量
weights = result.x

# 计算最佳组合策略的最大回撤
portfolio_return = daily_returns.dot(weights)
cum_returns = np.cumprod(portfolio_return + 1) - 1
rolling_max = np.maximum.accumulate(cum_returns)
drawdown = (cum_returns - rolling_max) / (rolling_max + 1)
min_max_drawdown = np.max(drawdown)

print(f"Best weights: {weights}")
print(f"Minimum Maximum Drawdown: {min_max_drawdown}")