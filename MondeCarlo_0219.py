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
import matplotlib.pyplot as plt
import seaborn as sns
#SQN##################################SQN#################################
#SQN##################################SQN#################################
#SQN##################################SQN#################################
# 读取数据
dfa.columns
dfa = dfa.rename(columns=lambda x: str(x).strip())
dfs=dfa
# 计算每个策略的平均每笔收益和标准差
avg_returns = dfs.mean()
std_returns = dfs.std()
# 定义SQN函数
def sqn(weights, avg_returns, std_returns):
    combined_returns = np.dot(weights, avg_returns)
    combined_std = np.sqrt(np.dot(weights.T, np.dot(np.cov(dfs.T), weights)))
    sqn = np.sqrt(len(dfs)) * combined_returns / combined_std
    return -sqn  # 目标函数是最小化负的SQN值

# 定义约束条件
n_assets = len(dfs.columns)
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
ws=result.x
dfx = pd.DataFrame({'Strategy': np.arange(1,ws.shape[0]+1),
                   'Weight': ws})

# 使用seaborn.barplot()繪製條形圖
plt.figure(figsize=(10, 8))
plt.rcParams.update({'font.size': 12})
ax = sns.barplot(x='Strategy', y='Weight', data=dfx)
# 在條形上添加標籤
for pa in ax.patches:
    ax.annotate(f'{pa.get_height():.2f}', (pa.get_x() + pa.get_width() / 2, pa.get_height() + 0.01),
                ha='center', color='black', fontsize=10)

# 設置圖表標題和坐標軸標籤

ax.set_title('SQN OPT Portfolio',fontdict={'fontsize': 18})
ax.set_xlabel('Strategy index',fontdict={'fontsize': 10})
ax.set_ylabel('Weight',fontdict={'fontsize': 10})
# 顯示圖形
plt.show()
plt.tight_layout()
plt.close()
hf=int((ws.shape[0])/2)
sqn_df = pd.DataFrame({'Strategy': np.arange(1,ws.shape[0]+1),'Weight': ws })
sqn_df=sqn_df.sort_values(by='Weight',ascending=False)
sqn7=sqn_df.iloc[0:hf]
sqn7 = sqn7[sqn7['Weight'] > 0.01]
# 将所有列名中的 ".0" 替换为 ""
dfa_s= dfa.copy()
dfa_s.rename(columns=lambda x: x.replace('.0', ''), inplace=True)
s7=sqn7.Strategy
lis_s7=s7.tolist()
lis_s7a = [str(x) for x in lis_s7]
s7_df = dfa_s.loc[:, lis_s7a].copy()
# 将每一行加总并存储为一个新的 Series
s7_sum = s7_df.sum(axis=1)
# 将新的 Series 添加到 DataFrame 中
s7_df['s7 Sum'] = s7_sum
# 計算B欄逐行加總，並新增一個欄位'C'保存結果
s7_df['cum_Sum'] = s7_df['s7 Sum'].cumsum()

s7_df_s=s7_df.copy()
s7_df_s['cum_Sum']=s7_df_s['cum_Sum']+600000
s7_df_s['cum_Sum_p']=s7_df_s['cum_Sum'].pct_change().dropna()


#s7_df['cum_Sum']=s7_df['cum_Sum']+600000
#s7_df = s7_df.drop('cum_Sum', axis=1)
#s7_df = s7_df.drop('s7 Sum', axis=1)

new_df = pd.DataFrame()

for column in s7_df.columns:
    # 計算累積損益
    cum_profit = s7_df[column].cumsum()
    # 加上本金
    cum_profit_with_capital = cum_profit + 100000
    # 把結果存儲到新 DataFrame
    new_df[column] = cum_profit_with_capital












#MondeCarlo
import numpy as np
import pandas as pd

# 假設您的原始dataframe為df，欄位為策略名，列為日期
# 先將原始dataframe中小於0的值設為0
#df[df < 0] = 0
new_df1=new_df.drop(new_df.columns[[-2,-1]], axis=1)
s7_dfa=new_df1

# 計算收益率
df_returns = s7_dfa.pct_change().dropna()

# 定義投資組合數量和蒙地卡羅模擬次數
n_portfolios = 30
n_iterations = 90

# 創建空的 DataFrame 來存儲模擬結果
simulated_returns = pd.DataFrame()

# 開始模擬
for i in range(n_portfolios):
    # 隨機生成一個權重向量
    weights = np.random.random(6)
    weights /= np.sum(weights)

    # 計算組合收益率
    portfolio_returns = (df_returns * weights).sum(axis=1)

    # 將收益率加入到 simulated_returns 中
    simulated_returns[f"Portfolio {i + 1}"] = portfolio_returns.values

# 輸出模擬結果
simulated_returns.to_csv("simulated_returns.csv", index=False)

# 計算每個投資組合的平均收益率和標準差
portfolio_means = simulated_returns.mean()
portfolio_stds = simulated_returns.std()

# 顯示平均收益率和標準差
print("Average Returns:")
print(portfolio_means)
print("Standard Deviations:")
print(portfolio_stds)

# 繪製累積收益圖
cumulative_returns = (1 + simulated_returns).cumprod()
cumulative_returns.plot(legend=False)
plt.xlabel("Day")
plt.ylabel("Cumulative Returns")
plt.title("Simulated Cumulative Returns")
plt.tight_layout()
plt.show()
# 獲取 AxesSubplot 物件
ax = plt.gca()
# 使用 setp 函數設置折線粗細
plt.setp(ax.lines, linewidth=0.5)

# 重新顯示圖形
plt.show()
sns.lineplot(x=s7_df.index, y='cum_Sum_p', data=s7_df_s,linestyle="-",color="purple",label="Mail")
