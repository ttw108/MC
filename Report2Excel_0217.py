import os
import xlwings
import  pandas as pd
import warnings
warnings.filterwarnings("ignore")
import xlwings as xw
import re
import string

def find_regex_value(sht1, regex_str1, r, c):
    all_rows=1000
    all_columns=300
    for row in range(1,all_rows):
        for col in range(1,all_columns):
            if sht1.cells(row,col).value is not None :
                if re.match(regex_str1, str(sht1.cells(row,col).value)):
                    return (sht1.cells(row + r, col + c).value)


def find_regex_rc(sht1, regex_str1):
    all_rows=1000
    all_columns=300
    for row in range(1,all_rows):
        for col in range(1,all_columns):
            if sht1.cells(row,col).value is not None :
                if re.match(regex_str1, str(sht1.cells(row,col).value)):
                #if cell.value == findtext:
                    print(sht1.cells(row,col).row)
                    print(sht1.cells(row,col).column)
                    return ([sht1.cells(row,col).row, sht1.cells(row,col).column])


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
wb = xw.Book()  # this will open a new workbook
#wb = xw.Book('Book1.xlsx')
#wb=app.books.open('Book1.xlsx')
sht=wb.sheets(1)
# 從當前工作表中取得第一個儲存格（A2）
cell = sht.range('A2')
# 使用Python中的datetime模組來得到當前月份的資訊
import datetime

# 取得當前月份的開始日期
#start_date = datetime.date(datetime.date.today().year, datetime.date.today().month, 1)
# 將開始日期輸入到A1儲存格中
dayrng=90
start_date=datetime.date.today() - datetime.timedelta(days=dayrng)
cell.value = start_date
# 將開始日期往下移動一格，並輸出相對應的日期資訊
sht.cells(1,1).value='日期'
for i in range(1,dayrng):
    cell.offset(row_offset=i,column_offset=0).value = start_date + datetime.timedelta(days=i)
    # 找到記錄XLS"Book1.xls"的最後一行
    # 找到最後一行
last_cell = sht.range('A2').current_region.rows.count

last_cell = last_cell + 1
print(last_cell)

#先找到目錄中    "出貨通知單開頭    .xlsx"    的檔案____
fd="./trading_xls"
files = os.listdir(fd)
fn=1
c=1
for filename in files:
    if filename.split(".")[-1] == "xls":
        if re.match('.*策略回測績效報告.*',filename):
            order_name=filename
            #print(order_name)
            app1 = xw.App(visible=False, add_book=False)
            app1.display_alerts = False
            app1.screen_updating = True
            wb1 = app.books.open(fd+"/"+order_name)
            sht1 = wb1.sheets('交易明細')

            # 將 sheet 內容轉換為 dataframe
            df = sht1.range('B3').options(pd.DataFrame, expand='table').value
            # 移除 column_name 為空的 row
            df.dropna(subset=[r"獲利(¤)"], inplace=True)
            df['日期'] = pd.to_datetime(df['日期']).dt.date


            #90日期中，一個一個日期字串拉出來 篩選 並取得加總：
            sht.cells(1, c+1).value = c
            d0=1
            for dstr in  sht.range('A2:A{}'.format(last_cell)).value:
                profit = df['獲利(¤)'][(df['日期'] == pd.to_datetime(dstr))].sum()
                sht.cells(d0+1,fn+1).value=profit
                d0=d0+1
            wb1.close()
            c=c+1

        sht.cells(95 + fn, 1).value =c
        sht.cells(95+fn, 2).value = filename
        fn=fn+1
        # Quit Excel application


#以下是針對關連性做排序
last_column=sht.range('A2').current_region.columns.count
rng_all = sht.range('A1:{}'.format(num_to_col(last_column)+str(last_cell)))
dfa =pd.DataFrame(rng_all.value, columns=rng_all.value[0])
# 將 '日期' 列轉換為索引
dfa=df = dfa.drop(dfa.index[0])
dfa['日期'] = pd.to_datetime(dfa['日期']).dt.date
dfa = dfa.set_index('日期')

dfa.columns
dfa = dfa.rename(columns=lambda x: str(x).strip())
#針對策略做相關性測試
dfb=dfa.corr().round(2)

sht_corr = wb.sheets.add('corr_1')
sht_corr.range('A1').options(index=True, header=True).value = dfb
# 將列按照相關性排序
dfc = dfb['2.0'].sort_values(ascending=False).index
# 重新排列 DataFrame 中的列
dfd = dfa[dfc]
sht_sortd = wb.sheets.add('re_sortd')
sht_sortd.range('A1').options(index=True, header=True).value = dfd

dfe=dfd.corr().round(2)
sht_corr_2 = wb.sheets.add('corr2')
sht_corr_2.range('A1').options(index=True, header=True).value = dfe


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
                ha='center', color='red', fontsize=12)

# 設置圖表標題和坐標軸標籤
ax.set_title('MaxProfit/MMSE Optimal weights')
ax.set_xlabel('Strategy index')
ax.set_ylabel('Weight')
# 顯示圖形
plt.show()
sht_pic = wb.sheets.add('OPT_Bar')
plt.savefig('temp.png')
pic_path=(os.path.join(os.getcwd(), "temp.png"))

sht_pic.pictures.add(pic_path)
plt.close()
# 删除临时文件
os.remove('temp.png')

#接下來做主成份分析
from sklearn.decomposition import PCA
returns=dfa
pca=PCA()
pca.fit(returns)

# 输出主成分分析结果
print('Explained variance ratio:', pca.explained_variance_ratio_)
print('Principal components:', pca.components_)
cc=pca.components_
#就此結果PCA繪圖
import matplotlib.pyplot as plt
# 绘制主成分分析结果散点图
plt.scatter(pca.components_[0], pca.components_[1])

# 添加坐标轴标签
plt.xlabel('PC1')
plt.ylabel('PC2')

# 添加每个交易策略的标签
strategies=list(returns.columns)

for i, strategy in enumerate(strategies):
    plt.annotate(strategy, (pca.components_[0][i], pca.components_[1][i]))

# 显示图形
pp=plt.plot
#plt.show()
plt.savefig('temp.png')
sht_pic = wb.sheets.add('PCA')

pic_path=(os.path.join(os.getcwd(), "temp.png"))
sht_pic.pictures.add(pic_path)
plt.close()
# 删除临时文件
os.remove('temp.png')
# 显示Excel文件
wb.app.visible = True

import seaborn as sns
sns.set(font_scale=0.5)
sns.heatmap(dfe, annot=True, cmap='coolwarm')
plt.savefig('temp.png')
sht_pic = wb.sheets.add('HeatMap')

pic_path=(os.path.join(os.getcwd(), "temp.png"))
sht_pic.pictures.add(pic_path)
plt.close()
# 删除临时文件
os.remove('temp.png')
#召入散點矩陣圖
import seaborn as sns
# 绘制散点矩阵图
#sns.pairplot(dfe, diag_kind='hist')
# 显示图像
plt.show()
plt.close()
from scipy.cluster import hierarchy
#dendrogram = hierarchy.dendrogram(hierarchy.linkage(dfe))


import networkx as nx
# 创建图
G = nx.Graph()

# 添加节点
for col in dfe.columns:
    G.add_node(col)
# 添加边
for i in range(len(dfe.columns)):
    for j in range(i + 1, len(dfe.columns)):
        if abs(dfe.iloc[i, j]) > 0.3:
            G.add_edge(dfe.columns[i], dfe.columns[j], weight=dfe.iloc[i, j])
# 绘制网络图
pos = nx.circular_layout(G)
# 繪製節點
nx.draw_networkx_nodes(G, pos, node_size=400, node_color='lightblue', node_shape='o', linewidths=1)
# 繪製節點標籤
labels = {node: node for node in G.nodes()}
nx.draw_networkx_labels(G, pos, labels, font_color='red')
# 繪製邊
nx.draw_networkx_edges(G, pos, style="dashed")

# 繪製邊權重標籤
nx.draw_networkx_edge_labels(G, pos, edge_labels={(u, v): round(d["weight"], 1) for u, v, d in G.edges(data=True)})
# 顯示圖形
plt.show()
plt.savefig('temp.png')
sht_pic = wb.sheets.add('NetworkX')
pic_path=(os.path.join(os.getcwd(), "temp.png"))
sht_pic.pictures.add(pic_path)
plt.close()





from scipy.optimize import minimize
returns=dfa
pct_df = pd.DataFrame(index=dfa.index, columns=dfa.columns)
for col in dfa.columns:
    pct_df[col] = dfa[col] / 100000 * 100
# 给定权重，求组合收益率、标准差、夏普比率
# 定义无风险收益率
rf = 0.02
# 获取股票平均收益率
mean_returns  = dfa.mean()
# 获取股票收益率的方差协方差矩阵
cov_matrix  = dfa.cov()
# 定义资产数量
number_assets = dfa.shape[1]

# 目标收益率约束条件
target_return = 0.5
# 最小化方差的目标函数
def portfolio_volatility(weights, cov_matrix):
    return np.sqrt(np.dot(weights.T, np.dot(cov_matrix, weights)))
# 约束条件
def constraint(weights, mean_returns, target_return):
    return np.sum(mean_returns * weights) - target_return
# 初始权重向量
n_assets = len(mean_returns)
weights_0 = np.ones(n_assets) / n_assets

# 定义边界条件
bounds = tuple((0, 1) for _ in range(n_assets))

# 定义约束条件
cons = ({'type': 'eq', 'fun': constraint, 'args': (mean_returns, target_return)})

# 调用 minimize 函数进行投资组合优化
opt_result = minimize(portfolio_volatility, weights_0, args=(cov_matrix,),
                      method='SLSQP', bounds=bounds, constraints=cons)

# 输出优化结果
print(opt_result.x)

####################################################################
####################################################################
import seaborn as sns
import matplotlib.pyplot as plt

df = pd.DataFrame({'Strategy': np.arange(1,16),
                   'Weight': opt_result.x })
# 使用seaborn.barplot()繪製條形圖
ax = sns.barplot(x='Strategy', y='Weight', data=df)

# 在條形上添加標籤
for p in ax.patches:
    ax.annotate(f'{p.get_height():.2f}', (p.get_x() + p.get_width() / 2, p.get_height() + 0.01),
                ha='center')

# 設置圖表標題和坐標軸標籤
ax.set_title('Optimal Least Square Error')
ax.set_xlabel('Strategy index')
ax.set_ylabel('Weight')
# 顯示圖形
plt.show()
plt.savefig('temp.png')
sht_pic = wb.sheets.add('MMSE Optimal Bar')
pic_path=(os.path.join(os.getcwd(), "temp.png"))
sht_pic.pictures.add(pic_path)
plt.close()

print("全部策略："+ str(c-1) + " 支")
wb.save('book.xlsx')
wb.close
xw.apps.active.quit()


