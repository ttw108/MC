import os
import xlwings
import  pandas as pd
import warnings
from sklearn.decomposition import PCA
import difflib
warnings.filterwarnings("ignore")
import xlwings as xw
import re
import string
import seaborn as sns
import matplotlib.pyplot as plt
from scipy.optimize import minimize
import numpy as np


def num_to_col(num):
    """將數字轉換為Excel欄位的英文命名法"""
    col = ''
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        col = string.ascii_uppercase[remainder] + col
    return col

#1 處理日期 開啟 book.xlsx '新增 profit 工作表 、加入日期180天
def open_xlsx():
    # 打開一個新的Excel文件
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = True
    wb = xw.Book('book.xlsx')
    global sht_profit
    if not 'profit' in wb.sheet_names:
        sht_profit = wb.sheets.add('profit')
    else:
        wb.sheets['profit'].delete()
        sht_profit = wb.sheets.add('profit')

    # 從當前工作表中取得第一個儲存格（A2）
    cell = sht_profit.range('A2')
    # 使用Python中的datetime模組來得到當前月份的資訊
    import datetime
    # 第一步基本資料處理
    # 將開始日期輸入到A1儲存格中
    dayrng = 180
    start_date = datetime.date.today() - datetime.timedelta(days=dayrng)
    cell.value = start_date
    # 將開始日期往下移動一格，並輸出相對應的日期資訊
    sht_profit.cells(1, 1).value = '日期'

    for i in range(1, dayrng):
        cell.offset(row_offset=i, column_offset=0).value = start_date + datetime.timedelta(days=i)
        # 找到記錄XLS"Book1.xls"的最後一行
        # 找到最後一行
    global last_cell
    last_cell = sht_profit.range('A2').current_region.rows.count
    print("Profit_工作表_日期建立完成")
open_xlsx()

#2先找到目錄中    ".xls" 把180天交易損益寫進 PROFIT工作表
def fill_strategy_profit():
    fd = "./trading_xls"
    files = os.listdir(fd)
    fn = 1
    global c
    c = 1
    for filename in files:
        if filename.split(".")[-1] == "xls":
            if re.match('.*策略回測績效報告.*', filename):
                order_name = filename
                # print(order_name)
                app1 = xw.App(visible=False, add_book=False)
                app1.display_alerts = False
                app1.screen_updating = False
                wb1 = app1.books.open(fd + "/" + order_name)
                sht1 = wb1.sheets('交易明細')

                # 將 sheet 內容轉換為 dataframe
                df = sht1.range('B3').options(pd.DataFrame, expand='table').value
                # 移除 column_name 為空的 row
                df.dropna(subset=[r"獲利(¤)"], inplace=True)
                df['日期'] = pd.to_datetime(df['日期']).dt.date

                # 180日期中，一個一個日期字串拉出來 篩選 並取得加總：
                sht_profit.cells(1, c + 1).value = c
                d0 = 1
                for dstr in sht_profit.range('A2:A{}'.format(last_cell)).value:
                    profit = df['獲利(¤)'][(df['日期'] == pd.to_datetime(dstr))].sum()
                    sht_profit.cells(d0 + 1, fn + 1).value = profit
                    d0 = d0 + 1
                wb1.close()
                c = c + 1

            sht_profit.cells(200 + fn, 1).value = c - 1
            sht_profit.cells(200 + fn, 2).value = filename
            print(str(fn)+": "+ filename)
            fn = fn + 1
fill_strategy_profit()

def conv_to_df():
    last_column = sht_profit.range('A2').current_region.columns.count
    rng_all = sht_profit.range('A1:{}'.format(num_to_col(last_column) + str(last_cell)))
    global dfa
    dfa = pd.DataFrame(rng_all.value, columns=rng_all.value[0])
    # 將 '日期' 列轉換為索引
    dfa = df = dfa.drop(dfa.index[0])
    dfa['日期'] = pd.to_datetime(dfa['日期']).dt.date
    dfa = dfa.set_index('日期')
    dfa = dfa.rename(columns=lambda x: str(x).strip())
    dfa = dfa.rename(columns=lambda x: x[:-2] if x.endswith('.0') else x)
    dfa.to_pickle('dfa.pkl')
    pd.read_pickle('dfa.pkl')
conv_to_df()

########################################################################################################################
########################################################################################################################
########################################################################################################################################################




import os
import xlwings
import  pandas as pd
import warnings
from sklearn.decomposition import PCA
import difflib
warnings.filterwarnings("ignore")
import xlwings as xw
import re
import string
import seaborn as sns
import matplotlib.pyplot as plt
from scipy.optimize import minimize
import numpy as np
###000000000可從這裡開始
################################################################
dfa = pd.read_pickle('dfa.pkl')
df_all = pd.read_pickle('df_all.pkl')
################################################################
#選出平加總及加權加總的優良策略組合
def sqn_func(dfa):
    # 读取数据
    dfs = dfa
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
    ws = result.x
    dfx = pd.DataFrame({'Strategy': np.arange(1, ws.shape[0] + 1),
                        'Weight': ws})
    return(dfx)

#將 a1,b1,c1 做平均加總(a1,b1,c1)及加權加總（a1*0.5+b1*0.3+c1*0.2）
def sqn_abc_all_export():
    global df_all
    # 本週90日的SQN
    dfa.tail(1).index
    df_a1 = dfa.iloc[-90:, :]
    # 從最後一天往前推 10 天，再選擇 90 天的範圍
    df_b1 = dfa.iloc[-110:-20, :]
    # 從最後一天往前推 20 天，再選擇 90 天的範圍
    df_c1 = dfa.iloc[-130:-40, :]
    # 從最後一天往前推 20 天，再選擇 90 天的範圍
    df_d1 = dfa.iloc[-150:-60, :]

    global  sqn_a1, sqn_b1, sqn_c1, sqn_d1
    sqn_a1 = sqn_func(df_a1)
    sqn_b1 = sqn_func(df_b1)
    sqn_c1 = sqn_func(df_c1)
    sqn_d1 = sqn_func(df_d1)

    sqn_a1.to_pickle('sqn_a1.pkl')
    sqn_b1.to_pickle('sqn_b1.pkl')
    sqn_c1.to_pickle('sqn_c1.pkl')
    sqn_d1.to_pickle('sqn_d1.pkl')
    # test=pd.read_pickle('sqn_a1.pkl')
sqn_abc_all_export()


# 策略名稱
# 選擇要導出的範圍

def data_load():
    global sqn_lis1, sqn_lis2,sqnb1_sum_cumsum0,sqnb2_sum_cumsum0,sqn_d1_b_cumsum0,sqn_cd_b_cumsum0
    df_all = pd.DataFrame()
    df_all['Strategy'] = sqn_a1['Strategy'].copy()
    app = xw.apps.active
    global sht_profit
    if app is not None:
        print("!")
        app = xw.App(visible=True, add_book=False)
        app.display_alerts = False
        app.screen_updating = True
        wb = xw.Book('book.xlsx')
        sht_profit = wb.sheets('profit')
    else:
        app = xw.App(visible=True, add_book=False)
        app.display_alerts = False
        app.screen_updating = True
        wb = xw.Book('book.xlsx')
        sht_profit = wb.sheets('profit')

    last_row = sht_profit.range('A201').current_region.rows.count + 200
    rangevalue = sht_profit.range('B201:B{}'.format(last_row))
    df_all['name'] = rangevalue.value
    df_all['a1'] = sqn_a1['Weight'].copy()
    df_all['b1'] = sqn_b1['Weight'].copy()
    df_all['c1'] = sqn_c1['Weight'].copy()
    df_all['d1'] = sqn_d1['Weight'].copy()
    df_all['sum'] = df_all.iloc[:, 2:5].sum(axis=1)
    weights = [0.5, 0.3, 0.2]
    df_all['weighted_sum'] = (df_all.iloc[:, 2:5] * weights).sum(axis=1)
    df_all.to_pickle('df_all.pkl')
    df_all = pd.read_pickle('df_all.pkl')

    df_sqn_sort1 = df_all.sort_values('sum', ascending=False)
    hf = int(len(df_sqn_sort1) / 2)  # 取策略的一半
    sqn_sun_best = df_sqn_sort1.iloc[:hf, :]
    sqn_sun_best = sqn_sun_best[sqn_sun_best['sum'] > 0.01]

    df_sqn_sort2 = df_all.sort_values('weighted_sum', ascending=False)
    sqn_wsun_best = df_sqn_sort2.iloc[:hf, :]
    sqn_wsun_best = sqn_sun_best[sqn_sun_best['weighted_sum'] > 0.01]

    ##################################################
    ##################################################
    sqn_lis1 = sqn_sun_best['Strategy'].to_list()
    # 使用 map() 和 str() 函数将数字列表转换为字符串列表
    sqn_lis1 = list(map(str, sqn_lis1))
    sqnb1 = dfa.loc[:, sqn_lis1]
    sqnb1 = sqnb1.tail(60)
    sqnb1_sum = sqnb1.groupby(sqnb1.columns, axis=1).sum().cumsum()
    # 计算累积总和
    sqnb1_sum_cumsum0 = sqnb1.sum(axis=1).cumsum()
    ##################################################
    ##################################################
    sqn_lis2 = sqn_wsun_best['Strategy'].to_list()
    # 使用 map() 和 str() 函数将数字列表转换为字符串列表
    sqn_lis2 = list(map(str, sqn_lis2))
    sqnb2 = dfa.loc[:, sqn_lis2]
    sqnb2 = sqnb2.tail(60)
    sqnb2_sum = sqnb2.groupby(sqnb2.columns, axis=1).sum().cumsum()
    # 计算累积总和
    sqnb2_sum_cumsum0 = sqnb2.sum(axis=1).cumsum()

    ##################################################
    from joblib import dump, load
    # 寫入檔案
    dump(sqn_lis1, 'sqn_list1.joblib')
    dump(sqn_lis2, 'sqn_list2.joblib')
    # 讀取檔案
    s1_list = load('sqn_list1.joblib')
    s2_list = load('sqn_list2.joblib')
data_load()

def sqn_d1_cd():
    ################################################################
    # 用D1的最佳策略來評估
    global sqn_d1_lis1,sqn_cd_lis1
    df_sqn_d1_sort = df_all.sort_values('d1', ascending=False)
    hf = int(len(df_sqn_d1_sort) / 2)  # 取策略的一半
    sqn_d1_best = df_sqn_d1_sort.iloc[:hf, :]
    sqn_d1_best = sqn_d1_best[sqn_d1_best['d1'] > 0.01]

    sqn_d1_lis1 = sqn_d1_best['Strategy'].to_list()
    # 使用 map() 和 str() 函数将数字列表转换为字符串列表
    sqn_d1_lis1 = list(map(str, sqn_d1_lis1))
    sqn_d1_b = dfa.loc[:, sqn_d1_lis1]
    sqn_d1_b = sqn_d1_b.tail(60)
    sqn_d1_sum = sqn_d1_b.groupby(sqn_d1_b.columns, axis=1).sum().cumsum()
    # 计算累积总和
    sqn_d1_b_cumsum0 = sqn_d1_b.sum(axis=1).cumsum()
    ########################################################################
    ########################################################################
    # 用C1+D1的最佳策略來評估
    df_all['cd_sum'] = df_all.iloc[:, 4:6].sum(axis=1)
    weights = [0.7, 0.3]
    df_all['weighted_cd_sum'] = (df_all.iloc[:, 4:6] * weights).sum(axis=1)
    df_sqn_cd_sort = df_all.sort_values('weighted_cd_sum', ascending=False)
    hf = int(len(df_sqn_cd_sort) / 2)  # 取策略的一半
    sqn_cd_wsun_best = df_sqn_cd_sort.iloc[:hf, :]
    sqn_cd_wsun_best = sqn_cd_wsun_best[sqn_cd_wsun_best['sum'] > 0.01]

    sqn_cd_lis1 = sqn_cd_wsun_best['Strategy'].to_list()
    # 使用 map() 和 str() 函数将数字列表转换为字符串列表
    sqn_cd_lis1 = list(map(str, sqn_cd_lis1))
    sqn_cd_b = dfa.loc[:, sqn_cd_lis1]
    sqn_cd_b = sqn_cd_b.tail(60)
    sqn_cd_sum = sqn_cd_b.groupby(sqn_cd_b.columns, axis=1).sum().cumsum()

    # 计算累积总和
    sqn_cd_b_cumsum0 = sqn_cd_b.sum(axis=1).cumsum()
sqn_d1_cd()

################################################################
################################################################
#PCA
def corr_f():
    # 以下是針對關連性做排序
    # 針對策略做相關性測試
    global dfb,dfaa,dfab,dfac,dfad
    dfaa=dfa.iloc[-90:,:]
    dfab=dfa.iloc[-110:-20,:]
    dfac = dfa.iloc[-130:-40, :]
    dfad = dfa.iloc[-150:-60, :]
    df_corr_a = dfaa.corr().round(2)
    df_corr_b = dfab.corr().round(2)
    df_corr_c = dfac.corr().round(2)
    df_corr_d = dfad.corr().round(2)
corr_f()

def pca_f(df_, color_list, label_dict):
    returns = df_
    pca = PCA()
    pca.fit(returns)
    #plt.tight_layout()
    # 输出主成分分析结果
    print('Explained variance ratio:', pca.explained_variance_ratio_)
    print('Principal components:', pca.components_)
    cc = pca.components_
    # 就此結果PCA繪圖
    import matplotlib.pyplot as plt
    # 绘制主成分分析结果散点图
    plt.figure(figsize=(6, 4))
    # 绘制 sqn7 的策略点
    #plt.scatter(pca.components_[0], pca.components_[1], c='red', s=30, marker='o', facecolors='none')
    # 添加坐标轴标签
    plt.title('PCA_{}'.format(color_list), fontsize=12)
    plt.xlabel('PC1')
    plt.ylabel('PC2')
    # 添加每个交易策略的标签
    strategies = list(returns.columns)
    #color_list=sqn_lis1
    # strategies = [x[:-2] if x.endswith('.0') else x for x in strategies]
    for i, strategy in enumerate(strategies):
        if strategy in color_list:
            plt.scatter(pca.components_[0][i], pca.components_[1][i], c='red', s=30, marker='o', facecolors='none')
        else:
            plt.scatter(pca.components_[0][i], pca.components_[1][i], c='blue', s=10, marker='o', facecolors='none')
        #plt.annotate(strategy, (pca.components_[0][i], pca.components_[1][i]), color='blue', fontsize=12)
        # 检查该策略的标签是否在label_dict中，如果是，则使用label_dict中的新标签
        if strategy in label_dict:
            label = strategy
            plt.annotate(label, (pca.components_[0][i], pca.components_[1][i]), color='red', fontsize=14)
        else:
            label = strategy
        plt.annotate(label, (pca.components_[0][i], pca.components_[1][i]), color='green', fontsize=8)



        # 显示图形
    pp = plt.plot
    plt.tight_layout()

pca_f(dfaa,sqn_lis1,sqn_lis1)
pca_f(dfad,sqn_d1_lis1,sqn_d1_lis1)
pca_f(dfab,sqn_cd_lis1,sqn_cd_lis1)

########################################################################
########################################################################
#t-sne
from sklearn.manifold import TSNE
import matplotlib.pyplot as plt
def tsne_f(df_, target_col):
    df_t = df_.T # 转置 dataframe
    tsne = TSNE(n_components=2, verbose=1, perplexity=3, n_iter=1000)
    tsne_results = tsne.fit_transform(df_t)
    df_tsne = pd.DataFrame({'X': tsne_results[:, 0], 'Y': tsne_results[:, 1], target_col: df_t.index})
    plt.figure(figsize=(10, 10))
    sns.scatterplot(x="X", y="Y", hue=target_col, palette=sns.color_palette("hls", len(df_tsne[target_col].unique())), data=df_tsne, legend="full", alpha=0.3)
    plt.title('t-SNE plot')
    plt.show()
tsne_f(dfa,','.join(dfa.columns))


import umap.umap_ as umap
def umap_f(df_, target_col, labels):
    df_t = df_.T # 转置 dataframe
    umap_results = umap.UMAP(n_neighbors=500, min_dist=1, n_components=10, repulsion_strength=0.1, learning_rate=0.001).fit_transform(df_t)
    df_umap = pd.DataFrame({'X': umap_results[:, 0], 'Y': umap_results[:, 1], target_col: df_t.index, 'label': labels})
    plt.figure(figsize=(6, 6))
    sns.scatterplot(x="X", y="Y", style=target_col, hue=target_col, palette=sns.color_palette("hls", len(df_umap[target_col].unique())), data=df_umap, alpha=0.9, legend=False)
    plt.title('UMAP plot')
    plt.show()
umap_f(dfa,','.join(dfa.columns),list(dfa.columns))

#TTTTTTTTTTTTTTTTTTTT
def umap_f(df_, target_col, labels):
    df_t = df_.T  # 转置 dataframe
    umap_results = umap.UMAP(n_neighbors=500, min_dist=1, n_components=10, repulsion_strength=0.1,
                             learning_rate=0.001).fit_transform(df_t)
    df_umap = pd.DataFrame({'X': umap_results[:, 0], 'Y': umap_results[:, 1], target_col: df_t.index, 'label': labels})

    # 绘制散点图
    plt.figure(figsize=(6, 6))
    colors = sns.color_palette("hls", len(df_umap[target_col].unique()))  # 定义颜色列表
    for i, label in enumerate(df_umap[target_col].unique()):
        x = df_umap[df_umap[target_col] == label]['X']
        y = df_umap[df_umap[target_col] == label]['Y']
        plt.scatter(x, y, c=colors[i], label=str(label))  # 将label参数转换为字符串类型并使用它
    plt.legend()
    plt.title('UMAP plot')
    plt.show()
umap_f(dfa, ','.join(dfa.columns), list(dfa.columns))






########################################################################
########################################################################

##################################################
# 绘制折线图
#ax = sqnb1_sum.plot.line(linewidth=0.5, alpha=0.15)
plt.figure(figsize=(10, 8))
sns.lineplot(x=sqnb1_sum_cumsum0.index, y=sqnb1_sum_cumsum0[:], data=sqnb1_sum_cumsum0, linestyle="-", color="red", label="abc_combined-{}")
sns.lineplot(x=sqnb2_sum_cumsum0.index, y=sqnb2_sum_cumsum0[:], data=sqnb2_sum_cumsum0, linestyle="--", color="gray", label="abc_Wcombined-{}")
sns.lineplot(x=sqn_d1_b_cumsum0.index, y=sqn_d1_b_cumsum0[:], data=sqn_d1_b_cumsum0, linestyle=":", color="green", label="d1-60Day ago-{}")
sns.lineplot(x=sqn_cd_b_cumsum0.index, y=sqn_cd_b_cumsum0[:], data=sqn_cd_b_cumsum0, linestyle="--", color="blue", label="c1,d1 Weighted-40-60Day{}")
# 繪製垂直線
last_date = sqnb1_sum_cumsum0.index[-1]
line_date = last_date - pd.DateOffset(days=40)
plt.axvline(x=line_date, linestyle="--", color="red", linewidth=0.3)


plt.xlabel("Date", fontdict={'fontsize': 10})
plt.ylabel("Total Return", fontdict={'fontsize': 10})
plt.title("SQN_rolling", fontdict={'fontsize': 18})
# 在 x 轴上绘制垂直线
#ax.axvline(x=70, color='r')
plt.legend(loc='upper left', borderaxespad=0., fontsize='large', fancybox=True, edgecolor='navy', framealpha=0.2,
           handlelength=1.5, handletextpad=0.5, borderpad=0.5, labelspacing=0.5)

plt.rcParams['xtick.labelsize']=8
plt.rcParams['ytick.labelsize']=8
plt.tight_layout()
plt.show()
##################################################
##################################################
