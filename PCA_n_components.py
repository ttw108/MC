from tkinter import *
import tkinter as tk
from tkinter import simpledialog
import warnings
warnings.filterwarnings("ignore")
import xlwings as xw
import re
import string
import pandas as pd
import os
from sklearn.preprocessing import MaxAbsScaler
from sklearn.decomposition import PCA
import matplotlib.pyplot as plt


root = tk.Tk()
root.withdraw()
print(root.tk.exprstring('$tcl_library'))
print(root.tk.exprstring('$tk_library'))
# Create an input dialog
n_conponents= simpledialog.askstring("Input", "選擇PCA n_conponents:",parent=root)
n_conponents=[int(ix) for ix in n_conponents.split(' ')]
n_conponents = list(map(str, n_conponents))

#將bookmark開啟
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
app.screen_updating=False
#wb = xw.Book()  # this will open a new workbook
wb = xw.Book('book.xlsx')
#wb=app.books.open('Book1.xlsx')
sht=wb.sheets("工作表1")
# 從當前工作表中取得第一個儲存格（A2）
cell = sht.range('A2')

last_column = sht.range('A2').current_region.columns.count
last_row = sht.range('A2').current_region.rows.count

rng_all = sht.range('A1:{}'.format(num_to_col(last_column) + str(last_row)))
dfa = pd.DataFrame(rng_all.value, columns=rng_all.value[0])

# 將 '日期' 列轉換為索引

dfa = dfa.drop(dfa.index[0])
dfa['日期'] = pd.to_datetime(dfa['日期']).dt.date
dfa = dfa.set_index('日期')

dfa = dfa.rename(columns=lambda x: str(x).strip())
dfa.rename(columns=lambda x: x.replace('.0', ''), inplace=True)

# 假設您的DataFrame名稱為dfa
# 先將第一列加上本錢100000元
dfa.iloc[0] = dfa.iloc[0] + 100000

# 對每一欄做cumsum
dfa = dfa.apply(lambda x: x.cumsum(), axis=0)
dfa = dfa.drop(dfa.index[-1])

#1 處理日期 開啟 book.xlsx '新增 profit 工作表 、加入日期180天
def open_xlsx():
    # 打開一個新的Excel文件
    global wb
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = True
    wb = xw.Book('book.xlsx')
    global sht_profit
    if not 'n_components' in wb.sheet_names:
        sht_profit = wb.sheets.add('n_components')
    else:
        wb.sheets['n_components'].delete()
        sht_profit = wb.sheets.add('n_components')
    print("n_components_工作表_日期建立完成")
open_xlsx()

from joblib import dump, load
# 讀取檔案
sqn_lis1 = load('sqn_a1_lis1.joblib')


def pca_f(df_, color_list, label_dict,n_c):
    from sklearn.preprocessing import MaxAbsScaler
    #returns = df_
    # 創建MaxAbsScaler對象
    scaler = MaxAbsScaler()
    # 對稀疏數據進行標準化
    data_scaled = scaler.fit_transform(df_)
    returns = pd.DataFrame(data_scaled, columns=df_.columns)
    pca = PCA(n_components=int(n_c))
    pca.fit(returns)
    #plt.tight_layout()
    # 输出主成分分析结果
    print('Explained variance ratio:', pca.explained_variance_ratio_)
    print('Principal components:', pca.components_)
    cc = pca.components_
    # 就此結果PCA繪圖
    import matplotlib.pyplot as plt
    # 绘制主成分分析结果散点图
    # 设置绘图区域的背景色为淡黄色
    plt.figure(figsize=(6, 4))
    # 绘制 sqn7 的策略点
    #plt.scatter(pca.components_[0], pca.components_[1], c='red', s=30, marker='o', facecolors='none')
    # 添加坐标轴标签
    plt.title('PCA_{}'.format(color_list), fontsize=12)
    plt.xlabel('PC1')
    plt.ylabel('PC2')
    # 添加每个交易策略的标签
    strategies = list(returns.columns)
    #color_list=sqn_a1_lis1
    # strategies = [x[:-2] if x.endswith('.0') else x for x in strategies]
    for i, strategy in enumerate(strategies):
        if strategy in color_list:
            plt.scatter(pca.components_[0][i], pca.components_[1][i], c='red', s=20, marker='o', facecolors='none')
        else:
            plt.scatter(pca.components_[0][i], pca.components_[1][i], c='blue', s=10, marker='o', facecolors='none')
        #plt.annotate(strategy, (pca.components_[0][i], pca.components_[1][i]), color='blue', fontsize=12)
        # 检查该策略的标签是否在label_dict中，如果是，则使用label_dict中的新标签
        if strategy in label_dict:
            label = strategy
            plt.annotate(label, (pca.components_[0][i], pca.components_[1][i]), color='red', fontsize=18)
        else:
            label = strategy
            plt.annotate(label, (pca.components_[0][i], pca.components_[1][i]), color='green', fontsize=10)
        # 显示图形
        # 設置座標軸的格線顏色
    plt.grid(color='#D3D3D3')
    # 将绘图区域的背景色改为淡黄色
    plt.gca().set_facecolor('#ffffcc')
    pp = plt.plot
    plt.tight_layout()

if 'n_components' not in [s.name for s in wb.sheets]:
    sht_pic = wb.sheets.add('n_components')
else:
    wb.sheets['n_components'].delete()
    sht_pic = wb.sheets.add('n_components')
    sht_pic = wb.sheets['n_components']

#dfaa##########################################################
for ii in n_conponents:
    pca_f(dfa, sqn_lis1, sqn_lis1,int(ii))
    plt.tight_layout()
    # plt.show()
    plt.savefig('temp.png')
    plt.close()
    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    sht_pic.pictures.add(pic_path)
    sht_pic.pictures.add(pic_path, name='dfa', left=sht_pic.range('A1').left, top=sht_pic.range('A1').top)
    ################################################################