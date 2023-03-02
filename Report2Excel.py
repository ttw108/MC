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
        if abs(dfe.iloc[i, j]) > 0.45: # 相关系数绝对值大于 0.5 时添加边
            G.add_edge(dfe.columns[i], dfe.columns[j], weight=dfe.iloc[i, j] )

# 绘制网络图
pos = nx.spring_layout(G, seed=42) # 使用spring布局算法
labels = {node: node for node in G.nodes()}
nx.draw_networkx_labels(G, pos, labels)
nx.draw_networkx_edges(G, pos)
nx.draw_networkx_edge_labels(G, pos, edge_labels={(u, v): round(d["weight"], 2) for u, v, d in G.edges(data=True)})
plt.axis("off")
plt.show()

#plt.show()
plt.savefig('temp.png')
sht_pic = wb.sheets.add('NetworkX')

pic_path=(os.path.join(os.getcwd(), "temp.png"))
sht_pic.pictures.add(pic_path)
plt.close()




print("全部策略："+ str(c-1) + " 支")

#wb.save('book.xlsx')
#wb.close



'''
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=True
wb=app.books.open(order_name)

#單檔操作也可以只用這行快速開啟  wb = xw.Book(order_name)

st1=wb.sheets('Sheet1')
aa=st1.range('A2:A10').value
bb=st1.cells(2,1).value
cc=st1.range('A:A').last_cell
print(st1.used_range.rows.count)

#找到最後一行
last_cell = st1.range('A1').current_region.rows.count
print(last_cell)
'''