from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter.filedialog import askdirectory
from PIL import ImageTk, Image
from tkinter import messagebox
from datetime import date
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import os
from sys import api_version
from Google import Create_Service

##patchnote 20220411
## 매출: 판매처-> 전체 선택 가능하게 됨
## 발주서: 신규 발주서에 발주일이 기존의 데이터와 겹치는 부분이 있으면 warning 메시지가 뜨고 발주서 정리 하지 않음!

root = Tk()
root.title("RENOVERA bot")
root.iconbitmap(r'icon.ico')
logo = ImageTk.PhotoImage(Image.open(r'logo.png'))
labl1 = Label(image = logo).grid(row = 0, column = 0, columnspan=3)

FOLDER_PATH = r'<Folder Path>'
# CLIENT_SECRET_FILE = 'client_secret.json'
CLIENT_SECRET_FILE = 'client_secret_share.json' #renov
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)


# spreadsheet_id = #################################
spreadsheet_id = '#################################' #renov

valueRanges_body = [
    '상품',
    '판매처',
    '판매가',
    'B2B',
    '행사'
]
response = service.spreadsheets().values().batchGet(
    spreadsheetId = spreadsheet_id,
    majorDimension = 'ROWS',
    ranges = valueRanges_body
).execute()

dataset = {}
for item in response['valueRanges']:
    dataset[item['range']] = item['values']

df = {}
for indx, k in enumerate(dataset):
    col = dataset[k][0]
    data = dataset[k]
    df[indx] = pd.DataFrame(data, columns = col )

for i in range(5):
    df[i].drop(df[i].index[0], inplace = True)
    df[i].reset_index(drop = True, inplace = True)



spreadsheet_id2 = ''#################################'' #renov
response = service.spreadsheets().values().get(
    spreadsheetId = spreadsheet_id2,
    majorDimension = 'ROWS',
    range = 'Sheet1'
).execute()
columns = response['values'][0]
data = response['values'][1:]

RENO = pd.DataFrame(data, columns = columns)



prd = df[0].copy()
shp = df[1].copy()
pri = df[2].copy()
B2B = df[3].copy()
evn = df[4].copy()



pri.set_index('판매처', inplace = True)
for colu in pri.columns:
    pri[colu] = pri[colu].str.replace(",","")
pri.fillna(0, inplace = True)
RENO["개수"] = RENO["개수"].astype(int)
RENO.fillna("", inplace=True)
B2B["개수"] = B2B["개수"].astype(int)
B2B.fillna("", inplace = True)
evn["가격"].fillna(0, inplace = True)
evn["가격"] = evn["가격"].astype(int)

RENO["발주일"] = pd.to_datetime(RENO["발주일"])
B2B["발주일"] = pd.to_datetime(B2B["발주일"])
evn["시작일"] = pd.to_datetime(evn["시작일"])
evn["마침일"] = pd.to_datetime(evn["마침일"])

### 판매처와 그 StringVar의 Dictionary
df1 = shp[["몰", "B2B_C"]].drop_duplicates().copy()
df1.reset_index(drop = True, inplace = True)
shops = dict()
shops["전체"] = IntVar()
i = 1
for shop in df1["몰"]:
    shops[shop.format(i)] = IntVar()
    i+=1
shops["그외"] = IntVar()


mall_B2C = []
mall_etc = []
for n in range(len(df1)):
    if df1["B2B_C"][n] == "C":
        mall_B2C.append(df1["몰"][n])
    elif df1["B2B_C"][n] == "etc":
        mall_etc.append(df1["몰"][n])

### 제품들과 그 StringVar의 Dictionary
df1 = prd["제품명"].copy()
prdcts = dict()
prdcts["전체"] = IntVar()
i = 1
for prdct in df1:
    prdcts[prdct.format(i)] = IntVar()
    i+=1

#### Functions
def RepurchaseSales(masked, table):
    df = masked.copy()
    df.drop(['메모', '판매처'], axis = 1, inplace = True)
    df = df.set_index(["수령자", "핸드폰번호"])
    df.dropna(inplace = True)
    df.reset_index(inplace = True)
    df = pd.pivot_table(df, values = "개수", index = ["주소", "수령자", "발주일"], aggfunc= np.sum)
    df.reset_index(inplace = True)

    # """Merging Addresses"""
    for i in range(len(df)-1):
        if df["수령자"][i+1] == df["수령자"][i]:
            if df["주소"][i+1] != df["주소"][i]:
                    df.loc[i+1,"주소"] = df["주소"][i]

    # """Making Repurchase Column"""
    df1 = df["주소"].copy()
    df2 = df1.groupby(df1.tolist()).size().reset_index()
    df2.columns = ['주소','Repurchase']

    list = []
    for i in range(len(df2)):
        for j in range(df2["Repurchase"][i]):
            list.append(df2["Repurchase"][i])

    df_Repurchase = pd.DataFrame(list, columns = ["Repurchase"])
    df = pd.concat([df, df_Repurchase], axis = 1)
    df.sort_values(by = ["주소", "발주일"], inplace = True)

    # """Only Repurchasing people"""
    mask1 = df["Repurchase"] > 1
    Repur = df[mask1].copy()
    Repur.reset_index(inplace = True, drop = True)

    # """Remove the First Purchase among the Re-Purhcasers"""
    inds = [0]
    for i in range(len(Repur)-1):
        if Repur["주소"][i] != Repur["주소"][i+1] :
            inds.append(i+1)
    Repur.drop(index = inds, inplace = True)
    Repur.reset_index(inplace = True, drop = True)

    # """Remove the purchases where 주소 = "" """
    inds2 = []
    for i in range(len(Repur)):
        if Repur["주소"][i] == "":
            inds2.append(i)
    Repur.drop(index=inds2, inplace=True)
    Repur.reset_index(inplace = True, drop = True)

    gropu = pd.pivot_table(Repur, values = "개수", index = "발주일", aggfunc = np.sum)

    Repur = table.join(gropu)["개수"].copy()
    Repur.fillna(0, inplace=True)

    return Repur

def graph_number(e1, e2, time_len):
    RENO_BB = pd.concat([RENO, B2B], axis = 0) #B2B 가 더해진 데이터
    RENO_BB.reset_index(drop = True, inplace = True)

    shp_lst = []
    if shops["전체"].get() == 1:
        table = pd.pivot_table(RENO_BB, values="개수", index="발주일", columns="상품명", aggfunc=np.sum, fill_value=0)
    else:
        for shop in shops:
            if shop != "그외":
                if shops[shop].get() != 0:
                    shp_lst.append(shop)
            else:
                if shops[shop].get() != 0:
                    for n in range(len(mall_etc)):
                        shp_lst.append(mall_etc[n])
        indx = []
        for shp_s in shp_lst:
            for i in range(len(RENO_BB)):
                if RENO_BB["판매처"][i] == shp_s:
                    indx.append(i)
        reno1 = RENO_BB.iloc[indx, :]
        table = pd.pivot_table(reno1, values="개수", index="발주일", columns="상품명", aggfunc=np.sum, fill_value=0)

    if time_len == "Weekly":
        table = table.resample("W").sum()
    elif time_len == "Monthly":
        table = table.resample("M").sum()
    elif time_len == "Quarterly":
        table = table.resample("3M").sum()


    prd_lst = []
    if prdcts["전체"].get() == 1:
        table2 = table.copy()
    else:
        for prdct in prdcts:
            if prdcts[prdct].get() != 0:
                prd_lst.append(prdct)
        table2 = table[prd_lst].copy()

    table2.loc[:, "Total"] = table2.sum(1)

    plt.plot(table2["Total"][e1:e2], label = str(shp_lst) + " " + str(prd_lst))
    plt.rc('font', family='Malgun Gothic')
    plt.legend()
    plt.show()

    return

def graph_sales(e1, e2, time_len):

    RENO1 = RENO.copy()
    B2B1 = B2B.copy()

    sale = []
    for i in range(len(B2B1)):
        try:
            sale.append(B2B1["개수"][i] * int(pri[B2B1["상품명"][i]][B2B1["판매처"][i]]))
        except:
            sale.append(0)
    B2B1["sale"] = sale
    
    po_set2 = []
    po_set3 = []
    so_set2 = []
    so_set3 = []
    single = []
    for i in range(len(RENO1)):
        if RENO1["상품명"][i] == "리노칼파 150g":
            if RENO1["개수"][i] % 3 == 0:
                po_set3.append(i)
            elif RENO1["개수"][i] % 2 == 0:
                po_set2.append(i)
            else:
                single.append(i)
        elif RENO1["상품명"][i] == "리노모팩 100g":
            if RENO1["개수"][i] % 3 == 0:
                so_set3.append(i)
            elif RENO1["개수"][i] % 2 == 0:
                so_set2.append(i)
            else:
                single.append(i)
        else:
            single.append(i)

    RENO1.loc[po_set2, "상품명_s"] = "리노칼파 150gx2"
    RENO1.loc[po_set3, "상품명_s"] = "리노칼파 150gx3"
    RENO1.loc[so_set2, "상품명_s"] = "리노모팩 100gx2"
    RENO1.loc[so_set3, "상품명_s"] = "리노모팩 100gx3"
    RENO1.loc[po_set2, "개수_s"] = RENO1["개수"] // 2
    RENO1.loc[po_set3, "개수_s"] = RENO1["개수"] // 3
    RENO1.loc[so_set2, "개수_s"] = RENO1["개수"] // 2
    RENO1.loc[so_set3, "개수_s"] = RENO1["개수"] // 3
    RENO1.loc[single, "개수_s"] = RENO1["개수"]
    RENO1.loc[single, "상품명_s"] = RENO1["상품명"]

    sale = []
    for i in range(len(RENO1)):
        try:
            sale.append(RENO1["개수_s"][i] * int(pri[RENO1["상품명_s"][i]][RENO1["판매처"][i]]))
        except:
            sale.append(0)

    RENO1["sale"] = sale

    for i in range(len(evn)):
        indxx = RENO1[ (RENO1["발주일"] >= evn["시작일"][i]) & (RENO1["발주일"] <= evn["마침일"][i]) & (RENO1["상품명"] == evn["제품"][i]) & (RENO1["판매처"] == evn["판매처"][i])].index
        RENO1.loc[indxx, "sale"] = RENO1["개수_s"] * evn["가격"][i]
    

    RENO_BB = pd.concat([RENO1, B2B1], axis = 0) #B2B 가 더해진 데이터
    RENO_BB.reset_index(drop = True, inplace = True)
    RENO_BB.fillna("", inplace=True)

    RENO_BB = RENO_BB[["발주일", "상품명", "sale", "판매처"]].copy()

    shp_lst = []
    if shops["전체"].get() == 1:
        RENO_BB = RENO_BB[["발주일", "상품명", "sale"]].copy()
        table = pd.pivot_table(RENO_BB, values="sale", index="발주일", columns="상품명", aggfunc=np.sum, fill_value=0)
    else:
        for shop in shops:
            if shop != "그외":
                if shops[shop].get() != 0:
                    shp_lst.append(shop)
            else:
                if shops[shop].get() != 0:
                    for n in range(len(mall_etc)):
                        shp_lst.append(mall_etc[n])
        indx = []
        for shp_s in shp_lst:
            for i in range(len(RENO_BB)):
                if RENO_BB["판매처"][i] == shp_s:
                    indx.append(i)
        reno1 = RENO_BB.iloc[indx, :].copy()
        table = pd.pivot_table(reno1, values="sale", index="발주일", columns="상품명", aggfunc=np.sum, fill_value= 0)
        

    if time_len == "Weekly":
        table = table.resample("W").sum()
    elif time_len == "Monthly":
        table = table.resample("M").sum()
    elif time_len == "Quarterly":
        table = table.resample("3M").sum()



    prd_lst = []
    if prdcts["전체"].get() == 1:
        table2 = table.copy()
    else:
        for prdct in prdcts:
            if prdcts[prdct].get() != 0:
                prd_lst.append(prdct)
        table2 = table[prd_lst].copy()
        
    table2.loc[:, "Total"] = table2.sum(1)

    plt.plot(table2["Total"][e1:e2], label = str(shp_lst) + " " + str(prd_lst))
    plt.rc('font', family='Malgun Gothic')
    plt.legend()
    plt.show()
    return

def xl(e1, e2, time_len):
    RENO_BB = pd.concat([RENO, B2B], axis = 0) #B2B 가 더해진 데이터
    RENO_BB.reset_index(drop = True, inplace = True)
    shp_lst = []
    if shops["전체"].get() == 1:
        table = pd.pivot_table(RENO_BB, values="개수", index="발주일", columns="상품명", aggfunc=np.sum, fill_value=0)
    else:
        for shop in shops:
            if shops[shop].get() != 0:
                shp_lst.append(shop)
        indx = []
        for shp_s in shp_lst:
            for i in range(len(RENO_BB)):
                if RENO_BB["판매처"][i] == shp_s:
                    indx.append(i)
        reno1 = RENO_BB.iloc[indx, :]
        table = pd.pivot_table(reno1, values="개수", index="발주일", columns="상품명", aggfunc=np.sum, fill_value=0)

    if time_len == "Weekly":
        table = table.resample("W").sum()
    elif time_len == "Monthly":
        table = table.resample("M").sum()
    elif time_len == "Quarterly":
        table = table.resample("3M").sum()


    prd_lst = []
    for prdct in prdcts:
        if prdcts[prdct].get() != 0:
            prd_lst.append(prdct)
    
    table2 = table[prd_lst][e1:e2].copy()

    file = askdirectory()
    if file:
        print(file)
        try:            
            table2.to_excel(file + "/" + e1 + "_" + e2 + ".xlsx")
            checkcheck = messagebox.showinfo("성공", file + " 파일에 저장되었습니다")         
        except:
            messagebox.showerror("Error", "해당 폴더에 저장 권한이 없습니다")
        return

def graph_Repur(cls, e1, e2, time_len):
    if cls == "리노베라 칼슘파우더":
        mask1 = RENO["상품명"] == prd["제품명"][0]
        d1 = RENO[mask1]

        mask2 = RENO["상품명"] == prd["제품명"][1]
        d2 = RENO[mask2]

        mask3 = RENO['상품명'] == prd["제품명"][2]
        d3 = RENO[mask3]

        mask4 = RENO["상품명"] == prd['제품명'][3]
        d4 = RENO[mask4]

        mask5 = RENO["상품명"] == prd['제품명'][4]
        d5 = RENO[mask5]

        d = pd.concat([d1,d2,d3,d4,d5], axis = 0)
    
    elif cls == "리노베라 모공팩바":
        mask1 = RENO["상품명"] == prd["제품명"][5]
        d1 = RENO[mask1]

        mask2 = RENO["상품명"] == prd["제품명"][6]
        d2 = RENO[mask2]

        d = pd.concat([d1,d2], axis = 0)
    
    elif cls == "그린그램":
        mask1 = RENO["상품명"] == prd["제품명"][7]
        d1 = RENO[mask1]

        mask2 = RENO["상품명"] == prd["제품명"][8]
        d2 = RENO[mask2]

        d = pd.concat([d1,d2], axis = 0)

    elif cls == "리노베라 살균탈취제":
        mask1 = RENO["상품명"] == prd["제품명"][9]

        d = RENO[mask1]

    d.reset_index(inplace=True, drop=True)
    table = pd.pivot_table(d, values = "개수", index = "발주일", columns="상품명",aggfunc=np.sum, fill_value=0)

    table.loc[:, "total"] = table.sum(1)

    table.loc[:, "repur"] = RepurchaseSales(d, table)


    if time_len == "Weekly":
        table = table.resample("W").sum()
    elif time_len == "Monthly":
        table = table.resample("M").sum()
    elif time_len == "Quarterly":
        table = table.resample("3M").sum()

    table.loc[:, "Repur%"] = np.where(table["total"]<1, table["total"], table["repur"]/table["total"])
    

    plt.plot(table["Repur%"][e1:e2])
    plt.show()

def showDel(three, e1):
    lbl_show = Label(three, text = len(RENO[RENO["발주일"] >= e1]))
    lbl_show.grid(row = 6, column = 1)

def ddelete(three, e1):
    response = messagebox.askokcancel("데이터 삭제하기", e1 + "날과 그 후의 모든 데이터를 삭제 하겠습니까?")
    if response == 1:
        global RENO
        RENO = RENO[RENO["발주일"] < e1].copy()
        Label(three, text = e1 + "일(포함) 이후의 데이터가 삭제 되었습니다").grid(row = 10, column = 1)
    else:
        Label(three, text = "데이터 삭제를 취소하셨습니다").grid(row = 10, column = 1)

def set150(e1, e2, time_len):
    REN = RENO[RENO["상품명"] == "리노칼파 150g"].copy()
    REN.reset_index(drop=True, inplace=True)

    po_set2 = []
    po_set3 = []
    single = []
    for i in range(len(REN)):
        if REN["개수"][i] % 3 == 0:
            po_set3.append(i)
        elif REN["개수"][i] % 2 == 0:
            po_set2.append(i)
        else:
            single.append(i)

    REN.loc[po_set2, "상품명_s"] = "리노칼파 150gx2"
    REN.loc[po_set3, "상품명_s"] = "리노칼파 150gx3"
    REN.loc[po_set2, "개수_s"] = REN["개수"] // 2
    REN.loc[po_set3, "개수_s"] = REN["개수"] // 3

    REN.loc[single, "상품명_s"] = REN["상품명"]
    REN.loc[single, "개수_s"] = REN["개수"]

    ### 넣기
    shp_lst = []
    if shops["전체"].get() == 1:
        df_hist = REN[(REN["발주일"] > e1) & (REN["발주일"] < e2)].copy()
        # table_his = pd.pivot_table(df_hist, index = "발주일",values = "개수_s", columns= "상품명_s", aggfunc = np.sum, fill_value = 0)
    else:
        for shop in shops:
            if shop != "그외":
                if shops[shop].get() != 0:
                    shp_lst.append(shop)
            else:
                if shops[shop].get() != 0:
                    for n in range(len(mall_etc)):
                        shp_lst.append(mall_etc[n])
        indx = []
        for shp_s in shp_lst:
            for i in range(len(REN)):
                if REN["판매처"][i] == shp_s:
                    indx.append(i)
        reno1 = REN.iloc[indx, :].copy()
        df_hist = reno1[(reno1["발주일"] > e1) & (reno1["발주일"] < e2)].copy()

    table_his = pd.pivot_table(df_hist, values="개수_s", index="발주일", columns="상품명_s", aggfunc=np.sum, fill_value=0)

    if time_len == "Weekly":
        table_his = table_his.resample("W").sum()
    elif time_len == "Monthly":
        table_his = table_his.resample("M").sum()
    elif time_len == "Quarterly":
        table_his = table_his.resample("3M").sum()

    table_his.loc[:, "단품 %"] = table_his['리노칼파 150g'] / (table_his['리노칼파 150gx2'] + table_his['리노칼파 150gx3'] + table_his['리노칼파 150g'])
    table_his.loc[:, "두개세트 %"] = table_his['리노칼파 150gx2'] / (table_his['리노칼파 150gx2'] + table_his['리노칼파 150gx3'] + table_his['리노칼파 150g'])
    table_his.loc[:, "세개세트 %"] = table_his['리노칼파 150gx3'] / (table_his['리노칼파 150gx2'] + table_his['리노칼파 150gx3'] + table_his['리노칼파 150g'])
    cross_tab_prop = table_his[["단품 %", "두개세트 %", "세개세트 %"]].copy()
    cross_tab = table_his[["리노칼파 150g", "리노칼파 150gx2", "리노칼파 150gx3"]].copy()

    # print(cross_tab)

    cross_tab_prop.plot(kind='bar', 
                        stacked=True, 
                        colormap='tab10', 
                        figsize=(10, 6))
    plt.rc('font', family='Malgun Gothic')
    plt.legend(loc="upper left", ncol=2)
    plt.xlabel("Release Year")
    plt.ylabel("Proportion")

    for n, x in enumerate([*cross_tab.index.values]):
        for (proportion, y_loc) in zip(cross_tab_prop.loc[x],
                                    cross_tab_prop.loc[x].cumsum()):
                    
            plt.text(x=n - 0.17,
                    y=(y_loc - proportion) + (proportion / 2),
                    s=f'{np.round(proportion * 100, 1)}%', 
                    color="black",
                    fontsize=12,
                    fontweight="bold")

    plt.show()

def dload(three, date_F):
    import glob
    global newDF
    files = glob.glob("*.xlsx")
    for file in files:
        if file.find("전체내역서") != -1:
                df = pd.read_excel(file)

    dfX = df.dropna(subset=['취소일'])
    df.drop(dfX.index, inplace=True)
    df.reset_index(inplace = True, drop = True)

    df["발주일"] = pd.to_datetime(df["발주일"])

    drp = ["송장번호", "주문번호", "공급처", "선택사항", "실제옵션", "원가", "주문일", "취소일", "발주시간", "송장입력일", "배송지우편번호", "배송일", "수령자전화", "관리번호", "선착불", "주문자"]
    df.drop(drp, axis = 1, inplace = True)
    df.fillna("", inplace = True)

    #### 겹치는 발주일이 있는지 확인
    df.sort_values(by = "발주일", inplace= True)
    date_N = df["발주일"].iloc[-1]
    if date_N <= date_F:
        messagebox.showwarning("경고!", "기존 데이터와 겹치는 발주서 입니다 \n발주서를 다시 확인해주세요")
        return



    df["판매처"] = df["판매처"].str.replace("에코-", "")
    list = ["재발송", "체험단", "직원구매", "사은품"]
    inds = []
    df['메모'].fillna("", inplace = True)
    for j in list:
        for i in range(len(df)):
            if df["메모"][i].find(j) != -1:
                inds.append(i)
            elif df["수령자"][i].find("남준호과장님") != -1:
                inds.append(i)
            elif df["수령자"][i].find("남준호 과장님") != -1:
                inds.append(i)

    df.drop(index=inds, inplace=True)
    df.reset_index(inplace = True, drop = True)

    for i in range(len(shp)):
        inds = []
        for j in range(len(df)):
            if df["판매처"][j].find(shp["발주서ID"][i]) != -1:
                inds.append(j)
        df.loc[inds, "판매처"] = shp["몰"][i]

    for i in range(len(shp)):
        inds = []
        for j in range(len(df)):
            if df["메모"][j].find(shp["발주서ID"][i]) != -1:
                inds.append(j)
        df.loc[inds, "판매처"] = shp["몰"][i]

    for i in range(len(prd)):
        inds = []
        for j in range(len(df)):
            if df["상품명"][j].find(prd["발주서ID"][i]) != -1:
                inds.append(j)
        df.loc[inds, "상품명"] = prd["제품명"][i]
    del_index = []
    for i in range(len(df)):
        cnt = 0
        inds = []
        for j in range(len(prd)):
            if df["상품명"][i].find(prd["제품명"][j]) != -1:
                cnt+=1
        if cnt == 0:
            del_index.append(i)
    df = df.drop(index=del_index) #그 외의 상품들, (쇼핑백 같은것)은 지운다.
    df.reset_index(inplace = True, drop = True)


    df = df.rename(columns={"판매개수":"개수"})
    df = df.rename(columns={"수령자핸드폰":"핸드폰번호"})
    df = df.rename(columns={"배송지주소":"주소"})

    df = pd.pivot_table(df, values = "개수", index = ["발주일", "주소", "수령자", "핸드폰번호", "판매처", "상품명", "메모"], aggfunc= np.sum)
    df.reset_index(inplace = True)
    df = df[["발주일", "판매처", "상품명", "개수", "메모", "수령자", "핸드폰번호","주소"]]
    df["발주일"] = pd.to_datetime(df["발주일"])
    
    newDF = df

    ######
    lbl_load1 = Label(three, text = df["발주일"].iloc[0])
    lbl_load2 = Label(three, text = df["발주일"].iloc[-1])
    lbl_txt = Label(three, text = "부터")
    lbl_dflen = Label(three, text = len(df))

    lbl_load1.grid(row = 4, column = 2)
    lbl_txt.grid(row = 6, column = 2)
    lbl_load2.grid(row = 8, column = 2)
    lbl_dflen.grid(row = 10, column = 2)


def dsave():
    file = askdirectory()
    if file:
        print(file)
        try:            
            newDF.to_excel(file + "/" + date.today().strftime("%Y-%m-%d") + ".xlsx")
            checkcheck = messagebox.showinfo("성공", file + " 파일에 저장되었습니다")         
        except:
            messagebox.showerror("Error", "해당 폴더에 저장 권한이 없습니다")
        return

###### Button One, Two, Three
def btn_sales():

    one = Toplevel() #anything you want in this window, you do it after this line
    one.title("판매량/매출")
    one.geometry("1100x700")

    ### Make a ScrollBar
    one_frame = Frame(one)
    one_frame.pack(fill = BOTH, expand=1)
    one_canvas = Canvas(one_frame)
    one_canvas.pack(side = LEFT, fill = BOTH, expand = 1)
    one_scrbar = ttk.Scrollbar(one_frame, orient=VERTICAL, command = one_canvas.yview)
    one_scrbar.pack(side = RIGHT, fill = Y)
    one_canvas.configure(yscrollcommand=one_scrbar.set)
    one_canvas.bind('<Configure>', lambda e: one_canvas.configure(scrollregion= one_canvas.bbox("all")))
    scnd_frame = Frame(one_canvas)
    one_canvas.create_window((0,0), window = scnd_frame, anchor="nw")

    ### 판매처 리스트 뽑기
    lbl_shp = Label(scnd_frame, text = "판매처")
    lbl_shp.grid(row = 0, column = 0, padx = 50)
    r = 1
    r2 = 2
    for shop in shops:
        if shop not in mall_etc:    
            c1 = Checkbutton(scnd_frame, text = shop, variable = shops[shop], onvalue = 1, offvalue = 0)
            c1.deselect()
            if shop in mall_B2C or shop == "전체":
                c1.grid(row = r, column= 0, sticky = 'w' )
                r +=1
            else:
                c1.grid(row = r2, column = 1, sticky = 'w')
                r2 +=1

    lbl_prd = Label(scnd_frame, text = "상품")
    lbl_prd.grid(row = 0, column = 2, padx = 50)
    r = 1
    for prdct in prdcts:
        c2 = Checkbutton(scnd_frame, text = prdct, variable = prdcts[prdct], onvalue = 1, offvalue = 0, padx = 50)
        c2.deselect()
        c2.grid(row = r, column = 2, sticky = 'w')
        r +=1

    lbl_dte = Label(scnd_frame, text = "기간")
    lbl_dte.grid(row = 0, column = 3, padx = 50)
    lbl_e1 = Label(scnd_frame, text = "시작일 입력: 연-월-일")
    lbl_e1.grid(row = 2, column = 3)
    lbl_e2 = Label(scnd_frame, text = "마침일 입력: 연-월-일")
    lbl_e2.grid(row = 4, column = 3)
    e1 = Entry(scnd_frame, width = 30)
    e1.grid(row = 3, column = 3)
    e1.insert(0, "2018-01-01")
    e2 = Entry(scnd_frame, width = 30)
    e2.grid(row = 5, column = 3)
    e2.insert(0, date.today().strftime("%Y-%m-%d"))


    lbl_len = Label(scnd_frame, text = "일/주/월/분기 단위")
    lbl_len.grid(row = 0, column = 4, padx = 50)

    option = ["Daily", "Weekly", "Monthly", "Quarterly"]
    time_len = StringVar()
    time_len.set(option[1])
    len = OptionMenu(scnd_frame, time_len, *option)
    len.grid(row = 1, column = 4)

    btn_num = Button(scnd_frame, text = "판매량 그래프", command = lambda: graph_number(e1.get(), e2.get(), time_len.get()))
    btn_num.config(height=10, width=20)
    btn_num.grid(row = 1, column = 5, padx = 50, rowspan = 6)

    btn_sal = Button(scnd_frame, text = "매출 그래프", command = lambda: graph_sales(e1.get(), e2.get(), time_len.get()))
    btn_sal.config(height = 10, width = 20)
    btn_sal.grid(row = 8, column = 5, padx = 50, rowspan = 6)

    btn_xl = Button(scnd_frame, text = "판매정보 엑셀 다운", command = lambda: xl(e1.get(), e2.get(), time_len.get()))
    btn_xl.config(height = 10, width = 20)
    btn_xl.grid(row = 15, column = 5, padx = 50, rowspan = 6)

    btn_set150 = Button(scnd_frame, text = "150g칼파 1,2,3개 \n 세트 판매 추이", command = lambda: set150(e1.get(), e2.get(), time_len.get()))
    btn_set150.config(width = 20)
    btn_set150.grid(row = 15, column = 4, padx = 50, rowspan = 2)

    lbl_set150 = Label(scnd_frame, text = "150g 리노칼파만 계산! \n판매처만 반영 됩니다.")
    lbl_set150.grid(row = 17, column = 4)

def btn_repr():
    two = Toplevel()
    two.geometry("1000x300")
    two.title("재구매율")

    lbl_class = Label(two, text = "제품군")
    lbl_class.grid(row = 0, column = 0, padx = 50)
    option1 = ["리노베라 칼슘파우더", "리노베라 모공팩바", "그린그램", "리노베라 살균탈취제"]
    cls = StringVar()
    cls.set(option1[0])
    clas = OptionMenu(two, cls, *option1)
    clas.grid(row = 3, column = 0, padx = 50)


    lbl_dte = Label(two, text = "기간")
    lbl_dte.grid(row = 0, column = 1)
    lbl_e1 = Label(two, text = "시작일 입력: 연-월-일")
    lbl_e1.grid(row = 2, column = 1)
    lbl_e2 = Label(two, text = "마침일 입력: 연-월-일")
    lbl_e2.grid(row = 4, column = 1)
    e1 = Entry(two, width = 30)
    e1.grid(row = 3, column = 1)
    e1.insert(0, "2019-01-01")
    e2 = Entry(two, width = 30)
    e2.grid(row = 5, column = 1)
    e2.insert(0, date.today().strftime("%Y-%m-%d"))

    lbl_len = Label(two, text = "일/주/월/분기 단위")
    lbl_len.grid(row = 0, column = 2, padx = 50)
    option = ["Daily", "Weekly", "Monthly", "Quarterly"]
    time_len = StringVar()
    time_len.set(option[2])
    len = OptionMenu(two, time_len, *option)
    len.grid(row = 3, column = 2)

    btn_num = Button(two, text = "그래프 그리기", command = lambda: graph_Repur(cls.get(),e1.get(), e2.get(), time_len.get()))
    btn_num.config(height=10, width=20)
    btn_num.grid(row = 3, column = 4, padx = 50, rowspan=10)

def btn_updt():
    three = Toplevel()
    three.title("데이터 업데이트")
    three.geometry("600x300")

    dates1 = RENO["발주일"].copy()
    dates1.sort_values(inplace = True)

    lbl_whatdata = Label(three, text = "현재 데이터베이스의 데이터")
    lbl_whatdata.grid(row = 0, column = 0)

    lbl_showdata = Label(three, text = dates1.iloc[0])
    lbl_showdata2 = Label(three, text = dates1.iloc[-1])
    lbl_text = Label(three, text = "부터")
    lbl_dbleng = Label(three, text = len(RENO))
    lbl_showdata.grid(row = 2, column = 0)
    lbl_text.grid(row = 4, column = 0)
    lbl_showdata2.grid(row = 6, column = 0)
    lbl_dbleng.grid(row = 8, column = 0)


    # lbl_e1 = Label(three, text = "지울 데이터 선택: 연-월-일")
    # lbl_e1.grid(row = 0, column = 1)
    # e1 = Entry(three, width = 30)
    # e1.grid(row = 2, column = 1)
    # e1.insert(0, date.today().strftime("%Y-%m-%d"))

    # btn_showDel = Button(three, text = "지워질 데이터 정보 보기", command = lambda: showDel(three, e1.get()))
    # btn_showDel.grid(row = 4, column = 1)

    # btn_Del = Button(three, text = "선택 데이터 지우기", command = lambda: ddelete(three, e1.get()))
    # btn_Del.grid(row = 8, column = 1)

    # global newDF
    # newDF = pd.DataFrame()
    lbl_add = Label(three, text = "베이스에 추가할 발주서")
    lbl_add.grid(row = 0, column = 2)
    btn_load = Button(three, text = "신규 발주서 불러오기", command = lambda: dload(three, dates1.iloc[-1]))
    btn_load.grid(row=2, column = 2)

    # btn_add = Button(three, text = "데이터 추가하기", command = lambda: dadd(newDF))
    # btn_add.grid(row = 12, column = 2)

    btn_dsave = Button(three, text = "신규 발주서 엑셀 저장하기", command = dsave)
    btn_dsave.grid(row = 14, column = 1, ipadx = 30, ipady = 20, pady = 30, columnspan = 2, rowspan = 2)


Button(root, text = "판매량/매출", command = btn_sales).grid(row = 1, column = 0)
Button(root, text = "재구매율", command = btn_repr).grid(row = 1, column= 1)
Button(root, text = "데이터 업데이트", command = btn_updt).grid(row = 1, column= 2)


root.mainloop()