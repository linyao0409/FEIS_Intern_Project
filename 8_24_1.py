import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont
import pandas as pd

file_path = "ISIN_code_test.xlsx"  # 記得替換成你的Excel檔案路徑
df = pd.read_excel(file_path, engine="openpyxl")




# 建立主視窗 -------------------------------------------------------------------------------
root = tk.Tk()
root.geometry("1270x720+400+200")
root.config(bg="gray")
root.title("證券經濟部報價平台")
# 設定視窗的icon--------------------------------------------------------------------------------
root.iconbitmap('feis_logo.ico')  # 將 'path/to/your/icon.ico' 替換為您的圖標檔案的路徑


# 建立分頁
notebook = ttk.Notebook(root)
# 第一個分頁 ################################################################################################################################################################################
#text_p1_response = None
################################
page1 = ttk.Frame(notebook)
notebook.add(page1, text='更新及刪除ISIN-code')
# 自定義第一個分頁的風格
style = ttk.Style()
style.configure("Page1.TFrame", background="white")  #其他分頁沿用
page1.configure(style="Page1.TFrame")

# 第一個分頁整理ISIN函式################################################################################################################################################################################
def load_excel_file():
    file_path = "ISIN_code_test.xlsx"  # 記得替換成你的Excel檔案路徑
    df = pd.read_excel(file_path, engine="openpyxl")

    # 建立新的視窗顯示股票代碼
    display_window = tk.Toplevel(page1)
    display_window.title("股票代碼一覽")
    display_window.geometry("200x800")

    # 使用自定義風格設定 Treeview 的字型大小
    style = ttk.Style()
    style.configure("Treeview", font=("微軟正黑體", 13))  # 調整這裡的字型大小

    # 使用Treeview元件顯示股票代碼
    tree = ttk.Treeview(display_window, columns=("ISIN"), show="headings")
    tree.heading("ISIN", text="ISIN --- code")
    tree.pack(fill="both", expand=True)
    i = 1
    for item in df["ISIN"]:
        stock_code = str(i)+" -- "+item
        i+=1
        tree.insert("", "end", values=(stock_code,))


def delete_isin(index):
    global text_p1_response
    global df
    global update_button
    global delete_button
    df = df.drop(index) # 去掉搜尋ISIN那一列
    # 設定輸出Excel檔案的路徑和檔名
    output_file = 'ISIN_code_test.xlsx' # 覆蓋掉原先ISIN_code_test.clsx檔案
    df.to_excel(output_file, index=False)# 將DataFrame輸出成Excel檔案
    text_p1_response.config(text="已刪除ISIN",fg="red")
    update_button.config(state="disabled")  # 禁用更新按鈕
    delete_button.config(state="disabled")  # 禁用刪除按鈕

def update_company(index,esun_var,dbs_var,cathay_var,sunny_var):
    global text_p1_response
    global df
    global update_button
    global delete_button
    df.iloc[index,1:] = [int(esun_var.get()),int(dbs_var.get()),int(cathay_var.get()),int(sunny_var.get())]

    output_file = 'ISIN_code_test.xlsx' # 覆蓋掉原先ISIN_code_test.clsx檔案
    df.to_excel(output_file, index=False)# 將DataFrame輸出成Excel檔案
    text_p1_response.config(text="已更新窗口")
    update_button.config(state="disabled")  # 禁用更新按鈕
    delete_button.config(state="disabled")  # 禁用刪除按鈕

def insert_new_isin(isin,esun_var,dbs_var,cathay_var,sunny_var):
    global df
    global text_p1_response
    global add_button
    # 新的資料
    new_data = {
        'ISIN': isin, 
        'ESUN': [int(esun_var.get())],
        'DBS': [int(dbs_var.get())],
        'CATHAY': [int(cathay_var.get())],
        'SUNNY': [int(sunny_var.get())],
    }

    # 新增新的列
    new_row_df = pd.DataFrame(new_data)
    # 合併原始DataFrame和新的DataFrame
    df = pd.concat([df, new_row_df], ignore_index=True)
    output_file = 'ISIN_code_test.xlsx' # 覆蓋掉原先ISIN_code_test.clsx檔案
    df.to_excel(output_file, index=False)# 將DataFrame輸出成Excel檔案
    text_p1_response.config(text="已新增此ISIN並更新")
    add_button.config(state="disabled")  # 禁用新增按鈕
    
def search_button_clicked():
    global text_p1_response
    global add_button
    global update_button
    global delete_button

    file_path = "ISIN_code_test.xlsx"  # Excel path
    df = pd.read_excel(file_path, engine="openpyxl")

    search_input = input_entry.get()  # Get the user input from the Entry widget
    new_window = tk.Toplevel(root)
    new_window.title("刪除更新ISIN/新增ISIN")
    new_window.geometry("800x600")

    # 自訂checkbox字型大小和樣式
    checkbox_font = tkfont.Font(family="微軟正黑體", size=16)
    
      

    if search_input in df["ISIN"].values: # isin exist already
        text_isin_exist = tk.Label(new_window,text="ISIN-code已存在",font=('微軟正黑體',18))
        text_isin_exist.pack()

        # Get the row corresponding to the searched ISIN
        row = df[df["ISIN"] == search_input].iloc[0]
        #index = list(df["ISIN"] == search_input).index(True)
        index = df[df["ISIN"]==search_input].index[0]
        
        global esun_var 
        esun_var = tk.BooleanVar()
        esun_var.set(int(row["ESUN"]==1))
        esun_cb = tk.Checkbutton(new_window,text="玉山銀行",variable=esun_var,font=checkbox_font)
        esun_cb.pack()

        global dbs_var 
        dbs_var = tk.BooleanVar()
        dbs_var.set(int(row["DBS"]==1))
        dbs_cb = tk.Checkbutton(new_window,text ="星展銀行",variable=dbs_var,font=checkbox_font)
        dbs_cb.pack()

        global cathay_var 
        cathay_var = tk.BooleanVar()
        cathay_var.set(int(row["CATHAY"]==1))
        cathay_cb = tk.Checkbutton(new_window,text="國泰世華",variable=cathay_var,font=checkbox_font)
        cathay_cb.pack()

        global sunny_var 
        sunny_var = tk.BooleanVar()
        sunny_var.set(int(row["SUNNY"]==1))
        sunny_cb = tk.Checkbutton(new_window,text="陽信銀行",variable=sunny_var,font=checkbox_font)
        sunny_cb.pack()

        # Add a "Delete ISIN" button
        delete_button = tk.Button(new_window, text="刪除此ISIN-code", font=("微軟正黑體", 16), command=lambda: delete_isin(index))

        #delete_button = tk.Button(new_window, text="刪除此ISIN-code", font=("微軟正黑體", 16),command=delete_isin(index))
        delete_button.pack()

        #add a "update company" button
        update_button = tk.Button(new_window, text="更新ISIN對應窗口", font=("微軟正黑體", 16),command=lambda: update_company(index,esun_var,dbs_var,cathay_var,sunny_var))
        update_button.pack()

    else:
        # BOOL is false, show a new window with checkboxes and an "Add ISIN" button
        text_isin_not_exist = tk.Label(new_window,text="ISIN-code目前不存在",bg="dark red",font=('微軟正黑體',18))
        text_isin_not_exist.pack()

        esun_var = tk.BooleanVar()
        dbs_var = tk.BooleanVar()
        cathay_var = tk.BooleanVar()
        sunny_var = tk.BooleanVar()

        eusan_checkbox = tk.Checkbutton(new_window, text='玉山銀行', variable=esun_var, font=("微軟正黑體", 16))
        dbs_checkbox = tk.Checkbutton(new_window, text='星展銀行', variable=dbs_var, font=("微軟正黑體", 16))
        catchy_checkbox = tk.Checkbutton(new_window, text='國泰世華', variable=cathay_var, font=("微軟正黑體", 16))
        sunny_checkbox = tk.Checkbutton(new_window, text='陽信銀行', variable=sunny_var, font=("微軟正黑體", 16))

        eusan_checkbox.pack()
        dbs_checkbox.pack()
        catchy_checkbox.pack()
        sunny_checkbox.pack()

        # Add an "Add ISIN" button
        #def insert_new_isin(isin,esun_var,dbs_var,cathay_var,sunny_var):
        add_button = tk.Button(new_window, text="新增ISIN-code", font=("微軟正黑體", 16),command= lambda: insert_new_isin(search_input,esun_var,dbs_var,cathay_var,sunny_var))
        add_button.pack() # BOOL is true, show a new window with checkboxes and a "Delete ISIN" button
       
    text_p1_response = tk.Label(new_window,text="",font=('微軟正黑體',18))
    text_p1_response.pack()

# 搜尋框和搜尋按鈕
input_entry = tk.Entry(page1,bg="white",font=("微軟正黑體",18))
search_button = tk.Button(page1, text='搜尋',fg='black',bg="grey",font=('微軟正黑體',18), command=search_button_clicked)
#
text_p1 = tk.Label(page1,text="請輸入ISIN-code",bg="white",font=('微軟正黑體',18))
observe_isin_button = tk.Button(page1, text="查看ISIN_CODE",font=('微軟正黑體',18) ,command=load_excel_file)
#
text_p1.grid(row=0,column=0,columnspan=5,padx = 50)
input_entry.grid(row=0, column=6, padx=10, pady=10)  
search_button.grid(row=0, column=7, padx=10, pady=10)
observe_isin_button.grid(row=0,column=8,padx=10,pady=10)

######################################################################################################
#                           page2 - > creating table
######################################################################################################
#import time
from p2 import process_data_temp # ~~~~
def page2_wraper():
    global text_p2_response
    text_p2_response.config(text="製作中...",fg="red")
    root.update()  # 更新界面，以便立即显示“製作中”状态
    ##time.sleep(1)  # 等待一秒
    process_data_temp()# ~~~~
    text_p2_response.config(text="前置報表製作完成!",fg="Black")
#--------------------------------------------------------------------------
page2 = ttk.Frame(notebook)
notebook.add(page2, text='製作先備報表')
# 自定義第2個分頁的風格
page2.configure(style="Page1.TFrame")

# 製作報表按鈕
report_button = tk.Button(page2, text='製作報表' )
report_button.config(font=("微軟正黑體", 16), foreground="black", background="grey",command = page2_wraper)  # 設定按鈕字型和顏色
report_button.pack(padx=30, pady=30)
#Response
text_p2_response = tk.Label(page2,text="尚未開始製做前置報表",fg="Black",font=("微軟正黑體",18))
text_p2_response.pack()
######################################################################################################
#                          page3 - >第三個分頁: select bank
######################################################################################################
page3 = ttk.Frame(notebook)
notebook.add(page3, text='選擇報價銀行')
page3.configure(style="Page1.TFrame")
########################################################
# 分頁三函式
#import linyaoTable from p3
def create_company_table(esun,dbs,cathay,sunny,linyao):
    if esun:
        pass
    if dbs:
        pass
    if cathay:
        pass
    if sunny:
        pass
    if linyao:
        #linyaoTable()
        pass
########################################################

# 自訂字型大小和樣式
checkbox_font = tkfont.Font(family="微軟正黑體", size=18)

all_var_p3 = tk.BooleanVar() # all bank button
esun_var_p3 = tk.BooleanVar()
dbs_var_p3 = tk.BooleanVar()
cathay_var_p3 = tk.BooleanVar()
sunny_var_p3 = tk.BooleanVar()
linyao_var_p3 = tk.BooleanVar() #test


bank_info_p3 = [
    ('全選', all_var_p3),
    ('玉山銀行', esun_var_p3),
    ('星展銀行', dbs_var_p3),
    ('國泰世華', cathay_var_p3),
    ('陽信銀行', sunny_var_p3),
    ('林曜銀行', linyao_var_p3)
]

# 創建並放置銀行選擇的複選框
checkboxes_p3 = []  # 用來存放複選框
for bank_name, bank_variable in bank_info_p3:
    checkbox = tk.Checkbutton(page3, text=bank_name, variable=bank_variable, font=checkbox_font)
    checkbox.pack()
    checkboxes_p3.append(checkbox)

# 選擇所有複選框的函式
def select_all():
    for checkbox in checkboxes_p3[1:]:  # 跳過全選複選框
        checkbox.select() if all_var_p3.get() else checkbox.deselect()

checkboxes_p3[0].config(command=select_all)




generate_bank_report_button = tk.Button(page3, text='製作勾選公司報表', command=lambda: create_company_table(esun_var_p3.get(), dbs_var_p3.get(), cathay_var_p3.get(), sunny_var_p3.get(),linyao_var_p3.get()))
generate_bank_report_button.config(font=("微軟正黑體", 16), foreground="black", background="grey")  # 設定按鈕字型和顏色
generate_bank_report_button.pack()

######################################################################################################
#                          page4 - >第四個分頁: mail table 
######################################################################################################
import page4_1 # 引入第四分頁寄信函式
def mail_table_to_bank():
    L = [
        ('玉山銀行', esun_var_p3.get()),
        ('星展銀行', dbs_var_p3.get()),
        ('國泰世華', cathay_var_p3.get()),
        ('陽信銀行', sunny_var_p3.get()),
        ('林曜銀行', linyao_var_p3.get())
    ]
    for bank,index in L:
        if index:
            page4_1.sendmail(bank)
    print(f"done")
     
#---------------------------------------------------------

page4 = ttk.Frame(notebook)
notebook.add(page4, text='郵寄至銀行窗口')
page4.configure(style="Page1.TFrame")

bank_info_p4 = [
    ('玉山銀行', esun_var_p3),
    ('星展銀行', dbs_var_p3),
    ('國泰世華', cathay_var_p3),
    ('陽信銀行', sunny_var_p3),
    ('林曜銀行', linyao_var_p3)
]

checkboxes_p4 = []  # 用來存放複選框
for bank_name, bank_variable in bank_info_p4:
    checkbox = tk.Checkbutton(page4, text=bank_name, variable=bank_variable, font=checkbox_font)
    checkbox.pack()
    checkboxes_p4.append(checkbox)


send_report_button = tk.Button(page4, text='寄出報表')
send_report_button.config(font=("微軟正黑體", 16), foreground="black", background="grey",command = mail_table_to_bank)  # 設定按鈕字型和顏色
send_report_button.pack()

######################################################################################################
#                                    page5 - > 第五個分頁:edit mail address
######################################################################################################
#Banks_Chinese_en = {"玉山":'玉山銀行',"星展":'DBS',"國泰":'國泰世華',"陽信":'陽信銀行',"林曜":"林曜銀行"}
# test_mail.xlsx  {bank:bank name in mail file}
page5 = ttk.Frame(notebook)
notebook.add(page5, text='編輯銀行郵寄窗口')
page5.configure(style="Page1.TFrame")

def quick_ckeck_5(bank_name):
    global mail_df
    file_path = "test_mail.xlsx"  # Excel檔案路徑
    mail_df = pd.read_excel(file_path, engine="openpyxl")

    # 建立新的視窗顯示EMAIL
    display_window_5 = tk.Toplevel(page5)
    display_window_5.title("快速查看"+bank_name+"信箱")
    display_window_5.geometry("450x300")

    # 使用自定義風格設定 Treeview 的字型大小
    style = ttk.Style()
    style.configure("Treeview", font=("微軟正黑體", 13))  # 調整這裡的字型大小

    # 使用Treeview元件顯示EMAIL欄位
    tree = ttk.Treeview(display_window_5, columns=("EMAIL",), show="headings", style="Treeview")
    tree.heading("EMAIL", text="EMAIL")  # 設置EMAIL欄位的標題
    tree.pack(fill="both", expand=True)

    for item in mail_df[bank_name].values:
        if pd.notna(item):  # 判斷是否不是 NaN
            mail = item
            tree.insert("", "end", values=(mail,))

def add_email(mail,name,bank_str):
    global mail_df
    global text_p5_response
    file_path = "test_mail.xlsx"  # Excel檔案路徑
    mail_df = pd.read_excel(file_path, engine="openpyxl")
    value = name+";"+mail
    bank_series = mail_df[bank_str]
    # 找到第一個空值的索引位置
    empty_indices = bank_series[bank_series.isnull()].index
    
    if len(empty_indices) > 0:
        # 如果有空值，則更新第一個空值的位置
        empty_index = empty_indices[0]
        mail_df.at[empty_index, bank_str] = value
    else:
        # 如果沒有空值，則在最下方新增一行
        new_index = mail_df.index.max() + 1
        mail_df.loc[new_index, bank_str] = value
    mail_df.to_excel(file_path,index=False,engine="openpyxl")
    text_p5_response.config(text="已新增 "+name+": "+mail+" 並更新")
    
    #pass

def delete_email(mail,name,bank_str):
    global mail_df
    global text_p5_response
    file_path = "test_mail.xlsx"  # Excel檔案路徑
    mail_df = pd.read_excel(file_path, engine="openpyxl")

    mail_series = mail_df[bank_str]
    updated_mail_series = pd.Series()
    bool_index = False
    for i, x in enumerate(mail_series):
        if isinstance(x, str):
            n, e = parse_input(x)
            if e != mail and n != name:
                updated_mail_series.at[i] = x  # 將非 "a@gmail.com" 的值添加到新的Series
                bool_index = True
        #else:
            ##print("NO mail")
            #break
    #text_p5_response.after(3000, clear_text_response)

    updated_mail_series.index = range(len(updated_mail_series)) # 更新update_mail_series的index
    mail_df[bank_str] = updated_mail_series
    mail_df.to_excel(file_path,index=False,engine="openpyxl")
    if bool_index:
        text_p5_response.config(text="已刪除 "+name+": "+mail+" 並更新",fg="red")
    else:
        text_p5_response.config(text="不存在此信箱和名稱")
    text_p5_response.after(3000, clear_text_response)
        

def bank_button_clicked(bank_str):
    global mail_df
    # 重新讀入MAIL EXCEL檔案
    file_path = "test_mail.xlsx"  # Excel檔案路徑
    mail_df = pd.read_excel(file_path, engine="openpyxl")
    bank_str 
    #跳出銀行窗口 
    bank_window = tk.Toplevel(page5)
    bank_window.title(bank_str + " 郵件地址新增更改")
    bank_window.geometry("500x400")

    #快速查看
    quick_check_button = tk.Button(bank_window,text="快速查看EMAIL",fg="Black",bg="grey",font=("微軟正黑體",18),command= lambda b = bank_str: quick_ckeck_5(b))
    quick_check_button.pack()

    #請輸入EMAIL(文字))
    text_p5_1 = tk.Label(bank_window,text = "請輸入EMAIL",font=('微軟正黑體',18))
    text_p5_1.pack()
    #新增email輸入框
    entry_p5_mail = tk.Entry(bank_window,bg="white",font=("微軟正黑體",18))
    entry_p5_mail.pack()
    #請輸入名稱(文字))
    text_p5_2 = tk.Label(bank_window,text = "請輸入名稱",font=('微軟正黑體',18))
    text_p5_2.pack()
    #新增名稱輸入框
    entry_p5_name = tk.Entry(bank_window,bg="white",font=("微軟正黑體",18))
    entry_p5_name.pack()

    def add_email_button_wraper():
        entry_mail = entry_p5_mail.get()
        entry_name = entry_p5_name.get()
        add_email(entry_mail,entry_name,bank_str)
    def delete_email_button_wraper():
        entry_mail = entry_p5_mail.get()
        entry_name = entry_p5_name.get()
        delete_email(entry_mail,entry_name,bank_str)
    add_email_button = tk.Button(bank_window,text="新增EMAIL,更新人名",fg="Black",bg="grey",font=("微軟正黑體",18),command=add_email_button_wraper,width=17)
    add_email_button.pack()
    delete_emil_button = tk.Button(bank_window,text="刪除EMAIL和人名",fg="Black",bg="grey",font=("微軟正黑體",18),command=delete_email_button_wraper,width=17)
    delete_emil_button.pack()


def parse_input(input_str):
    # 尋找分號的位置
    semicolon_index = input_str.find(';')
    # 分割NAME和EMAIL
    name = input_str[:semicolon_index].strip()
    email = input_str[semicolon_index + 1:].strip()
    return name, email
def clear_text_response():
    text_p5_response.config(text="")
#________________________________________________________________

bank_buttons_frame = tk.Frame(page5)  # 建立一個 Frame 放置按鈕
bank_buttons_frame.pack(pady=20)
file_path = "test_mail.xlsx"  # 記得替換成你的Excel檔案路徑
mail_df = pd.read_excel(file_path, engine="openpyxl")

banks = ["玉山銀行", "星展銀行", "國泰世華", "陽信銀行","林曜銀行"]
for bank in banks:
    bank_button = tk.Button(bank_buttons_frame, text=bank, command=lambda b=bank: bank_button_clicked(b))
    bank_button.config(font=("微軟正黑體", 16), foreground="black", background="grey")  # 設定按鈕字型和顏色
    bank_button.pack(side=tk.LEFT, padx=10)
text_p5_response = tk.Label(page5,text="",fg="Black",font=("微軟正黑體",18))
text_p5_response.pack()

######################################################################################################
#                                    page6 - > raning table adjustment
######################################################################################################
def insert_ranking(ranking_str, ranking_degree,institute):
    d = {"S&P":0,"Moody's":1,"Fitch":2}
    index = d[institute]
    file_path = "rating.xlsx"  # 記得替換成你的Excel檔案路徑
    ranking_df = pd.read_excel(file_path, engine="openpyxl")
    if ranking_str in ranking_df[institute].values:
        # 評級已經存在
        row_location = (ranking_df[institute] == ranking_str).idxmax()#row_index
        temp_d_value = ranking_df.iloc[row_location,index] # currently degree
        if ranking_degree == temp_d_value: # 與現值相等 ->無須修改
            return 
        else: # 與現值不同 -> 把現值更新為空 ，新增一橫列 ex:[0,1,0,degree]
            ranking_df.iloc[row_location,index] = ""

            new_row = ["" for i in range(4)]
            new_row[index] = ranking_str
            new_row[3] = ranking_degree
            ranking_df.loc[len(ranking_df)] = new_row
            ranking_df.to_excel(file_path,index=False,engine="openpyxl")
            text_p6_response.config(text=institute+"已更新: "+ranking_str+"評級")
    else:
        #評級不存在 ->新增評級
        new_row = ["" for i in range(4)] #　empty row
        new_row[index] = ranking_str
        new_row[3] = ranking_degree
        ranking_df.loc[len(ranking_df)] = new_row
        ranking_df.to_excel(file_path,index=False,engine="openpyxl")
        text_p6_response.config(text=institute+"已新增: "+ranking_str+"評級")
    text_p6_response.after(3000, clear_text_response)
    
def delete_ranking(ranking_str,institute):
    global text_p6_response  # 確保 text_p6_response 是全域變數
    d = {"S&P":0,"Moody's":1,"Fitch":2}
    index = d[institute]
    file_path = "rating.xlsx"  # 記得替換成你的Excel檔案路徑
    ranking_df = pd.read_excel(file_path, engine="openpyxl")
    if ranking_str in ranking_df[institute].values:
        row_location = (ranking_df[institute] == ranking_str).idxmax()
        ranking_df.iloc[row_location,index] = ""
        ranking_df.to_excel(file_path,index=False,engine="openpyxl")
        text_p6_response.config(text=institute+"已刪除: "+ranking_str+"評級",fg="red")
    else:
        text_p6_response.config(text=institute+"不存在: "+ranking_str+"評級",fg="red")
    text_p6_response.after(3000, clear_text_response)

def quick_check_6(institute):
    global ranking_df
    file_path = "rating.xlsx"  # 記得替換成你的Excel檔案路徑
    ranking_df = pd.read_excel(file_path, engine="openpyxl")

    style = ttk.Style()
    style.configure("Treeview", font=("微軟正黑體", 13))  # 調整這裡的字型大小
    # 建立新的視窗顯示EMAIL
    display_window = tk.Toplevel(page6)
    display_window.title("快速查看"+institute+"信箱")

    # 使用Treeview元件顯示EMAIL欄位
    tree = ttk.Treeview(display_window, columns=("Ranking",), show="headings")
    tree.heading("Ranking", text="Ranking")  # 設置EMAIL欄位的標題
    tree.pack(fill="both", expand=True)

    for r,d in ranking_df[[institute,"degree"]].values:
        if pd.notna(r):  # 判斷是否不是 NaN
            ranking = f"{r:<10} {str(d):<10}"
            tree.insert("", "end", values=(ranking,))


def clear_text_response():
    text_p6_response.config(text="")
def ranking_buttton_clicked(institute):
    #file_path = "rating.xlsx"  # 記得替換成你的Excel檔案路徑
    #raiting_df = pd.read_excel(file_path, engine="openpyxl")
    institute_window = tk.Toplevel(page6)
    institute_window.title(institute + " 評級新增，修改，刪除")
    institute_window.geometry("500x400")
    quick_check_button_6 = tk.Button(institute_window, text="快速查看"+institute +"評級", fg="black", bg="grey", font=("微軟正黑體", 18), command=lambda: quick_check_6(institute))
    quick_check_button_6.pack()

    text_p6_window = tk.Label(institute_window,text = "請輸入評級",font=('微軟正黑體',18))
    entry_p6_window = tk.Entry(institute_window,bg="white",font=("微軟正黑體",18))
    text_p6_window.pack()
    entry_p6_window.pack()

    text_p6_window_degree = tk.Label(institute_window,text = "請輸入degree",font=('微軟正黑體',18))
    entry_p6_window_degree = tk.Entry(institute_window,bg="white",font=("微軟正黑體",18))
    text_p6_window_degree.pack()
    entry_p6_window_degree.pack()

    def insert_ranking_wrapper():
        #global text_p6_response  # 確保 text_p6_response 是全域變數
        ranking_str = entry_p6_window.get()
        ranking_degree = entry_p6_window_degree.get()
        insert_ranking(ranking_str, ranking_degree,institute)
    def delete_ranking_wrapper():
        #global text_p6_response  # 確保 text_p6_response 是全域變數
        ranking_str = entry_p6_window.get()
        #ranking_degree = entry_p6_window_degree.get()
        delete_ranking(ranking_str,institute)

    insert_button = tk.Button(institute_window, text="新增,更新評級degree", fg="black", bg="grey", font=("微軟正黑體", 18), command=insert_ranking_wrapper,width=20)
    insert_button.pack(pady=10)

    delete_button = tk.Button(institute_window,text = "刪除評級",fg="Black",bg="grey",font=("微軟正黑體",18),command = delete_ranking_wrapper,width=20)
    delete_button.pack(pady=10)

page6 = ttk.Frame(notebook)
notebook.add(page6, text='調整評級')
page6.configure(style="Page1.TFrame")

ranking_institutes = ["S&P", "Moody's", "Fitch"]

# 設定列的權重，使按鈕置中
for i in range(3):
    page6.columnconfigure(i, weight=1)

for i, institute in enumerate(ranking_institutes):
    ranking_button = tk.Button(page6, text=institute, command=lambda c=institute: ranking_buttton_clicked(c))
    ranking_button.config(font=("微軟正黑體", 16), foreground="black", background="grey")  # 設定按鈕字型和顏色
    ranking_button.grid(row=0, column=i, padx=10, pady=10, sticky="nsew")

#Response text
text_p6_response = tk.Label(page6,text="",fg="Black",font=("微軟正黑體",18))
text_p6_response.grid(row=1, columnspan=3, pady=10)

##########################################################################################################
#                                    page7 -> adjust currency
##########################################################################################################
def quick_check_currency_fn():
    global currency_df
    file_path = "exchange_adjust.xlsx"  # 記得替換成你的Excel檔案路徑
    currency_df = pd.read_excel(file_path, engine="openpyxl")

    currency_window = tk.Toplevel(page7)
    currency_window.title("快速查看Currency")
    # 使用自定義風格設定 Treeview 的字型大小
    style = ttk.Style()
    style.configure("Treeview", font=("微軟正黑體", 15))  # 調整這裡的字型大小
    currency_window.geometry("200x400")
    # 使用Treeview元件顯示EMAIL欄位
    tree = ttk.Treeview(currency_window, columns=("幣別",), show="headings")
    tree.heading("幣別", text="幣別")  # 設置EMAIL欄位的標題
    tree.pack(fill="both", expand=True)
    for c, b, a in currency_df[["判斷B", "Bid", "Ask"]].values:
        item = f"{c:<8} {b:<8} {a:<8}"
        tree.insert("", "end", values=(item,))



def update_currency(cur,bid,ask):
    global currency_df
    file_path = "exchange_adjust.xlsx"  # 記得替換成你的Excel檔案路徑
    currency_df = pd.read_excel(file_path, engine="openpyxl")
    if cur in currency_df["判斷B"].values:
        row_location = (currency_df["判斷B"] == cur).idxmax()
        currency_df.iloc[row_location,1] = bid
        currency_df.iloc[row_location,2] = ask
        currency_df.to_excel(file_path,index=False,engine="openpyxl")
        text_p7_response.config(text="已更新"+cur+"貨幣及買價: "+bid+"賣價: "+ask)
    else:
        newrow = [cur,bid,ask]
        currency_df.loc[len(currency_df)] = newrow
        currency_df.to_excel(file_path,index=False,engine="openpyxl")
        text_p7_response.config(text="已新增"+cur+"貨幣及買價: "+bid+"賣價: "+ask)
    # 設定三秒後清空文字
    text_p7_response.after(3000, clear_text_response)


page7 = ttk.Frame(notebook)
notebook.add(page7,text="調整貨幣")
page7.configure(style="Page1.TFrame")
#text_p7 = tk.Label(page7,text="快速查看貨幣")
#text_p7.pack()
quick_check_currency_button = tk.Button(page7,text="快速查看貨幣",fg="Black",bg="grey",font=("微軟正黑體",18),command=quick_check_currency_fn)
quick_check_currency_button.pack()
#
text_p7 = tk.Label(page7,text="請輸入貨幣",fg="Black",font=("微軟正黑體",18))
text_p7.pack()
entry_p7_currency = tk.Entry(page7,bg="white",font=("微軟正黑體",18))
entry_p7_currency.pack()
#
def validate_float(input_value):
    if input_value == "":
        return True
    try:
        float(input_value)
        return True
    except ValueError:
        return False

text_p7_bid = tk.Label(page7,text="請輸入Bid",fg="Black",font=("微軟正黑體",18))
text_p7_bid.pack()
entry_p7_bid = tk.Entry(page7, bg="white", font=("微軟正黑體", 18), validate="key")
entry_p7_bid.configure(validatecommand=(entry_p7_bid.register(validate_float), '%P'))
entry_p7_bid.pack()

#entry_p7_bid = tk.Entry(page7,bg="white",font=("微軟正黑體",18))
#entry_p7_bid.pack()
#
text_p7_ask = tk.Label(page7,text="請輸入Ask",fg="Black",font=("微軟正黑體",18))
text_p7_ask.pack()
entry_p7_ask = tk.Entry(page7, bg="white", font=("微軟正黑體", 18), validate="key")
entry_p7_ask.configure(validatecommand=(entry_p7_ask.register(validate_float), '%P'))
entry_p7_ask.pack()

#entry_p7_ask = tk.Entry(page7,bg="white",font=("微軟正黑體",18))
#entry_p7_ask.pack()
#   
def delete_currency(currency):
    global currency_df
    file_path = "exchange_adjust.xlsx"  # 記得替換成你的Excel檔案路徑
    currency_df = pd.read_excel(file_path, engine="openpyxl")
    if currency in currency_df["判斷B"].values:
        row_location = (currency_df["判斷B"] == currency).idxmax()
        currency_df = currency_df.drop(currency_df.index[row_location])
        currency_df.to_excel(file_path,index=False,engine="openpyxl")
        text_p7_response.config(text="已成功刪除"+currency+"貨幣並更新",fg="red")
    else:
        text_p7_response.config(text="不存在"+currency+"貨幣",fg="red")
    # 設定三秒後清空文字
    text_p7_response.after(3000, clear_text_response)

# 設定三秒後清空文字
def clear_text_response():
    text_p7_response.config(text="")

def update_currency_wraper():
    cur = entry_p7_currency.get()
    bid = entry_p7_bid.get()
    ask = entry_p7_ask.get()
    #ranking_degree = entry_p6_window_degree.get()
    update_currency(cur,bid,ask)
def delete_currency_wraper():
    currency = entry_p7_currency.get()
    delete_currency(currency)

update_currency_button = tk.Button(page7, text="新增，更新", fg="Black", bg="grey", font=("微軟正黑體", 18), command=update_currency_wraper, width=10)
update_currency_button.pack(padx=10, pady=10)

delete_currency_button = tk.Button(page7, text="刪除", fg="Black", bg="grey", font=("微軟正黑體", 18), command=delete_currency_wraper, width=10)
delete_currency_button.pack(padx=10, pady=10)

text_p7_response = tk.Label(page7,text="",fg="Black",font=("微軟正黑體",18))
text_p7_response.pack()
###########################################################################################################################
# 設定分頁的頁籤大小
style = ttk.Style()
style.configure("TNotebook.Tab", font=('微軟正黑體', 16), padding=[20, 8])  # 設定頁籤的字型和內部padding

notebook.pack(expand=True, fill='both')


root.mainloop()







