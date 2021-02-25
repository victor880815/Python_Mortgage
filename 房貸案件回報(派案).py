import tkinter
from tkinter import *
import tkinter.tix
from tkinter.tix import Tk, Control, ComboBox
from tkinter.messagebox import showinfo, showwarning, showerror
import tkinter as tk
import tkinter.ttk as ttk
import PIL
from PIL import ImageTk, Image, ImageSequence
import time
import pyodbc
import win32com.client as client

from base64 import b16encode

def rgb_color(rgb):
    return(b'#' + b16encode(bytes(rgb)))

#連接資料庫
conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\10.240.172.69\r52500_電話行銷科tmr共用\書翊\房貸商機清單.accdb;UID=Administrator;PWD=520888')
cursor = conn.cursor()

#設定主視窗
root = tk.Tk()
root.title("房貸案件回報系統©Victor Lin_M06429")
root.geometry("1920x1080")
root.resizable(width=True, height=True)

#主視窗背景
canvas = tk.Canvas(root, width=1920,height=1080,bd=0, highlightthickness=0)
imgpath = r'\\10.240.172.69\r52500_電話行銷科tmr共用\書翊\Portable Python-3.8.2\房貸主視窗(派案).png'
img = Image.open(imgpath)
photo = ImageTk.PhotoImage(img)


canvas.create_image(960,525, image=photo)
canvas.pack()
#===============================================================================================================================

#定義清除函數
def Reset():
    entry.delete(0,END),
    entry2.delete(0,END),
    entry3.delete(0,END),
    entry4.delete(0,END),
    combo.delete(0,END),
    combo2.delete(0,END),
    entry5.delete(0,END),
    combo3.delete(0,END),
    entry6.delete(0,END),
    entry7.delete(0,END),
    combo4.delete(0,END),
    entry8.delete(0,END),
    entry9.delete(0,END),
    combo5.delete(0,END),
    combo6.delete(0,END),
    combo7.delete(0,END),
    entry11.delete(0,END),
    combo9.delete(0,END),


#===============================================================================================================================


#定義"客戶姓名"搜尋函數
def search():
    try:
        cursor=conn.cursor()
        sql = "select * from 房貸商機清單 where 客戶姓名 = ?"
        cursor.execute(sql,(entry2.get(),))
        row=cursor.fetchone()

        entryvar.set(row[0])
        entryvar2.set(row[1])
        entryvar3.set(row[2])
        entryvar4.set(row[3])
        combovar.set(row[9])
        combovar2.set(row[10])
        entryvar5.set(row[11])
        combovar3.set(row[12])
        entryvar6.set(row[13])
        entryvar7.set(row[14])
        combovar4.set(row[15])
        entryvar8.set(row[8])
        entryvar9.set(row[7])
        combovar5.set(row[5])
        combovar6.set(row[16])
        combovar7.set(row[4])
        entryvar11.set(row[17])
        combovar9.set(row[18])


        conn.commit()
    except:
        tkinter.messagebox.showinfo("房貸案件回報系統","查無此筆案件")
        Reset()




#建立"客戶姓名" search 按鈕
searchbutton = tk.Button(root, command=search,text='搜尋',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                            ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=3,height=1,bd=2)
searchbutton.pack()



canvas.create_window(700, 173,window=searchbutton)


#===============================================================================================================================

#定義"身分證字號"搜尋函數
def search2():
    try:
        cursor=conn.cursor()
        sql2 = "select * from 房貸商機清單 where 身分證字號 = ?"
        cursor.execute(sql2,(entry3.get(),))
        row=cursor.fetchone()

        entryvar.set(row[0])
        entryvar2.set(row[1])
        entryvar3.set(row[2])
        entryvar4.set(row[3])
        combovar.set(row[9])
        combovar2.set(row[10])
        entryvar5.set(row[11])
        combovar3.set(row[12])
        entryvar6.set(row[13])
        entryvar7.set(row[14])
        combovar4.set(row[15])
        entryvar8.set(row[8])
        entryvar9.set(row[7])
        combovar5.set(row[5])
        combovar6.set(row[16])
        combovar7.set(row[4])
        entryvar11.set(row[17])
        combovar9.set(row[18])


        conn.commit()
    except:
        tkinter.messagebox.showinfo("房貸案件回報系統","查無此筆案件")
        Reset()




#建立"身分證字號"search 按鈕
searchbutton2 = tk.Button(root, command=search2,text='搜尋',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                            ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=3,height=1,bd=2)
searchbutton2.pack()



canvas.create_window(700, 238,window=searchbutton2)


#===============================================================================================================================

#定義"連絡電話"搜尋函數
def search3():
    try:
        cursor=conn.cursor()
        sql3 = "select * from 房貸商機清單 where 連絡電話 = ?"
        cursor.execute(sql3,(entry4.get(),))
        row=cursor.fetchone()

        entryvar.set(row[0])
        entryvar2.set(row[1])
        entryvar3.set(row[2])
        entryvar4.set(row[3])
        combovar.set(row[9])
        combovar2.set(row[10])
        entryvar5.set(row[11])
        combovar3.set(row[12])
        entryvar6.set(row[13])
        entryvar7.set(row[14])
        combovar4.set(row[15])
        entryvar8.set(row[8])
        entryvar9.set(row[7])
        combovar5.set(row[5])
        combovar6.set(row[16])
        combovar7.set(row[4])
        entryvar11.set(row[17])
        combovar9.set(row[18])

        conn.commit()
    except:
        tkinter.messagebox.showinfo("房貸案件回報系統","查無此筆案件")
        Reset()




#建立"連絡電話"search 按鈕
searchbutton3 = tk.Button(root, command=search3,text='搜尋',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                            ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=3,height=1,bd=2)
searchbutton3.pack()



canvas.create_window(700, 305,window=searchbutton3)



#===============================================================================================================================

#定義更新函數
def update():
    cursor=conn.cursor()

    cursor.execute("UPDATE 房貸商機清單 SET 日期yyyymmdd=?, 客戶姓名=?, 身分證字號=?, 連絡電話=?, 身分別=?, 縣市=?, 地址=?, 需求類別=?, 備註=?, 原始編碼or推薦人員編=?, 來源=?, 序號=?, 派案=?, 線上申請=?, 初次聯繫回報=?, eLoan申編=?, 銷件處理中說明=? WHERE 編號=?",(
    entryvar.get(),
    entryvar2.get(),
    entryvar3.get(),
    entryvar4.get(),
    combovar.get(),
    combovar2.get(),
    entryvar5.get(),
    combovar3.get(),
    entryvar6.get(),
    entryvar7.get(),
    combovar4.get(),
    entryvar8.get(),
    combovar5.get(),
    combovar6.get(),
    combovar7.get(),
    entryvar11.get(),
    combovar9.get(),
    entryvar9.get(),

    ))

    conn.commit()
    tkinter.messagebox.showinfo("房貸案件回報系統","更新成功！")
    Reset()


#建立更新按鈕
updatebutton = tk.Button(root, command=update,text='更新',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                            ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=3,height=1,bd=2)
updatebutton.pack()



canvas.create_window(1600, 570,window=updatebutton)




#===============================================================================================================================


#定義組別清單函數
def display():
    cursor=conn.cursor()
    sql4 = "select * from 房貸商機清單 where 派案 = ?"
    cursor.execute(sql4,(combovar5.get(),))
    result = cursor.fetchall()
    if len(result)!=0:
        records.delete(*records.get_children(),)
        for row in result:
            records.insert("",END,values = row[:20])

        conn.commit()

def Info(ev):
    viewInfo = records.focus()
    learnerData = records.item(viewInfo)
    row = learnerData['values']

    entryvar.set(row[0])
    entryvar2.set(row[1])
    entryvar3.set(row[2])
    entryvar4.set(row[3])
    combovar.set(row[9])
    combovar2.set(row[10])
    entryvar5.set(row[11])
    combovar3.set(row[12])
    entryvar6.set(row[13])
    entryvar7.set(row[14])
    combovar4.set(row[15])
    entryvar8.set(row[8])
    entryvar9.set(row[7])
    combovar5.set(row[5])
    combovar6.set(row[16])
    combovar7.set(row[4])
    entryvar11.set(row[17])
    combovar9.set(row[18])



displaybutton = tk.Button(root, command=display,text='檢視清單',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                            ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=7,height=1,bd=2)
displaybutton.pack()

canvas.create_window(1625, 437,window=displaybutton)
#===============================================================================================================================

def callbackFunc(event):
    cursor=conn.cursor()
    sql5="select * from 信箱 where 業務員 = ?"
    cursor.execute(sql5,(combovar8.get(),))
    row=cursor.fetchone()

    combovar8.set(row[0])
    entryvar10.set(row[1])

    conn.commit()

#===============================================================================================================================

#定義Email函數
def Email():

    body = entryvar2.get()
    body2 = entryvar3.get()
    body3 = entryvar4.get()
    body4 = combovar5.get()
    body5 = combovar7.get()

    body10 = entryvar10.get()

    outlook = client.Dispatch('Outlook.Application')

    message = outlook.CreateItem(0) # 0 is the code for a mail item (see the enumerations)
    message.Display()

    message.To = body10
    message.CC = 'M06429@sinopac.com'


    message.Subject = '您有新的案件請至系統查收'
    message.Body = "😊您有新的案件請至系統查收"+"\n"+body+"⚡"+body2+"⚡"+body3+"⚡"+body4+"⚡"+body5


    message.Save() # save to drafts folder
    message.Send() # send to outbox
    tkinter.messagebox.showinfo("房貸案件派案系統","信件寄送成功！")


Emailbutton = tk.Button(root, command=Email,text='寄送',bg='#AE0000',fg='white',activeforeground="black",activebackground='white'
                                        ,font=("微軟正黑體", 13, 'bold'),cursor='hand2', width=3,height=1,bd=2)
Emailbutton.pack()

canvas.create_window(1770, 90,window=Emailbutton)

#===============================================================================================================================

#定義時間函數
def gettime():
    entryvar14.set(time.strftime("%Y-%m-%d  %H:%M:%S"))
    root.after(1000, gettime)

#===============================================================================================================================

style = ttk.Style()
#Pick a theme
style.theme_use("default")
style.configure("Treeview.Heading", font=("微軟正黑體", 16, 'bold'), background=rgb_color((84, 130, 53)),foreground="white", fieldbackground=rgb_color((84, 130, 53)))
style.configure("Treeview",rowheight= 24, font=("微軟正黑體", 16), background="lightgrey",foreground="white", fieldbackground="lightgrey")
# Change selected color
style.map('Treeview',
background=[('selected', '#0080FF')])


#Treeview 建立
scroll_y = Scrollbar(root,orient = VERTICAL)

records = ttk.Treeview(root,height = 10,columns = ("日期yyyymmdd","客戶姓名","身分證字號","連絡電話","初次聯繫回報","派案","時間","編號","序號","身分別","縣市","地址","需求類別"
                                                    ,"備註","原始編碼or推薦人員編","來源","線上申請"
                                                    ,"eLoan申編","銷件處理中說明"),yscrollcommand = scroll_y.set)
scroll_y.pack()




records.heading("日期yyyymmdd",text="日期yyyymmdd")
records.heading("客戶姓名",text="客戶姓名")
records.heading("身分證字號",text="身分證字號")
records.heading("連絡電話",text="連絡電話")
records.heading("初次聯繫回報",text="初次聯繫回報")
records.heading("派案",text="派案")
records.heading("時間",text="時間")
records.heading("編號",text="編號")
records.heading("序號",text="序號")
records.heading("身分別",text="身分別")
records.heading("縣市",text="縣市")
records.heading("地址",text="地址")
records.heading("需求類別",text="需求類別")
records.heading("備註",text="備註")
records.heading("原始編碼or推薦人員編",text="原始編碼or推薦人員編")
records.heading("來源",text="來源")
records.heading("線上申請",text="線上申請")
records.heading("eLoan申編",text="eLoan申編")
records.heading("銷件處理中說明",text="銷件處理中說明")

records['show'] = 'headings'

records.column("日期yyyymmdd", width = 180)
records.column("客戶姓名", width = 90)
records.column("身分證字號", width = 130)
records.column("連絡電話", width = 90)
records.column("初次聯繫回報", width = 150)
records.column("派案", width = 60)
records.column("時間", width = 60)
records.column("編號", width = 60)
records.column("序號", width = 60)
records.column("身分別", width = 70)
records.column("縣市", width = 65)
records.column("地址", width = 65)
records.column("需求類別", width = 90)
records.column("備註", width = 65)
records.column("原始編碼or推薦人員編", width = 235)
records.column("來源", width = 60)
records.column("線上申請", width = 90)
records.column("eLoan申編", width = 125)
records.column("銷件處理中說明", width = 175)

records.pack()
records.bind("<ButtonRelease-1>",Info)


canvas.create_window(959, 870,window=records)


#===============================================================================================================================

#日期yyyymmdd entry 建立
entryvar = tk.StringVar()
entry = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar)
entry.pack()
canvas.create_window(510, 110, width=280,height=50,
                               window=entry)


#客戶姓名 entry 建立
entryvar2=tk.StringVar()
entry2 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar2)
entry2.pack()
canvas.create_window(510, 173, width=280,height=50,
                                   window=entry2)

#身分證字號 entry 建立
entryvar3=tk.StringVar()
entry3 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar3)
entry3.pack()
canvas.create_window(510, 238, width=280,height=50,
                                   window=entry3)

#連絡電話 entry 建立
entryvar4=tk.StringVar()
entry4 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar4)
entry4.pack()
canvas.create_window(510, 305, width=280,height=50,
                                   window=entry4)

#身分別 combobox 建立
combovar=tk.StringVar()
combo = ttk.Combobox(root,value = ['上班族','企業主或自營商','自由業或家管'],font=("微軟正黑體",22),textvariable=combovar)

combo.pack()

canvas.create_window(510, 370, width=280, height=50,window=combo)

#縣市 combobox 建立
combovar2=tk.StringVar()
combo2 = ttk.Combobox(root,value = ['台中','台北','台南','宜蘭','花蓮','南投','屏東','桃園','高雄','基隆','新北','新竹','嘉義','彰化'],font=("微軟正黑體",22),textvariable=combovar2)

combo2.pack()

canvas.create_window(510, 437, width=280, height=50,window=combo2)


#地址 entry 建立
entryvar5=tk.StringVar()
entry5 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar5)
entry5.pack()
canvas.create_window(627, 503, width=515,height=50,
                                  window=entry5)


#需求類別 combobox 建立
combovar3=tk.StringVar()
combo3 = ttk.Combobox(root,value = ['他行房貸轉(增)貸','本行既有房貸戶欲增貸','原屋融資(房子目前已無貸款)','新購屋'],font=("微軟正黑體",22) ,textvariable=combovar3)

combo3.pack()

canvas.create_window(510, 568, width=280, height=50,window=combo3)


#eLoan申編 entry 建立
entryvar11=tk.StringVar()
entry11 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",18),highlightcolor=rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar11)
entry11.pack()
canvas.create_window(510, 636, width=280,height=50,
                                   window=entry11)


#備註 entry 建立
entryvar6=tk.StringVar()
entry6 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar6)
entry6.pack()
canvas.create_window(1140, 703, width=1510,height=50,
                                   window=entry6)


#原始編碼/推薦人員編 entry 建立
entryvar7=tk.StringVar()
entry7 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar7)
entry7.pack()
canvas.create_window(1410, 173, width=280,height=50,
                                   window=entry7)

#來源 combobox 建立
combovar4=tk.StringVar()
combo4 = ttk.Combobox(root,value = ['C01-客關商機','C01-客關部商機','M04-房貸商品申請','M05-豐雲e房貸','T01-電銷'],font=("微軟正黑體",22) ,textvariable=combovar4)

combo4.pack()

canvas.create_window(1410, 238, width=280, height=50,window=combo4)


#序號 entry 建立
entryvar8=tk.StringVar()
entry8 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar8)
entry8.pack()
canvas.create_window(1410, 305, width=280,height=50,
                                   window=entry8)


#編號 entry 建立
entryvar9=tk.StringVar()
entry9 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar9)
entry9.pack()
canvas.create_window(1410, 370, width=280,height=50,
                                   window=entry9)


#派案 combobox 建立
combovar5=tk.StringVar()
combo5 = ttk.Combobox(root,value = ['王念傑','杜茂霖','林秀玲','林孟翰','林婷婷','張瑜芬','陳杰','陳家華','陳偉仁','黃心怡','黃瑞珩','黃麗敏','葉子弘'
                                   ,'葉明艷','劉倢希','戴嘉欣'],font=("微軟正黑體",22) ,textvariable=combovar5)

combo5.pack()

canvas.create_window(1410, 437, width=280, height=50,window=combo5)


#線上申請 combobox 建立
combovar6=tk.StringVar()
combo6 = ttk.Combobox(root,value = ['已派過，登線上','預計線上申請'],font=("微軟正黑體",22) ,textvariable=combovar6)

combo6.pack()

canvas.create_window(1410, 503, width=280, height=50,window=combo6)

#初次聯繫回報 combobox 建立
combovar7=tk.StringVar()
combo7 = ttk.Combobox(root,value = ['未進件電聯中','未進件銷件(選擇原因)','eLoan起案(回報申編)','徵信完成','估價完成','未送簽銷件(選擇原因)','核貸書編輯中'
                                   ,'核貸書送簽中','未核准申覆中','簽核完成','簽約完成','對保完成','已撥款'],font=("微軟正黑體",22) ,textvariable=combovar7)

combo7.pack()

canvas.create_window(1410, 570, width=280, height=50,window=combo7)


#銷件處理中說明 combobox 建立
combovar9=tk.StringVar()
combo9 = ttk.Combobox(root,value = ['婉拒-客戶收支無空間','婉拒-已知信用瑕疵','婉拒-擔保品條件不符','婉拒-無法估價','婉拒-失聯','婉拒-借戶條件不佳'
                                   ,'客戶取消-利率條件不符','客戶取消-貸款額度不符','客戶取消-不願配合補件','客戶取消-轉貸他行','客戶取消-無資金需求'
                                   ,'客戶取消-還未找到標的物','客戶取消-銷件重進','處理中-待確認徵提保人的可能性','處理中-待補文件','處理中-準備送簽'
                                   ,'處理中-聯絡不上客戶','處理中-客戶考慮中','處理中-準備送撥','重複進件','已撥款','待移轉分行'],font=("微軟正黑體",22) ,textvariable=combovar9)

combo9.pack()

canvas.create_window(1490, 636, width=440, height=50,window=combo9)








#業務員信箱 entry 建立
combovar8=tk.StringVar()
combo8 = ttk.Combobox(root,value = ['林書翊','吳玉汝','衛佳筠'],font=("微軟正黑體",18),textvariable=combovar8)

combo8.pack()

canvas.create_window(1361, 90, width=180, height=40,window=combo8)

#信箱 entry 建立
entryvar10=tk.StringVar()
entry10 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",18),highlightcolor=rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar10)
entry10.pack()
canvas.create_window(1590, 90, width=280,height=40,
                                   window=entry10)
combo8.bind("<<ComboboxSelected>>", callbackFunc)


#時間 entry 建立
entryvar14=tk.StringVar()
entry14 = tk.Label(root,font=("微軟正黑體",14,"bold"),fg = "white",bg = rgb_color((84, 130, 53)),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar14)
entry14.pack()
canvas.create_window(1813, 17, width=210,height=15,
                                   window=entry14)
gettime()


# 下拉框颜色
#combostyle = ttk.Style()
#combostyle.theme_create('combostyle', parent='alt',
#                        settings={'TCombobox':
#                            {'configure':
#                                {
#                                    'foreground': 'blue',  # 前景色
#                                    'selectbackground': 'black',  # 选择后的背景颜色
#                                    'fieldbackground': 'white',  # 下拉框颜色
#                                    'background': 'red',  # 下拉按钮颜色
#                                }}}
#                        )
#combostyle.theme_use('combostyle')


root.mainloop()
