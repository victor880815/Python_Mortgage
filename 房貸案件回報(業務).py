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
root.geometry("1280x1024")
root.resizable(width=True, height=True)

#主視窗背景
canvas = tk.Canvas(root, width=1920,height=1080,bd=0, highlightthickness=0)
imgpath = r'\\10.240.172.69\r52500_電話行銷科tmr共用\書翊\Portable Python-3.8.2\房貸主視窗.png'
img = Image.open(imgpath)
photo = ImageTk.PhotoImage(img)


canvas.create_image(640,510, image=photo)
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



#===============================================================================================================================

#定義"客戶姓名"搜尋函數
def search():
    try:
        cursor=conn.cursor()
        sql = "select * from 房貸商機清單 where 客戶姓名 = ?"
        cursor.execute(sql,(entry2.get(),))
        row=cursor.fetchone()

        combovar.set(row[5])
        entryvar2.set(row[1])
        entryvar3.set(row[2])
        entryvar4.set(row[3])
        combovar2.set(row[4])
        entryvar5.set(row[7])


        conn.commit()
    except:
        tkinter.messagebox.showinfo("房貸案件回報系統","查無此筆案件")
        Reset()




#建立"客戶姓名" search 按鈕
searchbutton = tk.Button(root, command=search,text='搜尋',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                            ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=3,height=1,bd=2)
searchbutton.pack()



canvas.create_window(700, 207,window=searchbutton)


#===============================================================================================================================


#定義"身分證字號"搜尋函數
def search2():
    try:
        cursor=conn.cursor()
        sql2 = "select * from 房貸商機清單 where 身分證字號 = ?"
        cursor.execute(sql2,(entry3.get(),))
        row=cursor.fetchone()

        combovar.set(row[5])
        entryvar2.set(row[1])
        entryvar3.set(row[2])
        entryvar4.set(row[3])
        combovar2.set(row[4])
        entryvar5.set(row[7])

        conn.commit()
    except:
        tkinter.messagebox.showinfo("房貸案件回報系統","查無此筆案件")
        Reset()




#建立"身分證字號"search 按鈕
searchbutton2 = tk.Button(root, command=search2,text='搜尋',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                            ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=3,height=1,bd=2)
searchbutton2.pack()



canvas.create_window(700, 305,window=searchbutton2)


#===============================================================================================================================


#定義"連絡電話"搜尋函數
def search3():
    try:
        cursor=conn.cursor()
        sql3 = "select * from 房貸商機清單 where 連絡電話 = ?"
        cursor.execute(sql3,(entry4.get(),))
        row=cursor.fetchone()

        combovar.set(row[5])
        entryvar2.set(row[1])
        entryvar3.set(row[2])
        entryvar4.set(row[3])
        combovar2.set(row[4])
        entryvar5.set(row[7])

        conn.commit()
    except:
        tkinter.messagebox.showinfo("房貸案件回報系統","查無此筆案件")
        Reset()




#建立"連絡電話"search 按鈕
searchbutton3 = tk.Button(root, command=search3,text='搜尋',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                            ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=3,height=1,bd=2)
searchbutton3.pack()



canvas.create_window(700, 404,window=searchbutton3)



#===============================================================================================================================


#定義更新、新增函數
def update():
    cursor=conn.cursor()

    cursor.execute("UPDATE 房貸商機清單 SET 客戶姓名=?, 身分證字號=?, 連絡電話=?, 初次聯繫回報=? ,時間=? WHERE 編號=?",(

    entryvar2.get(),
    entryvar3.get(),
    entryvar4.get(),
    combovar2.get(),
    entryvar14.get(),

    entryvar5.get(),

    ))

    conn.commit()

    cursor2=conn.cursor()
    cursor2.execute("insert into 房貸商機清單歷史紀錄 values(?,?,?,?,?,?,?)",(

    entryvar2.get(),
    entryvar3.get(),
    entryvar4.get(),
    combovar2.get(),
    combovar.get(),
    entryvar14.get(),
    entryvar5.get(),

    ))
    conn.commit()

    tkinter.messagebox.showinfo("房貸案件回報系統","更新成功！")


#建立更新、新增按鈕
updatebutton = tk.Button(root, command=update,text='更新',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                            ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=3,height=1,bd=2)
updatebutton.pack()



canvas.create_window(700, 502,window=updatebutton)




#===============================================================================================================================

#定義組別清單函數
def display():
    cursor=conn.cursor()
    sql4 = "select * from 房貸商機清單 where 派案 = ?"
    cursor.execute(sql4,(combovar.get(),))
    result = cursor.fetchall()
    if len(result)!=0:
        records.delete(*records.get_children(),)
        for row in result:
            records.insert("",END,values = row[1:8])

        conn.commit()

def Info(ev):
    viewInfo = records.focus()
    learnerData = records.item(viewInfo)
    row = learnerData['values']


    entryvar2.set(row[0])
    entryvar3.set(row[1])
    entryvar4.set(row[2])
    combovar2.set(row[3])
    combovar.set(row[4])
    entryvar14.set(row[5])
    entryvar5.set(row[6])




displaybutton = tk.Button(root, command=display,text='檢視清單',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                            ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=7,height=1,bd=2)
displaybutton.pack()

canvas.create_window(1220, 108,window=displaybutton)
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

records = ttk.Treeview(root,height = 13,columns = ("客戶姓名","身分證字號","連絡電話","初次聯繫回報"
                                                    ,"業務員","時間","編號"),yscrollcommand = scroll_y.set)
scroll_y.pack()




records.heading("客戶姓名",text="客戶姓名")
records.heading("身分證字號",text="身分證字號")
records.heading("連絡電話",text="連絡電話")
records.heading("初次聯繫回報",text="初次聯繫回報")
records.heading("業務員",text="業務員")
records.heading("時間",text="時間")
records.heading("編號",text="編號")




records['show'] = 'headings'


records.column("客戶姓名", width = 100)
records.column("身分證字號", width = 180)
records.column("連絡電話", width = 180)
records.column("初次聯繫回報", width = 230)
records.column("業務員", width = 100)
records.column("時間", width = 230)
records.column("編號", width = 230)



records.pack()
records.bind("<ButtonRelease-1>",Info)


canvas.create_window(638, 820,window=records)


#===============================================================================================================================

imgpath2 = r'\\10.240.172.69\r52500_電話行銷科tmr共用\書翊\Portable Python-3.8.2\房貸子視窗.png'
img2 = Image.open(imgpath2)
photo2 = ImageTk.PhotoImage(img2)


#子視窗建立
def second_win():
    top = tk.Toplevel()
    top.title('房貸案件回報系統')
    top.geometry("800x800")

    #子視窗背景
    canvas2 = tk.Canvas(top, width=800,height=800,bd=0, highlightthickness=0)
    canvas2.create_image(400, 400, image=photo2)
    canvas2.pack()
    #===============================================================================================================================

    #定義子視窗"客戶姓名"搜尋函數
    def secondsearch():
        try:
            cursor=conn.cursor()
            sql5 = "select * from 房貸商機清單 where 客戶姓名 = ?"
            cursor.execute(sql5,(secondentry.get(),))
            row=cursor.fetchone()

            secondentryvar.set(row[1])
            secondentryvar2.set(row[2])
            secondentryvar3.set(row[3])
            secondentryvar4.set(row[5])
            secondentryvar5.set(row[6])
            secondentryvar6.set(row[4])
            secondentryvar7.set(row[7])

            conn.commit()
        except:
            tkinter.messagebox.showinfo("房貸案件回報系統","查無此筆案件")
            Reset()




    #建立子視窗"客戶姓名" search 按鈕
    secondsearchbutton = tk.Button(top, command=secondsearch,text='搜尋',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                                ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=3,height=1,bd=2)
    secondsearchbutton.pack()



    canvas2.create_window(320, 50,window=secondsearchbutton)

    #===============================================================================================================================

    #定義子視窗"身分證字號"搜尋函數
    def secondsearch2():
        try:
            cursor=conn.cursor()
            sql6 = "select * from 房貸商機清單 where 身分證字號 = ?"
            cursor.execute(sql6,(secondentry2.get(),))
            row=cursor.fetchone()

            secondentryvar.set(row[1])
            secondentryvar2.set(row[2])
            secondentryvar3.set(row[3])
            secondentryvar4.set(row[5])
            secondentryvar5.set(row[6])
            secondentryvar6.set(row[4])
            secondentryvar7.set(row[7])

            conn.commit()
        except:
            tkinter.messagebox.showinfo("房貸案件回報系統","查無此筆案件")
            Reset()




    #建立子視窗"身分證字號" search 按鈕
    secondsearchbutton2 = tk.Button(top, command=secondsearch2,text='搜尋',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                                ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=3,height=1,bd=2)
    secondsearchbutton2.pack()



    canvas2.create_window(320, 252,window=secondsearchbutton2)

    #===============================================================================================================================

    #定義子視窗"連絡電話"搜尋函數
    def secondsearch3():
        try:
            cursor=conn.cursor()
            sql7 = "select * from 房貸商機清單 where 連絡電話 = ?"
            cursor.execute(sql7,(secondentry3.get(),))
            row=cursor.fetchone()

            secondentryvar.set(row[1])
            secondentryvar2.set(row[2])
            secondentryvar3.set(row[3])
            secondentryvar4.set(row[5])
            secondentryvar5.set(row[6])
            secondentryvar6.set(row[4])
            secondentryvar7.set(row[7])

            conn.commit()
        except:
            tkinter.messagebox.showinfo("房貸案件回報系統","查無此筆案件")
            Reset()




    #建立子視窗"身分證字號" search 按鈕
    secondsearchbutton3 = tk.Button(top, command=secondsearch3,text='搜尋',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                                ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=3,height=1,bd=2)
    secondsearchbutton3.pack()



    canvas2.create_window(320, 455,window=secondsearchbutton3)

    #===============================================================================================================================
    #子視窗客戶姓名 entry 建立
    secondentryvar=tk.StringVar()
    secondentry = tk.Entry(top, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=secondentryvar)
    secondentry.pack()
    canvas2.create_window(205, 120, width=280,height=50,
                                       window=secondentry)

    #子視窗身分證字號 entry 建立
    secondentryvar2=tk.StringVar()
    secondentry2 = tk.Entry(top, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=secondentryvar2)
    secondentry2.pack()
    canvas2.create_window(205, 330, width=280,height=50,
                                       window=secondentry2)

    #子視窗連絡電話 entry 建立
    secondentryvar3=tk.StringVar()
    secondentry3 = tk.Entry(top, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=secondentryvar3)
    secondentry3.pack()
    canvas2.create_window(205, 535, width=280,height=50,
                                       window=secondentry3)

    #子視窗業務員 entry 建立
    secondentryvar4=tk.StringVar()
    secondentry4 = tk.Entry(top, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=secondentryvar4)
    secondentry4.pack()
    canvas2.create_window(205, 735, width=280,height=50,
                                       window=secondentry4)

    #子視窗最後聯繫時間 entry 建立
    secondentryvar5=tk.StringVar()
    secondentry5 = tk.Entry(top, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=secondentryvar5)
    secondentry5.pack()
    canvas2.create_window(580, 120, width=300,height=50,
                                       window=secondentry5)

    #子視窗初次聯繫回報 entry 建立
    secondentryvar6=tk.StringVar()
    secondentry6 = tk.Entry(top, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=secondentryvar6)
    secondentry6.pack()
    canvas2.create_window(580, 330, width=300,height=50,
                                       window=secondentry6)

    #子視窗編號 entry 建立
    secondentryvar7=tk.StringVar()
    secondentry7 = tk.Entry(top, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=secondentryvar7)
    secondentry7.pack()
    canvas2.create_window(580, 535, width=300,height=50,
                                       window=secondentry7)










    top,mainloop()

#建立按鈕
button=Button(root,text='其他業務員\n案件查詢',command=second_win,bg='#AE0000',fg='white',activeforeground="black",activebackground='#66B3FF',font=("微軟正黑體", 20, 'bold'),cursor='hand2',bd=3, width=10, height=2)
button.pack()

canvas.create_window(1173, 590,window=button)




#===============================================================================================================================

#定義密碼搜尋函數
def passwordsearch():
    try:
        cursor=conn.cursor()
        sql8 = "select * from 密碼 where 密碼 = ?"
        cursor.execute(sql8,(entryvar.get(),))
        row=cursor.fetchone()

        entryvar.set(row[0])
        combovar.set(row[1])


        conn.commit()
    except:
        tkinter.messagebox.showinfo("房貸案件回報系統","密碼輸入錯誤")
        Reset()




#建立子視窗"客戶姓名" search 按鈕
passwordbutton = tk.Button(root, command=passwordsearch,text='確認',bg=rgb_color((84, 130, 53)),fg='white',activeforeground="black",activebackground='white'
                            ,font=("微軟正黑體", 16, 'bold'),cursor='hand2', width=3,height=1,bd=2)
passwordbutton.pack()



canvas.create_window(700, 108,window=passwordbutton)

#===============================================================================================================================

#輸入使用者密碼 entry 建立
entryvar = tk.StringVar()
entry = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar, show='*')
entry.pack()
canvas.create_window(515, 108, width=280,height=50,
                               window=entry)

#業務員 combobox 建立
combovar=tk.StringVar()
combo = ttk.Combobox(root,value = ['王念傑','杜茂霖','林秀玲','林孟翰','林婷婷','張瑜芬','陳杰','陳家華','陳偉仁','黃心怡','黃瑞珩','黃麗敏','葉子弘'
                                   ,'葉明艷','劉倢希','戴嘉欣'],font=("微軟正黑體",22),textvariable=combovar, state="disable")

combo.pack()

canvas.create_window(1040, 108, width=230, height=50,window=combo)

#客戶姓名 entry 建立
entryvar2=tk.StringVar()
entry2 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar2)
entry2.pack()
canvas.create_window(515, 207, width=280,height=50,
                                   window=entry2)


#身分證字號 entry 建立
entryvar3=tk.StringVar()
entry3 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar3)
entry3.pack()
canvas.create_window(515, 305, width=280,height=50,
                                   window=entry3)


#連絡電話 entry 建立
entryvar4=tk.StringVar()
entry4 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar4)
entry4.pack()
canvas.create_window(515, 404, width=280,height=50,
                                   window=entry4)

#初次聯繫回報 combobox 建立
combovar2=tk.StringVar()
combo2 = ttk.Combobox(root,value = ['未進件電聯中','未進件銷件(選擇原因)','eLoan起案(回報申編)','徵信完成','估價完成','未送簽銷件(選擇原因)','核貸書編輯中'
                                   ,'核貸書送簽中','未核准申覆中','簽核完成','簽約完成','對保完成','已撥款'],font=("微軟正黑體",22) ,textvariable=combovar2)

combo2.pack()

canvas.create_window(515, 502, width=280, height=50,window=combo2)


#編號 entry 建立
entryvar5=tk.StringVar()
entry5 = tk.Entry(root, insertbackground='black',font=("微軟正黑體",22),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar5)
entry5.pack()
canvas.create_window(515, 602, width=280,height=50,
                                   window=entry5)

#時間 entry 建立
entryvar14=tk.StringVar()
entry14 = tk.Label(root,font=("微軟正黑體",14,"bold"),fg = "white",bg = rgb_color((84, 130, 53)),highlightcolor= rgb_color((84, 130, 53)) ,highlightthickness =2, textvariable=entryvar14)
entry14.pack()
canvas.create_window(1182, 17, width=210,height=15,
                                   window=entry14)
gettime()
#===============================================================================================================================

root.mainloop()
