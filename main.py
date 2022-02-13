from tkinter import *
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy
import os
from os import system,name
import xlrd
from PIL import ImageTk
from tkinter import messagebox
from tkinter import ANCHOR
import time
import random


path = os.getcwd()
os.chdir(path)

def login2(event):
    global user
    username=usern.get()
    password=passw.get()
    user=username

    if username.lower()=='admin' and password=='admin':
        loginw.destroy()
        opt()
    elif username.lower()!='admin' and password=='student':
        loginw.destroy()
        opt()
    else:
        messagebox.showinfo('Login unsuccessful','The username or password is incorrect.')
def login():
    global loginw
    global usern
    global passw
    loginw=Tk()
    loginw.title('BookWiki  (login page)')
    loginw.geometry('500x210')
    usern=StringVar()
    passw=StringVar()
    Label(loginw,text='Welcome to BookWiki',font=('kristen itc',30)).pack()
    Label(loginw,text='Username:').pack()
    Entry(loginw,textvariable=usern).pack()
    Label(loginw,text='').pack()
    Label(loginw,text='Password:').pack()
    Entry(loginw,textvariable=passw).pack()
    Label(loginw,text='').pack()
    lb=Button(loginw,text='login',bg='blue',fg='white',command=login2)
    lb.pack()
    lb.bind('<Return>',login2)

def submit():
    bookN_input=bn.get()
    bookAuth_input=ba.get()
    bookID_input=bid.get()
    book_genre=bg.get()
    book_details=bo.get()
    book_price=bp.get()

    cont=True
    if bookN_input=='' or bookAuth_input=='' or book_genre=='' or book_details=='' or book_price=='':
        messagebox.showinfo('Entry Error','All fields are compulsary')
        cont=False
    if cont==True and book_price.isnumeric()==False:
        messagebox.showinfo('Entry Error','Price should be an integer value')
        cont=False
    if cont==True:
        ori_wb='Book details.xls'
        rb = open_workbook(ori_wb)
        wb=copy(rb)

        s=wb.get_sheet(0)


        
        cell=open('cell.txt','r')
        row=int(cell.read())
        cell.close()

        s.write(row,0,bookID_input)
        s.write(row,1,bookN_input)
        s.write(row,2,bookAuth_input)
        s.write(row,3,book_genre)
        s.write(row,4,book_details)
        s.write(row,5,int(book_price))
        s.write(row,6,'AVAILABLE')
                
        row+=1

        os.remove(ori_wb)
        wb.save(ori_wb)
        os.remove('cell.txt')
        new_cell = open('cell.txt','w')
        new_row = new_cell.write(str(row))
        new_cell.close()
        bid.set('')
        bn.set('')
        ba.set('')
        bg.set('')
        bo.set('')
        bp.set('')

        messagebox.showinfo('Status','Registration Successful')
        addw.destroy()
        opt()
    
def aback():
    addw.destroy()
    opt()

def add():
    if user=='admin':
        global addw
        root.destroy()
        addw=Tk()
        addw.state('zoomed')

        canvas = Canvas(width = 1000, height = 500, bg = 'blue')
        canvas.pack(expand = YES, fill = BOTH)

        image = ImageTk.PhotoImage(file = path+"/a.png")
        canvas.create_image(0, 0, image = image, anchor = NW)

        global bid
        global bn
        global ba
        global bg
        global bo
        global bp

        bid=StringVar()
        bn=StringVar()
        ba=StringVar()
        bg=StringVar()
        bo=StringVar()
        bp=StringVar()
        
        add_title = Label(addw,text='Enter Book Details',font=('Arial Black',30),bg='blue')
        bidl = Label(addw,text='Book id:',font=('Bookman Old Style',18),bg='black',fg='white')
        bide = Entry(addw,textvariable=bid,width=50)
        
        bnl = Label(addw,text='Book name:',font=('Bookman Old Style',18),bg='black',fg='white')
        bne = Entry(addw,textvariable=bn,width=50)
        
        bal = Label(addw,text='Book Author:',font=('Bookman Old Style',18),bg='black',fg='white')
        bae = Entry(addw,textvariable=ba,width=50)
        
        bgl = Label(addw,text='Genre:',font=('Bookman Old Style',18),bg='black',fg='white')
        bge = Entry(addw,textvariable=bg,width=50)
        
        bol = Label(addw,text='Book Overview:',font=('Bookman Old Style',18),bg='black',fg='white')
        boe = Entry(addw,textvariable=bo,width=50)
        
        bpl = Label(addw,text='Price:',font=('Bookman old style',18),bg='black',fg='white')
        bpe = Entry(addw,textvariable=bp,width=50)
        
        bsb = Button(addw,text='Submit',font=('mv boli',30),bg='grey',command=submit)
        Button(addw,text='Back',font=('segoe script',20),fg='dark green',bg='yellow',command=aback).place(relx=0.88,rely=0.88)
        

        add_title.place(x=450,y=30)
        bidl.place(x=398,y=270)
        bide.place(x=586,y=275)
        bnl.place(x=398,y=298)
        bne.place(x=586,y=305)
        bal.place(x=398,y=329)
        bae.place(x=586,y=335)
        bgl.place(x=398,y=360)
        bge.place(x=586,y=365)
        bol.place(x=398,y=391)
        boe.place(x=586,y=399)
        bpl.place(x=398,y=423)
        bpe.place(x=586,y=433)
        bsb.place(x=600,y=500)
        
        addw.mainloop()
    else:
        messagebox.showinfo('Access denied','You do not have access to admin features')
def remback():
    remw.destroy()
    opt()
    
def r_name(value):
    global del_name
    
    del_name=value.lower()
    aq = messagebox.askquestion('Confirm action','Are you sure you want to remove "'+value+'" from the database?')
    if aq=='yes':
        remove2()

   

def remove1():
    if user=='admin':
        global di
        global cells
        global s
        global rb
        global remw
        
        try:        
            root.destroy()
        except:
            pass
        
        remw = Tk()
        remw.state('zoomed')
        
        canvas = Canvas(width = 1000, height = 500, bg = 'blue')
        canvas.pack(expand = YES, fill = BOTH)

        image = ImageTk.PhotoImage(file = path+"/lro.png")
        canvas.create_image(0, 0, image = image, anchor = NW)

        rb=xlrd.open_workbook('Book details.xls')
            
            
        s=rb.sheet_by_index(0)
        cell=open('cell.txt','r+')
        cells=cell.read()
        cell.close()
        
        Label(remw,text='Remove Book',font=('Gabriola',28),bg='sky blue').place(relx=0.45,rely=0.23)
        Button(remw,text='Back',font=('segoe script',20),fg='dark green',bg='yellow',command=remback).place(relx=0.88,rely=0.88)
        di={}

        for i in range(2,int(cells)):
                di.update({i:s.cell_value(i,1)})
                   
        variable = StringVar()
        variable.set('Choose a Book')
        
        p=OptionMenu(remw, variable, *di.values(), command=r_name)
        p.place(relx=0.4,rely=0.4)
        p.configure(bg='light green',fg='red',font=('papyrus',25,'bold'))
        remw.mainloop()
    else:
        messagebox.showinfo('Access denied','You do not have access to admin features')

def remove2():
        permission=0
        for i in range(0,int(cells)):
                    if  s.cell_value(i,1).lower() == del_name and s.cell_value(i,6)!='ISSUED':
                        row_n = i
                        permission=1
                        break
        if permission==0:
                        messagebox.showinfo('Already Issued','The book is currently issued')
                        remw.destroy()
                        remove1()
                        
        if permission==1:                               
            nwb=copy(rb)
            sheet=nwb.get_sheet(0)
                    
            for j in range(0,row_n):
                    for k in range(0,5):
                            sheet.write(j,k,s.cell_value(j,k))
                            
            for j in range(row_n,int(cells)-1):
                        for k in range(0,5):
                            sheet.write(j,k,s.cell_value(j+1,k))
                            
            for k in range(0,5):
                        sheet.write(j+1,k,'')               #removing last row

            new_cells = int(cells) - 1
            cell=open('cell.txt','w')
            cell.write(str(new_cells))
            cell.close()
                    
            os.remove('Book details.xls')
            nwb.save('Book details.xls')

            messagebox.showinfo('Status','Deleted book "'+del_name+'" from the database.')

            remw.destroy()
            opt()
        
def current(*event):
        global ovlabel
        global idlabel
        global autlabel
        global genlabel


        selected=listbox0.get(ANCHOR)

        num=''
        for i in range(len(selected)):
            if selected[i]==')':
                break
            else:
                num+=selected[i]
        index=int(num)+1
        
        alpha=False
        if alpha==False:
            selected_id = str(s.cell_value(index,0))
            selected_aut = str(s.cell_value(index,2))
            selected_gen = str(s.cell_value(index,3))
            selected_ov = str(s.cell_value(index,4))
            selected_price = str(s.cell_value(index,5))

        try :
            ovlabel.destroy()
        except:
            pass
        try:
           idlabel.destroy()
        except:
            pass
        try:
            genlabel.destroy()
        except:
            pass
        try:
            autlabel.destroy()
        except:
            pass
        
        idlabel=Label(show_w,text='Book id: '+selected_id,font=('courier new greek',20,'bold'))
        idlabel.place(relx=0.08,rely=0.65)
        autlabel=Label(show_w,text='Author: '+selected_aut,font=('courier new greek',20,'bold'))
        autlabel.place(relx=0.08,rely=0.7)
        genlabel=Label(show_w,text='Genre: '+selected_gen,font=('courier new greek',20,'bold'))
        genlabel.place(relx=0.08,rely=0.75)
        ovlabel=Label(show_w,text='Overview: '+selected_ov,font=('courier new greek',20,'bold'))
        ovlabel.place(relx=0.08,rely=0.8)
        

def currentalpha(*event):
        global ovlabel
        global idlabel
        global autlabel
        global genlabel
        global alpha
        
        selected=listbox0.get(ANCHOR)
        
        num=''
        for i in range(len(selected)):
            if selected[i]==')':
                break
            else:
                num+=selected[i]
        index=int(num)-1
        
                
        selected_id = rid[index]
        selected_aut = ra[index]
        selected_gen = rg[index]
        selected_ov = ro[index]
        

        try:
            ovlabel.destroy()
        except:
            pass
        try:
           idlabel.destroy()
        except:
            pass
        try:
            genlabel.destroy()
        except:
            pass
        try:
            autlabel.destroy()
        except:
            pass
        
        idlabel=Label(show_w,text='Book id: '+selected_id,font=('courier new greek',20,'bold'))
        idlabel.place(relx=0.08,rely=0.65)
        autlabel=Label(show_w,text='Author: '+selected_aut,font=('courier new greek',20,'bold'))
        autlabel.place(relx=0.08,rely=0.7)
        genlabel=Label(show_w,text='Genre: '+selected_gen,font=('courier new greek',20,'bold'))
        genlabel.place(relx=0.08,rely=0.75)
        ovlabel=Label(show_w,text='Overview: '+selected_ov,font=('courier new greek',20,'bold'))
        ovlabel.place(relx=0.08,rely=0.8)

        alpha=True


def alpha():                               #arrange alphabetically
    global listbox0
    global rn
    global rid
    global rg
    global ro
    global ra
    global dn
    global scrollbar

    ab.destroy()
    pb.destroy()
    gb.destroy()
    rn=sorted(dn.values())      #arrange names alphabetically
           
    scrollbar.destroy()
    listbox0.destroy()
    
    scrollbar=Scrollbar(f)
    scrollbar.pack(side=RIGHT,fill=Y)
    listbox0=Listbox(f,width=100)
    for i in range(0,len(rn)):
        n=i+1
        listbox0.insert(n,str(n)+') '+rn[i])

    listbox0.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=listbox0.yview)
             
   
        
    
    rid=[]
    rg=[]
    ra=[]
    ro=[]
    rp=[]

        
    for i in range(0,len(rn)):
        for j in range(2,int(cells)):
            if rn[i]==s.cell_value(j,1):
                rid.append(str(s.cell_value(j,0)))
                ra.append(str(s.cell_value(j,2)))
                rg.append(str(s.cell_value(j,3)))
                ro.append(str(s.cell_value(j,4)))
                rp.append(str(s.cell_value(j,5)))

    listbox0.pack()            
    listbox0.bind('<<ListboxSelect>>',currentalpha)
                
def price3(*event):
    global listbox0
    global ovlabel
    global idlabel
    global autlabel
    global genlabel
    global alpha
    global pricelabel

    
    selected=listbox0.get(ANCHOR)
        
    num=''
    for i in range(len(selected)):
            if selected[i]==')':
                break
            else:
                num+=selected[i]
    index=int(num)-1
             
        
    selected_id = pid[index]
    selected_aut = pa[index]
    selected_gen = pg[index]
    selected_ov = po[index]
    selected_price = pp[index]
        

    try:
        ovlabel.destroy()
    except:
        pass
    try:
        idlabel.destroy()
    except:
        pass
    try:
        genlabel.destroy()
    except:
        pass
    try:
        autlabel.destroy()
    except:
        pass
    try:
        pricelabel.destroy()
    except:
        pass
        
    idlabel=Label(show_w,text='Book id: '+selected_id,font=('courier new greek',20,'bold'))
    idlabel.place(relx=0.08,rely=0.65)
    autlabel=Label(show_w,text='Author: '+selected_aut,font=('courier new greek',20,'bold'))
    autlabel.place(relx=0.08,rely=0.7)
    genlabel=Label(show_w,text='Genre: '+selected_gen,font=('courier new greek',20,'bold'))
    genlabel.place(relx=0.08,rely=0.75)
    ovlabel=Label(show_w,text='Overview: '+selected_ov,font=('courier new greek',20,'bold'))
    ovlabel.place(relx=0.08,rely=0.8)

def price2():
    global scrollbar
    global ab
    global pb
    global listbox0
    global asp
    global aep
    global pn
    global pid
    global pp
    global pg
    global po
    global pa
    asp=sp.get()
    aep=ep.get()


    pn=[]
    pid=[]
    pa=[]
    pg=[]
    po=[]
    pp=[]
    for i in range(2,int(cells)):
        if (s.cell_value(i,5))>=float(asp) and (s.cell_value(i,5))<=float(aep):
                pn.append(s.cell_value(i,1))
                pid.append(str(s.cell_value(i,0)))
                pa.append(str(s.cell_value(i,2)))
                pg.append(str(s.cell_value(i,3)))
                po.append(str(s.cell_value(i,4)))
                pp.append(str(s.cell_value(i,5)))

    scrollbar.destroy()
    listbox0.destroy()
    
    scrollbar=Scrollbar(f)
    scrollbar.pack(side=RIGHT,fill=Y)
    listbox0=Listbox(f,width=100)
    for i in range(len(pn)):
        listbox0.insert(i+1,str(i+1)+') '+pn[i])

    listbox0.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=listbox0.yview)
    listbox0.pack() 
    listbox0.bind('<<ListboxSelect>>',price3)


def price():
    global sp
    global ep
    sp=StringVar()
    ep=StringVar()

    ab.destroy()
    pb.destroy()
    gb.destroy()
    Label(show_w,text='Enter price range',font=('times',20),bg='cyan').place(relx=0.7,rely=0.35)
    Entry(show_w,textvariable=sp).place(relx=0.72,rely=0.44)
    Label(show_w,text='to',bg='pink').place(relx=0.75,rely=0.5)
    Entry(show_w,textvariable=ep).place(relx=0.72,rely=0.56)

    
    Button(show_w,text='Filter',command=price2).place(relx=0.74,rely=0.6)

def author3(*event2):
    global listbox0
    global ovlabel
    global idlabel
    global autlabel
    global genlabel
    global pricelabel

    selected=listbox0.get(ANCHOR)
        
    num=''
    for i in range(len(selected)):
            if selected[i]==')':
                break
            else:
                num+=selected[i]
    index=int(num)-1
             
        
    selected_id = aid[index]
    selected_aut = alist[index]
    selected_gen = ag[index]
    selected_ov = ao[index]
    selected_price = ap[index]
        

    try:
        ovlabel.destroy()
    except:
        pass
    try:
        idlabel.destroy()
    except:
        pass
    try:
        genlabel.destroy()
    except:
        pass
    try:
        autlabel.destroy()
    except:
        pass
    try:
        pricelabel.destroy()
    except:
        pass
        
    idlabel=Label(show_w,text='Book id: '+selected_id,font=('courier new greek',20,'bold'))
    idlabel.place(relx=0.08,rely=0.65)
    autlabel=Label(show_w,text='Author: '+selected_aut,font=('courier new greek',20,'bold'))
    autlabel.place(relx=0.08,rely=0.7)
    genlabel=Label(show_w,text='Genre: '+selected_gen,font=('courier new greek',20,'bold'))
    genlabel.place(relx=0.08,rely=0.75)
    ovlabel=Label(show_w,text='Overview: '+selected_ov,font=('courier new greek',20,'bold'))
    ovlabel.place(relx=0.08,rely=0.8)

    
def author2(value):
    global scrollbar
    global listbox0
    global alist
    global an
    global aid
    global ag
    global ao
    global ap
    aname=value
    alist=[]
    an=[]
    aid=[]
    ag=[]
    ao=[]
    ap=[]
    for i in range(2,int(cells)):
        if s.cell_value(i,2)==aname:
            alist.append(s.cell_value(i,2))
            an.append(s.cell_value(i,1))
            aid.append(str(s.cell_value(i,0)))
            ag.append(str(s.cell_value(i,3)))
            ao.append(str(s.cell_value(i,4)))
            ap.append(str(s.cell_value(i,5)))

    listbox0.destroy()
    scrollbar.destroy()
    scrollbar=Scrollbar(f)
    scrollbar.pack(side=RIGHT,fill=Y)
    listbox0=Listbox(f,width=100)
    for i in range(len(an)):
        listbox0.insert(i+1,str(i+1)+') '+an[i])

    listbox0.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=listbox0.yview)
    listbox0.pack()    
    listbox0.bind('<<ListboxSelect>>',author3)
            
def author():
    global avariable

    ab.destroy()
    pb.destroy()
    gb.destroy()
        
    rb=xlrd.open_workbook('Book details.xls')
    s=rb.sheet_by_index(0)
    authors=[]
    avariable=StringVar()
    avariable.set('Select an Author')
    for i in range(2,int(cells)):
        authors.append(s.cell_value(i,2))
    om=OptionMenu(show_w, avariable, *authors, command=author2)
    om.place(relx=0.7,rely=0.41)
    om.configure(bg='cyan',font=('eras medium itc',20,'bold'))
    

def sback():
    show_w.destroy()
    opt()
def show():
        global scrollbar
        global listbox0
        global s
        global show_w
        global dn
        global f
        global cells
        global ab
        global pb
        global gb

        try:
            root.destroy()
        except:
            pass
        show_w = Tk()
        show_w.state('zoomed')

        canvas = Canvas(width = 1000, height = 500, bg = 'blue')
        canvas.pack(expand = YES, fill = BOTH)

        image = ImageTk.PhotoImage(file = path+"/ls2.jpg")
        canvas.create_image(-350, -250, image = image, anchor = NW)

        Label(canvas,text='Library',font=('dfkai-sb',85),fg='black',bg='orange').place(relx=0.4,rely=0.05)
        f=Frame(show_w)
        f.place(relx=0.07,rely=0.3)
        
        scrollbar=Scrollbar(f)
        scrollbar.pack(side=RIGHT,fill=Y)
        rb=xlrd.open_workbook('Book details.xls')
            
            
        s=rb.sheet_by_index(0)
        cell=open('cell.txt','r')
        cells=cell.read()
        cell.close()
        
        Label(f,text='Book List').pack()
        dn={}
        di={}
        da={}
        dg={}
        do={}

                
        for i in range(2,int(cells)):
                dn.update({i:s.cell_value(i,1)})

        for i in range(2,int(cells)):
                da.update({i:s.cell_value(i,2)})

        for i in range(2,int(cells)):
                di.update({i:s.cell_value(i,0)})
        for i in range(2,int(cells)):
                dg.update({i:s.cell_value(i,3)})
        for i in range(2,int(cells)):
                do.update({i:s.cell_value(i,4)})

        listbox0=Listbox(f,width=100)     
            
        for i in range(2,int(cells)):
            n=i-1
            listbox0.insert(n,str(n)+') '+dn[i])
        
        listbox0.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox0.yview)    
        listbox0.pack()
        
        
        

        listbox0.bind('<<ListboxSelect>>',current)
        ab=Button(canvas,text='Arrange alphabetically',font=('eras medium itc',20,'bold'),bg='cyan',fg='black',command=alpha)
        ab.place(relx=0.7,rely=0.31)
        pb=Button(canvas,text='Filter by Price',font=('eras medium itc',20,'bold'),bg='cyan',fg='black',command=price)
        pb.place(relx=0.7,rely=0.41)
        gb=Button(canvas,text='Filter by Author',font=('eras medium itc',20,'bold'),bg='cyan',fg='black',command=author)
        gb.place(relx=0.7,rely=0.51)
        Button(show_w,text='Back',font=('segoe script',20),fg='dark green',bg='yellow',command=sback).place(relx=0.88,rely=0.88)
    
        
        show_w.mainloop()

def issue3():
    rb=xlrd.open_workbook('Book details.xls')
    s=rb.sheet_by_index(0)
    nwb=copy(rb)
    sheet=nwb.get_sheet(0)
    cell=open('cell.txt','r')
    cells=cell.read()
    cell.close()

    cur_date=time.localtime()
    f_date=str(cur_date[2])+'/'+str(cur_date[1])+'/'+str(cur_date[0])
    for i in range(2,int(cells)):
        if s.cell_value(i,1)==ib:
            sheet.write(i,6,'ISSUED')
            sheet.write(i,7,f_date)
            sheet.write(i,8,user)
            
    os.remove('Book details.xls')
    nwb.save('Book details.xls')
    messagebox.showinfo('Book Issued',"You just issued the book '"+ib+"'")
    issuew.destroy()
    opt()

def issue2(issue_value):
    global ib
    ib=issue_value
    Button(issuew,text='Issue this book',command=issue3,font=('courier',30,'bold'),bg='sky Blue',fg='dark green').place(relx=0.45,rely=0.5)
    
def iback():
    issuew.destroy()
    opt()
def issue():
    global issuew
    global cells
    root.destroy()
    rb=xlrd.open_workbook('Book details.xls')
    s=rb.sheet_by_index(0)
    cell=open('cell.txt','r')
    cells=cell.read()
    cell.close()
    issuew=Tk()
    issuew.state('zoomed')
    issuew.title('BookWiki  (Issue Book)')

    canvas = Canvas(width = 1000, height = 500, bg = 'blue')
    canvas.pack(expand = YES, fill = BOTH)

    image = ImageTk.PhotoImage(file = path+"/li2.jpg")
    canvas.create_image(0, 0, image = image, anchor = NW)

    Label(canvas,text='Issue a Book',font=('book antiqua',30),bg='orange').pack()
    
    issue_list=[]
    variable=StringVar()
    variable.set('Select a book to issue')
    for i in range(2,int(cells)):
        if s.cell_value(i,6)!='ISSUED':
            issue_list.append(s.cell_value(i,1))
    if len(issue_list)==0:
        Label(canvas,text='All books are currently issued').place(relx=0.3,rely=0.3)
    else:
        om=OptionMenu(issuew, variable, *issue_list, command=issue2)
        om.place(relx=0.3,rely=0.3)
        om.configure(font=('pristina',28,'bold'),bg='pink')
    Button(issuew,text='Back',font=('segoe script',20),fg='dark green',bg='yellow',command=iback).place(relx=0.8,rely=0.8)
    issuew.mainloop()

def rback():
    retw.destroy()
    opt()
def ret3():
    rb=xlrd.open_workbook('Book details.xls')
    s=rb.sheet_by_index(0)
    nwb=copy(rb)
    sheet=nwb.get_sheet(0)
    cell=open('cell.txt','r')
    cells=cell.read()
    cell.close()
    
    for i in range(2,int(cells)):
        if s.cell_value(i,1)==rbn:
            sheet.write(i,6,'AVAILABLE')
                                    
    os.remove('Book details.xls')
    nwb.save('Book details.xls')
    retw.destroy()
    opt()
def ret2(return_value):
    global rbn
    rbn=return_value
    confirm = messagebox.askquestion('Return',"Are you sure you want to return the book '"+rbn+"' ?")
    if confirm=='yes':
        ret3()
    
def ret():
    global retw
    global cells
    root.destroy()
    rb=xlrd.open_workbook('Book details.xls')
    s=rb.sheet_by_index(0)
    cell=open('cell.txt','r')
    cells=cell.read()
    cell.close()
    retw=Tk()
    retw.state('zoomed')

    canvas = Canvas(width = 10, height = 400, bg = 'blue')
    canvas.pack(expand = YES, fill = BOTH)

    image = ImageTk.PhotoImage(file = path+"\\ret.jpg")
        
    canvas.create_image(-100, -100, image = image, anchor = NW)

    Label(canvas,text='Return Book',font=('courier new',50,'bold'),bg='yellow').place(relx=0.38,rely=0.1)
    return_list=[]
    variable=StringVar()
    variable.set('Select a book to return')
    for i in range(2,int(cells)):
        if s.cell_value(i,6)=='ISSUED' and s.cell_value(i,8)==user:
            return_list.append(s.cell_value(i,1))
    if len(return_list)==0:
        Label(retw,text='You have no issued books!',font=('courier new',30,'bold'),bg='pink').place(relx=0.3,rely=0.4)
    else:
        om=OptionMenu(canvas, variable, *return_list, command=ret2)
        om.place(relx=0.4,rely=0.5)
        om.configure(font=('pristina',28,'bold'),bg='pink')
    Button(retw,text='Back',font=('segoe script',20),fg='dark green',bg='yellow',command=rback).place(relx=0.8,rely=0.8)
    retw.mainloop()    
    
def stback():
    stockw.destroy()
    opt()

def stock():
    global stockw

    if user=='admin':
        
        root.destroy()
        rb=xlrd.open_workbook('Book details.xls')
        s=rb.sheet_by_index(0)
        cell=open('cell.txt','r')
        cells=cell.read()
        cell.close()

        stockw=Tk()
        stockw.state('zoomed')

        canvas = Canvas(width = 10, height = 400, bg = 'blue')
        canvas.pack(expand = YES, fill = BOTH)

        image = ImageTk.PhotoImage(file = path+"\\stock.jpg")
            
        canvas.create_image(0, 0, image = image, anchor = NW)

        Label(canvas,text='STOCK',font=('viner hand itc',38,'bold'),bg='orange').pack()

        f1=Frame(canvas)
        f2=Frame(canvas)
        
        
        s1=Scrollbar(f1)
        s1.pack(side=RIGHT,fil=Y)
        s2=Scrollbar(f2)
        s2.pack(side=RIGHT,fil=Y)


        Label(canvas,text='Issued books',font=('snap itc',20),bg='cyan').place(relx=0.25,rely=0.2)
        f1.place(relx=0.2,rely=0.25)
        
        listbox1=Listbox(f1,width=100)

        n=0
        for i in range(2,int(cells)):
            if s.cell_value(i,6)=='ISSUED':
                n+=1
                listbox1.insert(n,str(n)+') '+s.cell_value(i,1)+' by '+s.cell_value(i,2)+'   (Issued by '+s.cell_value(i,8)+' on '+s.cell_value(i,7)+')')

        listbox1.config(yscrollcommand=s1.set)
        s1.config(command=listbox1.yview)
        listbox1.pack()
        
        
        Label(canvas,text='Available books',font=('snap itc',20),bg='cyan').place(relx=0.25,rely=0.5)
        f2.place(relx=0.2,rely=0.55)
        listbox2=Listbox(f2,width=100)

        n2=0
        for i in range(2,int(cells)):
            if s.cell_value(i,6)=='AVAILABLE' and s.cell_value(i,8)!='':
                n2+=1
                listbox2.insert(n2,str(n2)+') '+s.cell_value(i,1)+' by '+s.cell_value(i,2)+'   (Last Issued by '+s.cell_value(i,8)+' on '+s.cell_value(i,7)+')')
            elif s.cell_value(i,6)=='AVAILABLE':
                n2+=1
                listbox2.insert(n2,str(n2)+') '+s.cell_value(i,1)+' by '+s.cell_value(i,2)+'   (not issued yet)')

        listbox2.config(yscrollcommand=s2.set)
        s2.config(command=listbox2.yview)
        listbox2.pack()

        
        books=str(int(cells)-2)
        Label(canvas,text='Total books:'+books,font=('dfgothic-eb',30),bg='pink').place(relx=0.4,rely=0.8)
        Button(canvas,text='Back',font=('segoe script',20),fg='dark green',bg='yellow',command=stback).place(relx=0.8,rely=0.8)
        stockw.mainloop()
    else:
        messagebox.showinfo('Access denied','You do not have access to admin features')

def opt():
        
        
        global root
        global quotes
        root=Tk()
        root.title('Book Wiki  (Home)')
        root.state('zoomed')

        quotes={'Frank Zappa': '“So many books, so little time.”', 'Marcus Tullius Cicero': '“A room without books is like a body without a soul.”', ' Jane Austen, Northanger Abbey': '“The person, be it gentleman or lady, who has not pleasure in a good novel, must be intolerably stupid.”', 'Mark Twain': '“Good friends, good books, and a sleepy conscience: this is the ideal life.”', 'Neil Gaiman, Coraline': '“Fairy tales are more than true: not because they tell us that dragons exist, but because they tell us that dragons can be beaten.”', 'Jorge Luis Borges': '“I have always imagined that Paradise will be a kind of library.”', 'Lemony Snicket':'“Never trust anyone who has not brought a book with them.”'} 
        quoter=random.choice(list(quotes.keys()))
        quote=quotes[quoter]

        a=time.localtime()
        

        canvas = Canvas(width = 10, height = 400, bg = 'blue')
        canvas.pack(expand = YES, fill = BOTH)

        image = ImageTk.PhotoImage(file = path+"\l2.png")
        
        canvas.create_image(0, 0, image = image, anchor = NW)
        

        title= Label(root,text='Welcome to BookWiki',font=('jokerman',50),bg='brown')
        subtitle=Label(root,text='Your Library Managemrnt Software',font=('dfgothic-eb',20),bg='yellow')
        title.place(relx=0.24,rely=0.08)
        subtitle.place(relx=0.33,rely=0.25)
        
        o1=Button(root,text='Add books',font=('Arial Body',20),command=add)
        
        o2=Button(root,text='Remove books',font=('Arial Body',20),command=remove1)
        
        o3=Button(root,text='Show books',font=('Arial Body',20),command=show)
        
        o4=Button(root,text='Issue',font=('arial body',20),command=issue)
        
        o5=Button(root,text='Return',font=('arial body',20),command=ret)
        
        o6=Button(root,text='Stock',font=('arial body',20),command=stock)

        
        o3.place(relx=0.7,rely=0.49)
        o5.place(relx=0.4,rely=0.65)
        o4.place(relx=0.09,rely=0.65)
        o1.place(relx=0.09,rely=0.49)
        o2.place(relx=0.37,rely=0.49)
        o6.place(relx=0.7,rely=0.65)

        Label(canvas,text='~'+quoter,font=('comic sans ms',12,'bold'),bg='pink',fg='black').pack(side='bottom')
        Label(canvas,text=quote,font=('tempus sans itc',15,'bold'),bg='orange',fg='black').pack(side='bottom')

        root.mainloop()


        
       

login()
