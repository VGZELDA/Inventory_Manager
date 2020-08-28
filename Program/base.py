import tkinter as tk
import pandas as pd
from datetime import datetime
from tabulate import tabulate
from datetime import timedelta
from ttkwidgets import autocomplete
import glob
import os

HEIGHT=1080
WIDTH=1920
months=["NONE","JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"]
labelfont=("Fixedsys",16)
fontExample = ("Fixedsys", 13)
today,yesterday='',''

def decide_date(x):
    global yesterday
    global today
    global dateroot
    yesterday=str(dayentry.get())+'-'+str(monthentry.get())+'-'+str(yearentry.get())
    today=str(datetime.strptime(yesterday,'%d-%m-%Y')+timedelta(days=1)).split()[0]
    today=today.split('-')
    today=today[2]+'-'+today[1]+'-'+today[0]
    dateroot.destroy()
    
def dfo1(x):
    monthentry.focus_set()
    
def dfo2(x):
    yearentry.focus_set()
    
dateroot=tk.Tk()
datecanvas=tk.Canvas(dateroot,width=500,height=150,bg="#FBFBFB")
datecanvas.pack()
daylabel=tk.Label(dateroot,text="DAY (DD)",anchor="w",bg="#FBFBFB",fg="#4C3822")
daylabel.place(relx=.1,rely=.3,relheight=.2,relwidth=.2)
daylabel.config(font=labelfont)
dayentry=tk.Entry(dateroot,font=fontExample,highlightthickness=3)
dayentry.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
dayentry.bind("<Return>",dfo1)
dayentry.place(relx=.1,rely=.7,relwidth=.1)
monthlabel=tk.Label(dateroot,text="MONTH (MM)",anchor="w",bg="#FBFBFB",fg="#4C3822")
monthlabel.place(relx=.32,rely=.3,relheight=.2,relwidth=.2)
monthlabel.config(font=labelfont)
monthentry=tk.Entry(dateroot,font=fontExample,highlightthickness=3)
monthentry.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
monthentry.bind("<Return>",dfo2)
monthentry.place(relx=.36,rely=.7,relwidth=.1)
yearlabel=tk.Label(dateroot,text="YEAR",anchor="w",bg="#FBFBFB",fg="#4C3822")
yearlabel.place(relx=.61,rely=.3,relheight=.2,relwidth=.2)
yearlabel.config(font=labelfont)
yearentry=tk.Entry(dateroot,font=fontExample,highlightthickness=3)
yearentry.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
yearentry.bind("<Return>",decide_date)
yearentry.place(relx=.6,rely=.7,relwidth=.1)
yearentry.focus_set()

list_of_files = glob.glob('Stock/*') # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)[6:16].split('-')
dayentry.insert(0,latest_file[0])
monthentry.insert(0,latest_file[1])
yearentry.insert(0,latest_file[2]) 

tk.mainloop()

daybeforeyesterday=str(datetime.strptime(yesterday,'%d-%m-%Y')+timedelta(days=-1)).split()[0]
daybeforeyesterday=daybeforeyesterday.split('-')
daybeforeyesterday=daybeforeyesterday[2]+'-'+daybeforeyesterday[1]+'-'+daybeforeyesterday[0]

def printsheet():
    try:
        rejdf_copy=pd.read_excel("Rejects\\"+yesterday+"-REJECTS.xlsx")
    except:
        rejdf_copy=pd.DataFrame([["","","",0]],columns=["R E","J E","C T","E D"])
    rejdf_copy['E D']=rejdf_copy["E D"].apply(float)
    toprint=pd.read_excel("Stock\\"+yesterday+".xlsx")   
    totalin=sum(toprint['WEIGHT_IN'])
    totalout=sum(toprint['WEIGHT_OUT'])
    totalsale=sum(toprint['SALE'])
    toprint=toprint.drop('SALE',axis=1)
    toprint=toprint.drop('WEIGHT_IN',axis=1)
    toprint=toprint.drop('WEIGHT_OUT',axis=1)
    toprint.columns=['SIZE','PARTY','Grade','STOCK']
    ingot,billet,bloom=0,0,0
    rejdf_copy=rejdf_copy[rejdf["E D"]>0]
    Entries=rejdf_copy.values.tolist()+toprint.values.tolist()
    for i in(Entries):
        s1,s2=i[0].split('x')
        s1,s2=float(s1),float(s2)
        if(s1<76)and(s1!=s2):
            ingot=ingot+i[3]
        elif(s1==s2):
            billet=billet+i[3]
        else:
            bloom=bloom+i[3]
    total_STOCK=billet+bloom+ingot
    total_entries=len(toprint)
    toprint = toprint[toprint['STOCK'] >= 0.001]
    entries=toprint.values.tolist()
    s1=entries[(total_entries//2)-1][0].split('x')[0]
    for e in range((total_entries//2)-1,total_entries):
        if(entries[e][0].split('x')[0]==s1):
            final=e+1
            continue
        else:
            final=e
            break
    x=rejdf_copy.values.tolist()
    x.insert(0,["R E","J E","C T","E D"])
    x4=pd.DataFrame([[ingot,billet,bloom,total_STOCK],["IN","PRODUCTION","SALE"],[totalin,totalout,totalsale]],columns=["INGOTS","BILLETS","BLOOMS","TOTAL STOCK"])
    x1=pd.DataFrame([["INGOTS","BILLETS","BLOOMS","TOTAL STOCK"]]+x4.values.tolist()+[['SIZE','PARTY','Grade','STOCK']]+entries[0:final],columns=["DATE -->",today.split('-')[0],months[int(today.split('-')[1])],today.split('-')[2]])
    x3=pd.concat([pd.DataFrame(entries[final:],columns=['size','party','grade','stock']),pd.DataFrame(x,columns=['size','party','grade','stock'])],ignore_index=True)
    df1=pd.concat([x1,x3],axis=1)
    df1.to_excel("Print\\"+today+" - print.xlsx",index=False)
    printdf=pd.read_excel("Print\\"+today+" - print.xlsx")
    printdf=printdf.fillna("")
    printdf.to_html("Print\\"+today+" - print.html",index=False)

def floatmod(num):
    num=float(num)
    if(num-int(num)==0):
        return int(num)
    else:
        return num

def findrelevant(x):
    PARTYin=str(comp_drop.get())
    SIZEin=str(SIZE_drop.get())
    GRADEin=str(qual_drop.get())
    if(len(PARTYin)==0):
        PARTYin=''
    if(len(SIZEin)==0):
        SIZEin=''
    if(len(GRADEin)==0):
        GRADEin=''
    global df
    matched=[]
    matchwith=df.values.tolist()
    for i in matchwith:
        if(i[0][:len(SIZEin)]==SIZEin)and(i[1][:len(PARTYin)].upper()==PARTYin.upper())and(i[2][:len(GRADEin)].upper()==GRADEin.upper()):
            matched.append(i)
    matched=pd.DataFrame(matched,columns=df.columns)
    matched.insert(loc=0, column='##', value=[kl+1 for kl in range(len(matched))])
    tab_label.delete("1.0", "end")
    tab_label.insert(tk.END,str(tabulate(matched, headers='keys', tablefmt='fancy_grid',showindex="never")))

def selection(x):
    try:
        index_to_fill=int(tab_label.selection_get())-1
        PARTYin=str(comp_drop.get())
        SIZEin=str(SIZE_drop.get())
        GRADEin=str(qual_drop.get())
        SIZE_drop.delete(first=0,last=100)
        comp_drop.delete(first=0,last=100)
        qual_drop.delete(first=0,last=100)
        if(len(PARTYin)==0):
            PARTYin=''
        if(len(SIZEin)==0):
            SIZEin=''
        if(len(GRADEin)==0):
            GRADEin=''
        global df
        matched=[]
        matchwith=df.values.tolist()
        for i in matchwith:
            if(i[0][:len(SIZEin)]==SIZEin)and(i[1][:len(PARTYin)].upper()==PARTYin.upper())and(i[2][:len(GRADEin)].upper()==GRADEin.upper()):
                matched.append(i)
        matched=pd.DataFrame(matched,columns=df.columns)
        matched.insert(loc=0, column='##', value=[kl+1 for kl in range(len(matched))])
        index_to_fill=int(tab_label.selection_get())-1
        SIZE_drop.insert(0,matched["SIZE"][index_to_fill])
        comp_drop.insert(0,matched["PARTY"][index_to_fill])
        qual_drop.insert(0,matched["GRADE"][index_to_fill])
        findrelevant(0)
    except:
        print("",end="")
SIZE=pd.read_excel("Program\\"+"SIZE.xlsx")
SIZE['SIZE']=SIZE['SIZE'].apply(str)
sorted_SIZE=[i.split('x') for i in SIZE['SIZE'].values.tolist()]
for i in range(len(sorted_SIZE)):
    for j in range(2):
        sorted_SIZE[i][j]=floatmod(sorted_SIZE[i][j])
sorted_SIZE=sorted(sorted_SIZE)
for i in range(len(sorted_SIZE)):
    sorted_SIZE[i]=str(sorted_SIZE[i][0])+'x'+str(sorted_SIZE[i][1])
SIZE['SIZE']=sorted_SIZE

comp=pd.read_excel("Program\\"+"PARTY.xlsx")
comp=comp.sort_values('PARTY')
comp = comp.reset_index(drop=True)
qual=pd.read_excel("Program\\"+"GRADE.xlsx")
qual['GRADE']=qual['GRADE'].apply(str)
qual=qual.sort_values('GRADE')
qual = qual.reset_index(drop=True)
try:
    df=pd.read_excel("Stock\\"+yesterday+".xlsx")
except:
    print("Yesterday's sheet not found.")
    
df['GRADE']=df['GRADE'].apply(str)
df['WEIGHT_OUT']=df['WEIGHT_OUT'].apply(float)
df['WEIGHT_IN']=df['WEIGHT_IN'].apply(float)
df['WEIGHT_OUT']=df['WEIGHT_OUT'].apply(float)
df['SALE']=df['SALE'].apply(float)


try:
    rejdf=pd.read_excel("Rejects\\"+yesterday+"-REJECTS.xlsx")
except:
    try:
        rejdf=pd.read_excel("Rejects\\"+daybeforeyesterday+"-REJECTS.xlsx")
        rejdf.to_excel("Rejects\\"+yesterday+"-REJECTS.xlsx",index=False)
    except:
        rejdf=pd.DataFrame([["","","",0]],columns=["R E","J E","C T","E D"])
        rejdf.to_excel("Rejects\\"+yesterday+"-REJECTS.xlsx",index=False)

def done():
    global root
    printsheet()
    root.destroy()


def reject():
    rejected=reject_entry.get()
    rejsize=SIZE_drop.get()
    rejparty=comp_drop.get()
    rejgrade=qual_drop.get()
    global df
    try:
        matched=df.loc[:,'SIZE':'GRADE'].values.tolist().index([str(rejsize),str(rejparty),str(rejgrade)])
    except:
        matched=-1
    if(matched!=-1):
        df.at[matched,'STOCK']=df.at[matched,'STOCK']-float(rejected)
        df = df.reset_index(drop=True)
    else:
        df=df.append({'PARTY':rejparty,'GRADE':rejgrade,'SIZE':rejsize,'WEIGHT_IN':0,'WEIGHT_OUT':0,'SALE':0,'STOCK':-1*float(rejected)},ignore_index=True)
    df.to_excel("Stock\\"+yesterday+".xlsx",index=False)
    nextdf=df.copy()
    nextdf['WEIGHT_IN']=[0]*len(df)
    nextdf['WEIGHT_OUT']=[0]*len(df)
    nextdf['SALE']=[0]*len(df)
    nextdf=nextdf[nextdf['STOCK']>=.001]
    nextdf.to_excel("Stock\\"+today+".xlsx",index=False)
    global rejdf
    rejdf['R E']=rejdf['R E'].apply(str)
    rejdf['J E']=rejdf['J E'].apply(str)
    rejdf['C T']=rejdf['C T'].apply(str)
    rejdf['E D']=rejdf['E D'].apply(float)
    try:
        ind=rejdf.loc[:,'R E':'C T'].values.tolist().index([str(rejsize),str(rejparty),str(rejgrade)])
    except:
        ind=-1
    if(ind==-1):
        rejdf=rejdf.append({"R E":rejsize,"J E":rejparty,"C T":rejgrade,"E D":float(rejected)},ignore_index=True)
    else:
        rejdf.at[ind,"E D"]=float(rejdf["E D"][ind]+float(rejected))
    rejdf=rejdf[rejdf["E D"]>=.001]
    rejdf.to_excel("Rejects\\"+yesterday+"-REJECTS.xlsx",index=False)
    SALE.delete(first=0,last=100)
    WEIGHT_OUT.delete(first=0,last=100)
    WEIGHT_IN.delete(first=0,last=100)
    SIZE_drop.delete(first=0,last=100)
    comp_drop.delete(first=0,last=100)
    qual_drop.delete(first=0,last=100)
    reject_entry.delete(first=0,last=100)
    SIZE_drop.focus_set()

def enter():
    R=reject_entry.get()
    if(len(R)!=0):
        reject()
    else:
        global SIZE
        global comp
        global qual
        global df
        global sorted_SIZE
        
        inPARTY=str(comp_drop.get()).upper()
        inSIZE=str(SIZE_drop.get())
        inSIZE=str(floatmod((inSIZE.split('x'))[0]))+'x'+str(floatmod((inSIZE.split('x'))[1]))
        inGRADE=str(qual_drop.get()).upper()
        inWEIGHT_IN=WEIGHT_IN.get()
        inWEIGHT_OUT=WEIGHT_OUT.get()
        inSALE=SALE.get()
        
        if(len(inWEIGHT_IN)==0):
            inWEIGHT_IN=0.0
        if(len(inWEIGHT_OUT)==0):
            inWEIGHT_OUT=0.0
        if(len(inSALE)==0):
            inSALE=0.0
            
        if((len(inPARTY)!=0) and (len(inSIZE)!=0) and (len(inGRADE)!=0)):
            if inPARTY not in list(comp['PARTY']):
                comp=comp.append({'PARTY':inPARTY},ignore_index=True)
                comp.to_excel("Program\\"+"PARTY.xlsx",index=False)
            if inSIZE not in list(SIZE['SIZE']):
                SIZE=SIZE.append({'SIZE':inSIZE},ignore_index=True)
                SIZE.to_excel("Program\\"+"SIZE.xlsx",index=False)
                
                SIZE=pd.read_excel("Program\\"+"SIZE.xlsx")
                SIZE['SIZE']=SIZE['SIZE'].apply(str)
                sorted_SIZE=[i.split('x') for i in SIZE['SIZE'].values.tolist()]
                for i in range(len(sorted_SIZE)):
                    for j in range(2):
                        sorted_SIZE[i][j]=floatmod(sorted_SIZE[i][j])
                sorted_SIZE=sorted(sorted_SIZE)
                for i in range(len(sorted_SIZE)):
                    sorted_SIZE[i]=str(sorted_SIZE[i][0])+'x'+str(sorted_SIZE[i][1])
                SIZE['SIZE']=sorted_SIZE
                SIZE.to_excel("Program\\"+"SIZE.xlsx",index=False)
    
            if inGRADE not in list(qual['GRADE']):
                qual=qual.append({'GRADE':inGRADE},ignore_index=True)
                qual.to_excel("Program\\"+"GRADE.xlsx",index=False)
            try:
                ind=df.loc[:,'SIZE':'GRADE'].values.tolist().index([str(inSIZE),str(inPARTY),str(inGRADE)])
            except:
                ind=-1
            
            if(ind==-1):
                inSIZE=str(floatmod((inSIZE.split('x'))[0]))+'x'+str(floatmod((inSIZE.split('x'))[1]))
                df=df.append({'PARTY':inPARTY,'GRADE':inGRADE,'SIZE':inSIZE,'WEIGHT_IN':floatmod(inWEIGHT_IN),'WEIGHT_OUT':floatmod(inWEIGHT_OUT),'SALE':floatmod(inSALE),'STOCK':(floatmod(inWEIGHT_IN)-floatmod(inWEIGHT_OUT)-floatmod(inSALE))},ignore_index=True)
                df.to_excel("Stock\\"+yesterday+".xlsx",index=False)
            else:
                if(floatmod(df['STOCK'][ind])+floatmod(inWEIGHT_IN)-floatmod(inWEIGHT_OUT)-floatmod(inSALE)>=-.001):
                    df['WEIGHT_IN']=df['WEIGHT_IN'].apply(float)
                    df['WEIGHT_OUT']=df['WEIGHT_OUT'].apply(float)
                    df['SALE']=df['SALE'].apply(float)
                    s=floatmod(df['STOCK'][ind])
                    df.at[ind,'WEIGHT_IN']=floatmod(inWEIGHT_IN)+floatmod(df['WEIGHT_IN'][ind])
                    df['WEIGHT_OUT'][ind]=floatmod(floatmod(inWEIGHT_OUT)+floatmod(df['WEIGHT_OUT'][ind]))
                    df.at[ind,'SALE']=floatmod(inSALE)+floatmod(df['SALE'][ind])
                    df.at[ind,'STOCK']=s+floatmod(inWEIGHT_IN)-floatmod(inWEIGHT_OUT)-floatmod(inSALE)
                    df.to_excel("Stock\\"+yesterday+".xlsx",index=False)
                else:
                    tab_label.delete("1.0", "end")
                    tab_label.insert(tk.END,"INVALID ENTRY!!! \n STOCK BECOMES NEGATIVE.\n PLEASE CHECK WEIGHT")
                    return
            df['A']=[floatmod(i.split('x')[0]) for i in df['SIZE']]
            df['B']=[floatmod(i.split('x')[1]) for i in df['SIZE']]
            df['WEIGHT_IN']=df['WEIGHT_IN'].apply(float)
            df['WEIGHT_OUT']=df['WEIGHT_OUT'].apply(float)
            df['SALE']=df['SALE'].apply(float)
            df=df.sort_values(by=['A','B','PARTY'])
            df=df.drop('A',axis=1)
            df=df.drop('B',axis=1)
            df = df.reset_index(drop=True)
            df.to_excel("Stock\\"+yesterday+".xlsx",index=False)
            
            nextdf=df.copy()
            nextdf['WEIGHT_IN']=[0]*len(df)
            nextdf['WEIGHT_OUT']=[0]*len(df)
            nextdf['SALE']=[0]*len(df)
            nextdf=nextdf[nextdf['STOCK']>=.001]
            nextdf.to_excel("Stock\\"+today+".xlsx",index=False)
            
            tab_label.delete("1.0", "end")
            tab_label.insert(tk.END,str(tabulate(df, headers='keys', tablefmt='fancy_grid',showindex="never")))
            SIZE_drop['completevalues']=list(SIZE['SIZE'])
            SIZE_drop.set_completion_list(list(SIZE['SIZE']))
            qual_drop['completevalues']=list(qual['GRADE'])
            qual_drop.set_completion_list(list(qual['GRADE']))
            comp_drop['completevalues']=list(comp['PARTY'])
            comp_drop.set_completion_list(list(comp['PARTY']))
            
            wintext.set("TOTAL IN \n "+str(sum(df["WEIGHT_IN"])))
            woutext.set("TOTAL OUT \n "+str(sum(df["WEIGHT_OUT"])))
            saletext.set("TOTAL SALE \n "+str(sum(df["SALE"])))
            
            SALE.delete(first=0,last=100)
            WEIGHT_OUT.delete(first=0,last=100)
            WEIGHT_IN.delete(first=0,last=100)
            SIZE_drop.delete(first=0,last=100)
            comp_drop.delete(first=0,last=100)
            qual_drop.delete(first=0,last=100)
            SIZE_drop.focus_set()
    
def enter_and_print():
    enter()
    printsheet()

# =============================================================================
# def recalculate():
#     global df
#     df=pd.read_excel("Stock\\"+yesterday+".xlsx")
#     recalculated=df.copy()
#     recalculated['WEIGHT_IN']=[0]*len(df)
#     recalculated['WEIGHT_OUT']=[0]*len(df)
#     recalculated['SALE']=[0]*len(df)
#     recalculated.to_excel("Stock\\"+today+".xlsx",index=False)
# =============================================================================

def show_rejtable():
    global rejdf
    tab_label.delete("1.0", "end")
    tab_label.insert(tk.END,str(tabulate(rejdf[rejdf["E D"]!=0], headers='keys', tablefmt='fancy_grid',showindex="never")))
    
def clearall():
    SALE.delete(first=0,last=100)
    WEIGHT_OUT.delete(first=0,last=100)
    WEIGHT_IN.delete(first=0,last=100)
    SIZE_drop.delete(first=0,last=100)
    comp_drop.delete(first=0,last=100)
    qual_drop.delete(first=0,last=100)
    findrelevant(0)
def next1(x):
    comp_drop.focus_set()
def next2(x):
    qual_drop.focus_set()
def next3(x):
    WEIGHT_IN.focus_set()
def next4(x):
    WEIGHT_OUT.focus_set()
def next5(x):
    SALE.focus_set()
def next6(x):
    reject_entry.focus_set()

root=tk.Tk()
root.attributes('-fullscreen', True)
root.bind("<KeyPress>",findrelevant)

canvas=tk.Canvas(root,width=WIDTH,height=HEIGHT,bg="#FBFBFB")
gif1 = tk.PhotoImage(file="Program\logo.png")
canvas.create_image(0,0, image=gif1,anchor='nw')
canvas.pack()

minimize=tk.Button(root, text = "-",font=fontExample,bg='red',fg='white', command = lambda: root.wm_state("iconic"))
minimize.place(relheight=.03,relwidth=.03,relx=.94)
close_button=tk.Button(root,text="X",bg='red',fg='white',command=done,font=fontExample)
close_button.place(relheight=.03,relwidth=.03,relx=.97)

today_label=tk.Label(root,fg="#4C3822",font=("Fixedsys",15,"bold underline"),text=yesterday,bg='#FBFBFB')
today_label.place(anchor='n',relx=.5,rely=1/540)

SIZE_label=tk.Label(root,text="Choose SIZE",anchor="w",bg="#FBFBFB",fg="#4C3822",font=labelfont)
SIZE_label.place(relx=.1,rely=.1,relheight=.05,relwidth=.15)

SIZE_drop=autocomplete.AutocompleteCombobox(root,completevalues=sorted_SIZE)
SIZE_drop.set_completion_list(sorted_SIZE)
SIZE_drop.bind("<Return>",next1)
root.option_add('*TCombobox*Listbox.font', fontExample)
SIZE_drop.place(relx=.1,rely=.15,relwidth=.14)

comp_label=tk.Label(root,text="Choose PARTY",anchor="w",bg="#FBFBFB",fg="#4C3822")
comp_label.place(relx=.1,rely=.2,relheight=.05,relwidth=.15)
comp_label.config(font=labelfont)

comp_drop=autocomplete.AutocompleteCombobox(root,completevalues=sorted(list(comp['PARTY'])))
comp_drop.set_completion_list(sorted(list(comp['PARTY'])))
comp_drop.bind("<Return>",next2)
root.option_add('*TCombobox*Listbox.font', fontExample)
comp_drop.place(relx=.1,rely=.25,relwidth=.14)

qual_label=tk.Label(root,text="Choose GRADE",anchor="w",bg="#FBFBFB",fg="#4C3822")
qual_label.place(relx=.1,rely=.3,relheight=.05,relwidth=.15)
qual_label.config(font=labelfont)

qual_drop=autocomplete.AutocompleteCombobox(root,completevalues=sorted(list(qual['GRADE'])))
qual_drop.set_completion_list(sorted(list(qual['GRADE'])))
qual_drop.bind("<Return>",next3)
root.option_add('*TCombobox*Listbox.font', fontExample)
qual_drop.place(relx=.1,rely=.35,relwidth=.14)

WEIGHT_IN_label=tk.Label(root,text="Enter <Weight In>",anchor="w",bg="#FBFBFB",fg="#4C3822")
WEIGHT_IN_label.place(relx=.1,rely=.4,relheight=.05,relwidth=.2)
WEIGHT_IN_label.config(font=labelfont)

WEIGHT_IN=tk.Entry(root,font=fontExample,highlightthickness=3)
WEIGHT_IN.bind("<Return>",next4)
WEIGHT_IN.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
WEIGHT_IN.place(relx=.1,rely=.45)

WEIGHT_OUT_label=tk.Label(root,text="Enter <Weight Out>",anchor="w",bg="#FBFBFB",fg="#4C3822")
WEIGHT_OUT_label.place(relx=.1,rely=.5,relheight=.05,relwidth=.2)
WEIGHT_OUT_label.config(font=labelfont)

WEIGHT_OUT=tk.Entry(root,font=fontExample,highlightthickness=3)
WEIGHT_OUT.bind("<Return>",next5)
WEIGHT_OUT.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
WEIGHT_OUT.place(relx=.1,rely=.55)

SALE_label=tk.Label(root,text="Enter SALE",anchor="w",bg="#FBFBFB",fg="#4C3822")
SALE_label.place(relx=.1,rely=.6,relheight=.05,relwidth=.2)
SALE_label.config(font=labelfont)

SALE=tk.Entry(root,font=fontExample,highlightthickness=3)
SALE.bind("<Return>",next6)
SALE.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
SALE.place(relx=.1,rely=.65)

printable=df.copy()
printable.insert(loc=0, column='##', value=[i+1 for i in range(len(printable))])
tab_label=tk.Text(root,fg="#458BC6",bg="#FBFBFB",font=("Courier", 13,"bold"),highlightthickness=2)
tab_label.config(highlightbackground = "black", highlightcolor="black")
tab_label.insert(tk.END,str(tabulate(printable, headers='keys', tablefmt='fancy_grid',showindex="never")))
tab_label.place(relx=.25,rely=.1,relheight=.65,relwidth=.75 )
tab_label.bind("<Button-1>", selection)

reject_label=tk.Label(root,text="Enter Reject",anchor="w",bg="#FBFBFB",fg="#4C3822")
reject_label.place(relx=.1,rely=.7,relheight=.05,relwidth=.15)
reject_label.config(font=labelfont)

reject_entry=tk.Entry(root,font=fontExample,highlightthickness=3)
reject_entry.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
reject_entry.place(relx=.1,rely=.75)

enter_button=tk.Button(root,text="Enter",font=("Fixedsys",13),command=enter_and_print,bg="#458BC6")
enter_button.place(relx=.04,rely=.69)

enter_button2=tk.Button(root,text="Enter",font=("Fixedsys",13),command=enter_and_print,bg="#458BC6")
enter_button2.place(relx=.04,rely=.49)

clear_button=tk.Button(root,text="Clear",font=("Fixedsys",13),command=clearall,bg="#458BC6")
clear_button.place(relx=.04,rely=.25)

show_reject=tk.Button(root,text="Show REJECTS",font=("Fixedsys",13),command=show_rejtable,bg="#458BC6")
show_reject.place(relx=.25,rely=.055)

wintext=tk.StringVar()
wintext.set("TOTAL IN \n "+str(sum(df["WEIGHT_IN"])))
display_win=tk.Label(root,textvariable=wintext,anchor="w",bg="#FBFBFB",fg="#4C3822")
display_win.place(relx=.45,rely=.77,relheight=.05,relwidth=.15)
display_win.config(font=labelfont)

woutext=tk.StringVar()
woutext.set("TOTAL OUT \n "+str(sum(df["WEIGHT_OUT"])))
display_wout=tk.Label(root,textvariable=woutext,anchor="w",bg="#FBFBFB",fg="#4C3822")
display_wout.place(relx=.55,rely=.77,relheight=.05,relwidth=.15)
display_wout.config(font=labelfont)

saletext=tk.StringVar()
saletext.set("TOTAL SALE \n "+str(sum(df["SALE"])))
display_sale=tk.Label(root,textvariable=saletext,anchor="w",bg="#FBFBFB",fg="#4C3822")
display_sale.place(relx=.65,rely=.77,relheight=.05,relwidth=.15)
display_sale.config(font=labelfont)
#print_button=tk.Button(root,text="Print",font=("Fixedsys",13),command=printsheet,bg="#458BC6")
#print_button.place(relx=.25,rely=.8)

#recalculate_button=tk.Button(root,text="Recalculate",font=("Fixedsys",13),command=recalculate,bg="#458BC6")
#recalculate_button.place(relx=.9,rely=.8)

#reject_button=tk.Button(root,text="Reject",font=("Fixedsys",13),command=reject,bg="#458BC6")
#   reject_button.place(relx=.04,rely=.446)

tk.mainloop()
