import tkinter as tk
import pandas as pd
from datetime import datetime
from datetime import timedelta
from ttkwidgets import autocomplete
from tksheet import Sheet
import glob
import os
import sys
def restart_program():
    """Restarts the current program.
    Note: this function does not return. Any cleanup action (like
    saving data) must be done before calling this function."""
    python = sys.executable
    os.execl(python, python, * sys.argv)

HEIGHT=1080
WIDTH=1920
rejopen=0
fromhigh=0
tohigh=0
# counter=0
transfer_open=0
everyone_state=1
months=["NONE","JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"]
labelfont=("Fixedsys",16)
fontExample = ("Fixedsys", 11)
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
    global df
    rejdf_copy=df[df['REJECTED']>0]
    del rejdf_copy['WEIGHT_IN']
    del rejdf_copy['WEIGHT_OUT']
    del rejdf_copy['SALE']
    del rejdf_copy['STOCK']
    del rejdf_copy['OPENING']
    rejdf_copy['REJECTED']=rejdf_copy["REJECTED"].apply(float)
    toprint=pd.read_excel("Stock\\"+yesterday+".xlsx")   
    toprint=toprint[toprint["STOCK"]>=.001]
    totalin=sum(toprint['WEIGHT_IN'])
    totalout=sum(toprint['WEIGHT_OUT'])
    totalsale=sum(toprint['SALE'])
    toprint=toprint.drop('SALE',axis=1)
    toprint=toprint.drop('WEIGHT_IN',axis=1)
    toprint=toprint.drop('WEIGHT_OUT',axis=1)
    toprint=toprint.drop('OPENING',axis=1)
    toprint=toprint.drop('REJECTED',axis=1)
    toprint.columns=['SIZE','PARTY','Grade','STOCK']
    ingot,billet,bloom=0,0,0
    rejdf_copy=rejdf_copy[rejdf_copy['REJECTED']>0]
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
    x.insert(0,["REJECTED","REJECTED","REJECTED","REJECTED"])
    x4=pd.DataFrame([[ingot,billet,bloom,total_STOCK],["DATE -->",yesterday.split('-')[0],months[int(yesterday.split('-')[1])],yesterday.split('-')[2]],["IN","PRODUCTION","SALE"],[totalin,totalout,totalsale]],columns=["INGOTS","BILLETS","BLOOMS","OPENING STOCK"])
    x1=pd.DataFrame([["INGOTS","BILLETS","BLOOMS","OPENING STOCK"]]+x4.values.tolist()+[['SIZE','PARTY','Grade','STOCK']]+entries[0:final],columns=["DATE -->",today.split('-')[0],months[int(today.split('-')[1])],today.split('-')[2]])
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
    matched['WEIGHT_IN']=matched['WEIGHT_IN'].apply(lambda x: round(x, 3))
    matched['WEIGHT_OUT']=matched['WEIGHT_OUT'].apply(lambda x: round(x, 3))
    matched['SALE']=matched['SALE'].apply(lambda x: round(x, 3))
    matched['REJECTED']=matched['REJECTED'].apply(lambda x: round(x, 3))
    matched['OPENING']=matched['OPENING'].apply(lambda x: round(x, 3))
    matched['STOCK']=matched['STOCK'].apply(lambda x: round(x, 3))
    global sheet
    sheet.set_sheet_data(data=matched.values.tolist())
    global rejopen
    rejopen=0
# def selection(x):
#     try:
        # index_to_fill=int(tab_label.selection_get())-1
        # PARTYin=str(comp_drop.get())
        # SIZEin=str(SIZE_drop.get())
        # GRADEin=str(qual_drop.get())
        # SIZE_drop.delete(first=0,last=100)
        # comp_drop.delete(first=0,last=100)
        # qual_drop.delete(first=0,last=100)
        # if(len(PARTYin)==0):
        #     PARTYin=''
        # if(len(SIZEin)==0):
        #     SIZEin=''
        # if(len(GRADEin)==0):
        #     GRADEin=''
        # global df
        # matched=[]
        # matchwith=df.values.tolist()
        # for i in matchwith:
        #     if(i[0][:len(SIZEin)]==SIZEin)and(i[1][:len(PARTYin)].upper()==PARTYin.upper())and(i[2][:len(GRADEin)].upper()==GRADEin.upper()):
        #         matched.append(i)
        # matched=pd.DataFrame(matched,columns=df.columns)
        # matched.insert(loc=0, column='##', value=[kl+1 for kl in range(len(matched))])
        # index_to_fill=int(tab_label.selection_get())-1
        # SIZE_drop.insert(0,matched["SIZE"][index_to_fill])
        # comp_drop.insert(0,matched["PARTY"][index_to_fill])
        # qual_drop.insert(0,matched["GRADE"][index_to_fill])
        # findrelevant(0)
#     except:
#         print("",end="")


def row_select(x):
    index_to_fill=x[1]
    PARTYin=str(comp_drop.get())
    SIZEin=str(SIZE_drop.get())
    global rejopen
    GRADEin=str(qual_drop.get())
    try:
        PARTYinH=str(comp_drop.selection_get())
    except:
        PARTYinH=''
    try:
        SIZEinH=str(SIZE_drop.selection_get())
    except:
        SIZEinH=''
    try:
        GRADEinH=str(qual_drop.selection_get())
    except:
        GRADEinH=''
    PARTYin=PARTYin.replace(PARTYinH,'')
    SIZEin=SIZEin.replace(SIZEinH,'')
    GRADEin=GRADEin.replace(GRADEinH,'')
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
    if(rejopen==1):
        matched=df[df["REJECTED"]>0]
        matched=matched.values.tolist()
        matched=pd.DataFrame(matched,columns=df.columns)
    SIZE_drop.delete(first=0,last=100)
    comp_drop.delete(first=0,last=100)
    qual_drop.delete(first=0,last=100)
    SIZE_drop.insert(0,matched["SIZE"][index_to_fill])
    comp_drop.insert(0,matched["PARTY"][index_to_fill])
    qual_drop.insert(0,matched["GRADE"][index_to_fill])
    global sheet
    findrelevant(0)
    # sheet.dehighlight_rows(index_to_fill+1)
    sheet.deselect("all")
        

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
    # df=df[df["STOCK"]>=.001 and df["REJECTED"]>=.001]
    # df.to_excel("Stock\\"+yesterday+".xlsx",index=False)
except:
    print("Yesterday's sheet not found.")
    
try:
    next_df=pd.read_excel("Stock\\"+today+".xlsx")
except:
    next_df=pd.read_excel("Stock\\"+yesterday+".xlsx")
    next_df['WEIGHT_IN']=[0]*len(df)
    next_df['WEIGHT_OUT']=[0]*len(df)
    next_df['SALE']=[0]*len(df)
    next_df['OPENING']=next_df['STOCK']
    next_df.to_excel("Stock\\"+today+".xlsx",index=False)

sorted_list_of_files=[]
list_of_files = glob.glob('Stock/*') # * means all if need specific format then *.csv
for s in list_of_files:
    s=s[6:16]
    sorted_list_of_files.append([int(s.split('-')[2]),int(s.split('-')[1]),int(s.split('-')[0])])
sorted_list_of_files.sort()
for x in range(len(sorted_list_of_files)):
    for y in range(3):
        sorted_list_of_files[x][y]=str(sorted_list_of_files[x][y])
        if(len(sorted_list_of_files[x][y])==1):
            sorted_list_of_files[x][y]='0'+sorted_list_of_files[x][y]
sorted_list_of_files=['-'.join(i[::-1]) for i in sorted_list_of_files]
sorted_list_of_files=sorted_list_of_files[sorted_list_of_files.index(yesterday)+1:]

df['GRADE']=df['GRADE'].apply(str)
df['WEIGHT_OUT']=df['WEIGHT_OUT'].apply(float)
df['WEIGHT_IN']=df['WEIGHT_IN'].apply(float)
df['WEIGHT_OUT']=df['WEIGHT_OUT'].apply(float)
df['SALE']=df['SALE'].apply(float)
df['REJECTED']=df['REJECTED'].apply(float)
df['OPENING']=df['OPENING'].apply(float)


def done():
    global root
    printsheet()
    #recalculate()
    root.destroy()


def reject():
    global rejopen
    rejected=everyone.get()
    rejsize=SIZE_drop.get()
    rejparty=comp_drop.get()
    rejgrade=qual_drop.get()
    global df
    try:
        matched=df.loc[:,'SIZE':'GRADE'].values.tolist().index([str(rejsize),str(rejparty),str(rejgrade)])
    except:
        matched=-1
    if(matched!=-1):
        if(df.at[matched,'STOCK']-float(rejected)>=0):
            df.at[matched,'STOCK']=df.at[matched,'STOCK']-float(rejected)
            df.at[matched,'REJECTED']=df.at[matched,'REJECTED']+float(rejected)
            df = df.reset_index(drop=True)
        else:
            global sheet
            sheet.set_sheet_data(data=[["INVALID","ENTRY"],["STOCK","BECOMES","NEGATIVE"],["PLEASE","CHECK","WEIGHT"]])
            rejopen=0
            return
    else:
        df=df.append({'PARTY':rejparty,'GRADE':rejgrade,'SIZE':rejsize,'WEIGHT_IN':0,'WEIGHT_OUT':0,'SALE':0,'STOCK':-1*float(rejected),'REJECTED':float(rejected)},ignore_index=True)
    df.to_excel("Stock\\"+yesterday+".xlsx",index=False)
    stocktext.set("CLOSING STOCK \n "+str(round(sum(df["STOCK"]),3)))
    for i in sorted_list_of_files:
        chain_df=pd.read_excel("Stock\\"+i+".xlsx")
        chain_df['GRADE']=chain_df['GRADE'].apply(str)
        chain_df['WEIGHT_OUT']=chain_df['WEIGHT_OUT'].apply(float)
        chain_df['WEIGHT_IN']=chain_df['WEIGHT_IN'].apply(float)
        chain_df['WEIGHT_OUT']=chain_df['WEIGHT_OUT'].apply(float)
        chain_df['SALE']=chain_df['SALE'].apply(float)
        chain_df['REJECTED']=chain_df['REJECTED'].apply(float)
        chain_df['OPENING']=chain_df['OPENING'].apply(float)
        try:
            matched=chain_df.loc[:,'SIZE':'GRADE'].values.tolist().index([str(rejsize),str(rejparty),str(rejgrade)])
        except:
            matched=-1
        if(matched!=-1):
            chain_df.at[matched,'STOCK']=chain_df.at[matched,'STOCK']-float(rejected)
            chain_df.at[matched,'OPENING']=chain_df.at[matched,'OPENING']-float(rejected)
            chain_df.at[matched,'REJECTED']=chain_df.at[matched,'REJECTED']+float(rejected)
            chain_df = chain_df.reset_index(drop=True)
        else:
            chain_df=chain_df.append({'PARTY':rejparty,'GRADE':rejgrade,'SIZE':rejsize,'WEIGHT_IN':0,'WEIGHT_OUT':0,'SALE':0,'STOCK':-1*float(rejected),'OPENING':-1*float(rejected),'REJECTED':float(rejected)},ignore_index=True)
        chain_df.to_excel("Stock\\"+i+".xlsx",index=False)
    # nextdf=df.copy()
    # nextdf['WEIGHT_IN']=[0]*len(df)
    # nextdf['WEIGHT_OUT']=[0]*len(df)
    # nextdf['SALE']=[0]*len(df)
    # nextdf=nextdf[nextdf['STOCK']>=.001]
    # nextdf.to_excel("Stock\\"+today+".xlsx",index=False)
    # try:
    #     ind=rejdf.loc[:,'R E':'C T'].values.tolist().index([str(rejsize),str(rejparty),str(rejgrade)])
    # except:
    #     ind=-1
    # if(ind==-1):
    #     rejdf=rejdf.append({"R E":rejsize,"J E":rejparty,"C T":rejgrade,"E D":float(rejected)},ignore_index=True)
    # else:
    #     rejdf.at[ind,"E D"]=float(rejdf["E D"][ind]+float(rejected))
    # rejdf=rejdf[rejdf["E D"]>=.001]
    # rejdf.to_excel("Rejects\\"+yesterday+"-REJECTS.xlsx",index=False)
    

    SIZE_drop.delete(first=0,last=100)
    comp_drop.delete(first=0,last=100)
    qual_drop.delete(first=0,last=100)
    everyone.delete(first=0,last=100)
    SIZE_drop.focus_set()

def enter():
    if(everyone_state==4):
        reject()
    else:
        global rejopen
        global SIZE
        global comp
        global qual
        global df
        global sorted_SIZE
        
        inPARTY=str(comp_drop.get()).upper()
        inSIZE=str(SIZE_drop.get())
        try:
            inSIZE=str(floatmod((inSIZE.split('x'))[0]))+'x'+str(floatmod((inSIZE.split('x'))[1]))
        except:
            return
        inGRADE=str(qual_drop.get()).upper()
        inWEIGHT_IN=everyone.get()
        inWEIGHT_OUT=everyone.get()
        inSALE=everyone.get()
        if(everyone_state==1):
            inWEIGHT_OUT=""
            inSALE=""
        if(everyone_state==2):
            inWEIGHT_IN=""
            inSALE=""
        if(everyone_state==3):
            inWEIGHT_OUT=""
            inWEIGHT_IN=""
            
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
                df=df.append({'PARTY':inPARTY,'GRADE':inGRADE,'SIZE':inSIZE,'WEIGHT_IN':floatmod(inWEIGHT_IN),'WEIGHT_OUT':floatmod(inWEIGHT_OUT),'SALE':floatmod(inSALE),'STOCK':(floatmod(inWEIGHT_IN)-floatmod(inWEIGHT_OUT)-floatmod(inSALE)),'OPENING':(floatmod(inWEIGHT_IN)-floatmod(inWEIGHT_OUT)-floatmod(inSALE)),'REJECTED':0},ignore_index=True)
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
                    global sheet
                    sheet.set_sheet_data(data=[["INVALID","ENTRY"],["STOCK","BECOMES","NEGATIVE"],["PLEASE","CHECK","WEIGHT"]])
                    rejopen=0
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
            stocktext.set("CLOSING STOCK \n "+str(round(sum(df["STOCK"]),3)))
            
            for j in range(len(sorted_list_of_files)):
                i=sorted_list_of_files[j]
                chain_df=pd.read_excel("Stock\\"+i+".xlsx")
                try:
                    cind=chain_df.loc[:,'SIZE':'GRADE'].values.tolist().index([str(inSIZE),str(inPARTY),str(inGRADE)])
                except:
                    cind=-1
                
                if(cind==-1):
                    inSIZE=str(floatmod((inSIZE.split('x'))[0]))+'x'+str(floatmod((inSIZE.split('x'))[1]))
                    chain_df=chain_df.append({'PARTY':inPARTY,'GRADE':inGRADE,'SIZE':inSIZE,'WEIGHT_IN':0,'WEIGHT_OUT':0,'SALE':0,'STOCK':(floatmod(inWEIGHT_IN)-floatmod(inWEIGHT_OUT)-floatmod(inSALE)),'OPENING':(floatmod(inWEIGHT_IN)-floatmod(inWEIGHT_OUT)-floatmod(inSALE)),'REJECTED':0},ignore_index=True)
                    chain_df.to_excel("Stock\\"+i+".xlsx",index=False)
                else:
                    s=floatmod(chain_df['STOCK'][cind])
                    chain_df.at[cind,'STOCK']=s+floatmod(inWEIGHT_IN)-floatmod(inWEIGHT_OUT)-floatmod(inSALE)
                    if(j!=0):
                        pind=pd.read_excel("Stock\\"+sorted_list_of_files[j-1]+".xlsx").loc[:,'SIZE':'GRADE'].values.tolist().index([str(inSIZE),str(inPARTY),str(inGRADE)])
                        chain_df.at[cind,'OPENING']=pd.read_excel("Stock\\"+sorted_list_of_files[j-1]+".xlsx").at[pind,'STOCK']
                    else:
                        chain_df.at[cind,'OPENING']=df.at[ind,'STOCK']
                    chain_df.to_excel("Stock\\"+i+".xlsx",index=False)
                chain_df['A']=[floatmod(i.split('x')[0]) for i in chain_df['SIZE']]
                chain_df['B']=[floatmod(i.split('x')[1]) for i in chain_df['SIZE']]
                chain_df['WEIGHT_IN']=chain_df['WEIGHT_IN'].apply(float)
                chain_df['WEIGHT_OUT']=chain_df['WEIGHT_OUT'].apply(float)
                chain_df['SALE']=chain_df['SALE'].apply(float)
                chain_df['OPENING']=chain_df['OPENING'].apply(float)
                chain_df=chain_df.sort_values(by=['A','B','PARTY'])
                chain_df=chain_df.drop('A',axis=1)
                chain_df=chain_df.drop('B',axis=1)
                chain_df = chain_df.reset_index(drop=True)
                chain_df.to_excel("Stock\\"+i+".xlsx",index=False)
            # nextdf=df.copy()
            # nextdf['WEIGHT_IN']=[0]*len(df)
            # nextdf['WEIGHT_OUT']=[0]*len(df)
            # nextdf['SALE']=[0]*len(df)
            # nextdf=nextdf[nextdf['STOCK']>=.001]
            # nextdf.to_excel("Stock\\"+today+".xlsx",index=False)
            df['WEIGHT_IN']=df['WEIGHT_IN'].apply(lambda x: round(x, 3))
            df['WEIGHT_OUT']=df['WEIGHT_OUT'].apply(lambda x: round(x, 3))
            df['SALE']=df['SALE'].apply(lambda x: round(x, 3))
            df['REJECTED']=df['REJECTED'].apply(lambda x: round(x, 3))
            df['OPENING']=df['OPENING'].apply(lambda x: round(x, 3))
            df['STOCK']=df['STOCK'].apply(lambda x: round(x, 3))
            sheet.set_sheet_data(data=df.values.tolist())
            rejopen=0
            SIZE_drop.set_completion_list(list(SIZE['SIZE']))
            qual_drop['completevalues']=list(qual['GRADE'])
            qual_drop.set_completion_list(list(qual['GRADE']))
            comp_drop['completevalues']=list(comp['PARTY'])
            comp_drop.set_completion_list(list(comp['PARTY']))
            
            wintext.set("TOTAL IN \n "+str(round(sum(df["WEIGHT_IN"]),3)))
            woutext.set("TOTAL OUT \n "+str(round(sum(df["WEIGHT_OUT"]),3)))
            saletext.set("TOTAL SALE \n "+str(round(sum(df["SALE"]),3)))
            
            everyone.delete(first=0,last=100)
            SIZE_drop.delete(first=0,last=100)
            comp_drop.delete(first=0,last=100)
            qual_drop.delete(first=0,last=100)
            SIZE_drop.focus_set()
    
def enter_and_print():
    enter()
    printsheet()

def transfer():
    # global counter
    # counter=0
    labelfont=("Fixedsys",12)
    global df
    df['STOCK']=df['STOCK'].apply(lambda x: round(x, 3))
    troot=tk.Tk()
    troot.attributes("-topmost", True)
    troot.attributes('-fullscreen', True)
    tcanvas=tk.Canvas(troot,width=1920,height=1080,bg="#FBFBFB")
    tcanvas.pack()
    
    
    
    TSIZEL1=tk.Label(troot,text="FROM SIZE",anchor="w",bg="#FBFBFB",fg="#4C3822")
    TSIZEL1.place(relx=.1,rely=.1,relheight=.05,relwidth=.2)
    TSIZEL1.config(font=labelfont)
    
    TSIZE1=tk.Entry(troot,font=fontExample,highlightthickness=3)
    TSIZE1.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
    TSIZE1.place(relx=.1,rely=.15)
    
    TCOMPL1=tk.Label(troot,text="FROM PARTY",anchor="w",bg="#FBFBFB",fg="#4C3822")
    TCOMPL1.place(relx=.1,rely=.2,relheight=.05,relwidth=.2)
    TCOMPL1.config(font=labelfont)
    
    TCOMP1=tk.Entry(troot,font=fontExample,highlightthickness=3)
    TCOMP1.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
    TCOMP1.place(relx=.1,rely=.25)
    
    TGRADL1=tk.Label(troot,text="FROM GRADE",anchor="w",bg="#FBFBFB",fg="#4C3822")
    TGRADL1.place(relx=.1,rely=.3,relheight=.05,relwidth=.2)
    TGRADL1.config(font=labelfont)
    
    TGRAD1=tk.Entry(troot,font=fontExample,highlightthickness=3)
    TGRAD1.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
    TGRAD1.place(relx=.1,rely=.35)
    
    TSIZEL2=tk.Label(troot,text="TO SIZE",anchor="w",bg="#FBFBFB",fg="#4C3822")
    TSIZEL2.place(relx=.1,rely=.4,relheight=.05,relwidth=.2)
    TSIZEL2.config(font=labelfont)
    
    TSIZE2=tk.Entry(troot,font=fontExample,highlightthickness=3)
    TSIZE2.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
    TSIZE2.place(relx=.1,rely=.45)   
    
    TCOMPL2=tk.Label(troot,text="TO PARTY",anchor="w",bg="#FBFBFB",fg="#4C3822")
    TCOMPL2.place(relx=.1,rely=.5,relheight=.05,relwidth=.2)
    TCOMPL2.config(font=labelfont)
    
    TCOMP2=tk.Entry(troot,font=fontExample,highlightthickness=3)
    TCOMP2.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
    TCOMP2.place(relx=.1,rely=.55)    
    
    TGRADL2=tk.Label(troot,text="TO GRADE",anchor="w",bg="#FBFBFB",fg="#4C3822")
    TGRADL2.place(relx=.1,rely=.6,relheight=.05,relwidth=.2)
    TGRADL2.config(font=labelfont)
    
    TGRAD2=tk.Entry(troot,font=fontExample,highlightthickness=3)
    TGRAD2.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
    TGRAD2.place(relx=.1,rely=.65)
    
    QUANL=tk.Label(troot,text="QUANTITY",anchor="w",bg="#FBFBFB",fg="#4C3822")
    QUANL.place(relx=.1,rely=.7,relheight=.05,relwidth=.2)
    QUANL.config(font=labelfont)
    
    QUAN=tk.Entry(troot,font=fontExample,highlightthickness=3)
    QUAN.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
    QUAN.place(relx=.1,rely=.75)
    
    def done2():
        troot.destroy()
    close_button2=tk.Button(troot,text="BACK",bg='blue',fg='white',command=done2,font=fontExample)
    close_button2.place(relheight=.03,relwidth=.03,relx=.97)
    
    def trow_select(y):
        # global counter
        # counter=counter+1
        
        findex=y[1]
        tsheet.highlight_rows(findex, "yellow")
        if(str(TSIZE1.get())==''):
            TSIZE1.delete(first=0,last=100)
            TCOMP1.delete(first=0,last=100)
            TGRAD1.delete(first=0,last=100)
            TSIZE1.insert(0,df['SIZE'][findex])
            TCOMP1.insert(0,df['PARTY'][findex])
            TGRAD1.insert(0,df['GRADE'][findex])
            global fromhigh
            fromhigh=findex
        else:
            TSIZE2.delete(first=0,last=100)
            TCOMP2.delete(first=0,last=100)
            TGRAD2.delete(first=0,last=100)
            TSIZE2.insert(0,df['SIZE'][findex])
            TCOMP2.insert(0,df['PARTY'][findex])
            TGRAD2.insert(0,df['GRADE'][findex])
            global tohigh
            tohigh=findex
        tsheet.deselect("all")
    tsheet = Sheet(troot,
                   page_up_down_select_row = True,
                   #empty_vertical = 0,
                   column_width = 120,
                   startup_select = (0,1,"rows"),
                   #row_height = "4",
                   #default_row_index = "numbers",
                   #default_header = "both",
                   #empty_horizontal = 0,
                   #show_vertical_grid = False,
                   #show_horizontal_grid = False,
                   #auto_resize_default_row_index = False,
                   #header_height = "3",
                   #row_index_width = 100,
                   #align = "center",
                   #header_align = "w",
                    #row_index_align = "w",
                    data =df.values.tolist(), #to set sheet data at startup
                    headers = list(df.columns),
                    #row_index = [f"Row {r}\nnewline1\nnewline2" for r in range(2000)],
                    #set_all_heights_and_widths = True, #to fit all cell sizes to text at start up
                    #headers = 0, #to set headers as first row at startup
                    #headers = [f"Column {c}\nnewline1\nnewline2" for c in range(30)],
                   #theme = "light green",
                    #row_index = 0, #to set row_index as first column at startup
                    #total_rows = 2000, #if you want to set empty sheet dimensions at startup
                    #total_columns = 30, #if you want to set empty sheet dimensions at startup
                    # height = 500, #height and width arguments are optional
                    # width = 1200 #For full startup arguments see DOCUMENTATION.md
                    )
    #self.sheet.hide("row_index")
    #self.sheet.hide("header")
    #self.sheet.hide("top_left")
    tsheet.enable_bindings(("single_select", #"single_select" or "toggle_select"
                                     "drag_select",   #enables shift click selection as well
                                     "column_drag_and_drop",
                                     "row_drag_and_drop",
                                     "column_select",
                                     "row_select",
                                     "column_width_resize",
                                     "double_click_column_resize",
                                     #"row_width_resize",
                                     #"column_height_resize",
                                     "arrowkeys",
                                     "row_height_resize",
                                     "double_click_row_resize",
                                     "right_click_popup_menu",
                                     "rc_select",
                                     "rc_insert_column",
                                     "rc_delete_column",
                                     "rc_insert_row",
                                     "rc_delete_row",
                                     "copy",
                                     "cut",
                                     "paste",
                                     "delete",
                                     "undo",
                                     "edit_cell"))
    tsheet.extra_bindings([("row_select",trow_select)])
    tsheet.change_theme("light green")
    tsheet.place(relx=.25,rely=.1,relheight=.8,relwidth=.75 )
    
    def apply_transfer():
        fromsize=TSIZE1.get()
        global df
        global rejopen
        fromparty=TCOMP1.get()
        fromgrade=TGRAD1.get()
        tosize=TSIZE2.get()
        toparty=TCOMP2.get()
        tograde=TGRAD2.get()
        quantity=round(float(QUAN.get()),3)
        fromindex=df.loc[:,'SIZE':'GRADE'].values.tolist().index([str(fromsize),str(fromparty),str(fromgrade)])
        toindex=df.loc[:,'SIZE':'GRADE'].values.tolist().index([str(tosize),str(toparty),str(tograde)])
        df['STOCK']=df['STOCK'].apply(lambda x: round(x, 3))
        if(df.at[fromindex,'STOCK']-quantity>=0):
            df.at[fromindex,'STOCK']=df.at[fromindex,'STOCK']-quantity
        else:
            wrong=tk.Tk()
            wrong.attributes("-topmost", True)
            def des():
                wrong.destroy()
            wlabel=tk.Label(wrong,text="INVALID ENTRY \n WEIGHT BECOMES NEGATIVE \n CHECK QUANTITY")
            wlabel.place(anchor='n',relx=.5,rely=0)
            wbutton=tk.Button(wrong,text="CLOSE",command=des,bg="#458BC6")
            wbutton.place(anchor='n',relx=.5,rely=.5)
            tk.mainloop()
            return
        df.at[toindex,'STOCK']=df.at[toindex,'STOCK']+quantity
        df.to_excel("Stock\\"+yesterday+".xlsx",index=False)
        tsheet.set_sheet_data(data=df.values.tolist())
        sheet.set_sheet_data(data=df.values.tolist())
        rejopen=0
        for j in range(len(sorted_list_of_files)):
                i=sorted_list_of_files[j]
                chain_df=pd.read_excel("Stock\\"+i+".xlsx")
                
                cfindex=chain_df.loc[:,'SIZE':'GRADE'].values.tolist().index([str(fromsize),str(fromparty),str(fromgrade)])
                try:
                    ctindex=chain_df.loc[:,'SIZE':'GRADE'].values.tolist().index([str(tosize),str(toparty),str(tograde)])
                except:
                    ctindex=-1
                chain_df.at[cfindex,"STOCK"]-=quantity
                chain_df.at[cfindex,"OPENING"]-=quantity
                if(ctindex==-1):
                    chain_df.append({'PARTY':toparty,'GRADE':tograde,'SIZE':tosize,'WEIGHT_IN':0,'WEIGHT_OUT':0,'SALE':0,'STOCK':quantity,'OPENING':quantity,'REJECTED':0},ignore_index=True)
                else:
                    chain_df.at[ctindex,"STOCK"]+=quantity
                    chain_df.at[ctindex,"OPENING"]+=quantity
                chain_df['A']=[floatmod(i.split('x')[0]) for i in chain_df['SIZE']]
                chain_df['B']=[floatmod(i.split('x')[1]) for i in chain_df['SIZE']]
                chain_df['WEIGHT_IN']=chain_df['WEIGHT_IN'].apply(float)
                chain_df['WEIGHT_OUT']=chain_df['WEIGHT_OUT'].apply(float)
                chain_df['SALE']=chain_df['SALE'].apply(float)
                chain_df['OPENING']=chain_df['OPENING'].apply(float)
                chain_df=chain_df.sort_values(by=['A','B','PARTY'])
                chain_df=chain_df.drop('A',axis=1)
                chain_df=chain_df.drop('B',axis=1)
                chain_df = chain_df.reset_index(drop=True)
                chain_df.to_excel("Stock\\"+i+".xlsx",index=False)                                                                                     
                TSIZE1.delete(first=0,last=100)
                TCOMP1.delete(first=0,last=100)
                TGRAD1.delete(first=0,last=100)
                TSIZE2.delete(first=0,last=100)
                TCOMP2.delete(first=0,last=100)
                TGRAD2.delete(first=0,last=100)
                QUAN.delete(first=0,last=100)
                # tsheet.dehighlight_rows('all')
    #@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    def CLEARFROM():
        # global counter
        # counter+=1
        global fromhigh
        TSIZE1.delete(first=0,last=100)
        TCOMP1.delete(first=0,last=100)
        TGRAD1.delete(first=0,last=100)
        tsheet.dehighlight_rows(fromhigh)
    def CLEARTO():
        # global counter
        # counter+=1
        global tohigh
        TSIZE2.delete(first=0,last=100)
        TCOMP2.delete(first=0,last=100)
        TGRAD2.delete(first=0,last=100)
        tsheet.dehighlight_rows(tohigh)
    enter_button3=tk.Button(troot,text="ENTER",font=("Fixedsys",13),command=apply_transfer,bg="#458BC6")
    enter_button3.place(relx=.1,rely=.85)
    clear_from=tk.Button(troot,text="CLEAR",font=("Fixedsys",13),command=CLEARFROM,bg="#458BC6")
    clear_from.place(relx=.05,rely=.25)
    clear_to=tk.Button(troot,text="CLEAR",font=("Fixedsys",13),command=CLEARTO,bg="#458BC6")
    clear_to.place(relx=.05,rely=.55)

    tk.mainloop()

# def recalculate():
#     global df
#     df=pd.read_excel("Stock\\"+yesterday+".xlsx")
#     recalculated=df.copy()
#     recalculated['WEIGHT_IN']=[0]*len(df)
#     recalculated['WEIGHT_OUT']=[0]*len(df)
#     recalculated['SALE']=[0]*len(df)
#     recalculated.to_excel("Stock\\"+today+".xlsx",index=False)

def show_rejtable():
    global df
    global sheet
    global rejopen
    rejtable=df[df["REJECTED"]>0]
    rejtable['WEIGHT_IN']=rejtable['WEIGHT_IN'].apply(lambda x: round(x, 3))
    rejtable['WEIGHT_OUT']=rejtable['WEIGHT_OUT'].apply(lambda x: round(x, 3))
    rejtable['SALE']=rejtable['SALE'].apply(lambda x: round(x, 3))
    rejtable['REJECTED']=rejtable['REJECTED'].apply(lambda x: round(x, 3))
    rejtable['OPENING']=rejtable['OPENING'].apply(lambda x: round(x, 3))
    rejtable['STOCK']=rejtable['STOCK'].apply(lambda x: round(x, 3))
    sheet.set_sheet_data(data=rejtable.values.tolist())
    rejopen=1
    
def clearall():
    print('a'+str(SIZE_drop.get())+'b')
    SIZE_drop.delete(first=0,last=100)
    comp_drop.delete(first=0,last=100)
    qual_drop.delete(first=0,last=100)
    everyone.delete(first=0,last=100)
    findrelevant(0)
    
def only_win():
    global everyone_state
    global everyone_var
    everyone_var.set("Weight In Quantity")
    everyone_state=1
def only_wout():
    global everyone_state
    global everyone_var
    everyone_var.set("Weight Out Quantity")
    everyone_state=2
def only_sale():
    global everyone_state
    global everyone_var
    everyone_var.set("Sale Quantity")
    everyone_state=3
def only_reject():
    global everyone_state
    global everyone_var
    everyone_var.set("Reject Quantity")
    everyone_state=4
def next1(x):
    comp_drop.focus_set()
def next2(x):
    qual_drop.focus_set()

def next3(x):
    everyone.focus_set()

root=tk.Tk()
root.attributes('-fullscreen', True)
root.bind("<KeyPress>",findrelevant)

canvas=tk.Canvas(root,width=WIDTH,height=HEIGHT,bg="#FBFBFB")
# gif1 = tk.PhotoImage(file="Program\logo.png")
# canvas.create_image(0,0, image=gif1,anchor='nw')
canvas.pack()

change_date=tk.Button(root, text = "CHANGE DATE",font=("Fixedsys",13),bg="#458BC6",command = restart_program)
change_date.place(relheight=.04,relx=.36)

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

everyone_var = tk.StringVar(value="Weight In Quantity")
everyone_label=tk.Label(root,textvariable=everyone_var,anchor="w",bg="#FBFBFB",fg="#4C3822")
everyone_label.place(relx=.1,rely=.4,relheight=.05,relwidth=.2)
everyone_label.config(font=labelfont)

everyone=tk.Entry(root,font=fontExample,highlightthickness=3)
everyone.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
everyone.place(relx=.1,rely=.45)

WEIGHT_IN_label=tk.Label(root,text="Enter <Weight In>",anchor="w",bg="#FBFBFB",fg="#4C3822")
# WEIGHT_IN_label.place(relx=.1,rely=.4,relheight=.05,relwidth=.2)
WEIGHT_IN_label.config(font=labelfont)

WEIGHT_IN=tk.Entry(root,font=fontExample,highlightthickness=3)
# WEIGHT_IN.bind("<Return>",next4)
WEIGHT_IN.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
# WEIGHT_IN.place(relx=.1,rely=.45)

WEIGHT_OUT_label=tk.Label(root,text="Enter <Weight Out>",anchor="w",bg="#FBFBFB",fg="#4C3822")
# WEIGHT_OUT_label.place(relx=.1,rely=.5,relheight=.05,relwidth=.2)
WEIGHT_OUT_label.config(font=labelfont)

WEIGHT_OUT=tk.Entry(root,font=fontExample,highlightthickness=3)
# WEIGHT_OUT.bind("<Return>",next5)
WEIGHT_OUT.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
# WEIGHT_OUT.place(relx=.1,rely=.55)

SALE_label=tk.Label(root,text="Enter SALE",anchor="w",bg="#FBFBFB",fg="#4C3822")
# SALE_label.place(relx=.1,rely=.6,relheight=.05,relwidth=.2)
SALE_label.config(font=labelfont)

SALE=tk.Entry(root,font=fontExample,highlightthickness=3)
# SALE.bind("<Return>",next6)
SALE.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
# SALE.place(relx=.1,rely=.65)

# -------------------------------------
# -------------------------------------
printable=df.copy()
printable['WEIGHT_IN']=printable['WEIGHT_IN'].apply(lambda x: round(x, 3))
printable['WEIGHT_OUT']=printable['WEIGHT_OUT'].apply(lambda x: round(x, 3))
printable['SALE']=printable['SALE'].apply(lambda x: round(x, 3))
printable['REJECTED']=printable['REJECTED'].apply(lambda x: round(x, 3))
printable['OPENING']=printable['OPENING'].apply(lambda x: round(x, 3))
printable['STOCK']=printable['STOCK'].apply(lambda x: round(x, 3))
sheet = Sheet(root,
                   page_up_down_select_row = True,
                   #empty_vertical = 0,
                   column_width = 120,
                   startup_select = (0,1,"rows"),
                   #row_height = "4",
                   #default_row_index = "numbers",
                   #default_header = "both",
                   #empty_horizontal = 0,
                   #show_vertical_grid = False,
                   #show_horizontal_grid = False,
                   #auto_resize_default_row_index = False,
                   #header_height = "3",
                   #row_index_width = 100,
                   #align = "center",
                   #header_align = "w",
                    #row_index_align = "w",
                    data =printable.values.tolist(), #to set sheet data at startup
                    headers = list(printable.columns),
                    #row_index = [f"Row {r}\nnewline1\nnewline2" for r in range(2000)],
                    #set_all_heights_and_widths = True, #to fit all cell sizes to text at start up
                    #headers = 0, #to set headers as first row at startup
                    #headers = [f"Column {c}\nnewline1\nnewline2" for c in range(30)],
                   #theme = "light green",
                    #row_index = 0, #to set row_index as first column at startup
                    #total_rows = 2000, #if you want to set empty sheet dimensions at startup
                    #total_columns = 30, #if you want to set empty sheet dimensions at startup
                    # height = 500, #height and width arguments are optional
                    # width = 1200 #For full startup arguments see DOCUMENTATION.md
                    )
#self.sheet.hide("row_index")
#self.sheet.hide("header")
#self.sheet.hide("top_left")
sheet.enable_bindings(("single_select", #"single_select" or "toggle_select"
                                 "drag_select",   #enables shift click selection as well
                                 "column_drag_and_drop",
                                 "row_drag_and_drop",
                                 "column_select",
                                 "row_select",
                                 "column_width_resize",
                                 "double_click_column_resize",
                                 #"row_width_resize",
                                 #"column_height_resize",
                                 "arrowkeys",
                                 "row_height_resize",
                                 "double_click_row_resize",
                                 "right_click_popup_menu",
                                 "rc_select",
                                 "rc_insert_column",
                                 "rc_delete_column",
                                 "rc_insert_row",
                                 "rc_delete_row",
                                 "copy",
                                 "cut",
                                 "paste",
                                 "delete",
                                 "undo",
                                 "edit_cell"))
sheet.extra_bindings([("row_select",row_select)])
sheet.change_theme("light green")
sheet.place(relx=.25,rely=.1,relheight=.65,relwidth=.75 )
# printable.insert(loc=0, column='##', value=[i+1 for i in range(len(printable))])
# tab_label=tk.Text(root,fg="#458BC6",bg="#FBFBFB",font=("Courier", 9,"bold"),highlightthickness=2)
# tab_label.config(highlightbackground = "black", highlightcolor="black")
# tab_label.insert(tk.END,str(tabulate(printable, headers='keys', tablefmt='fancy_grid',showindex="never")))
# tab_label.place(relx=.25,rely=.1,relheight=.65,relwidth=.75 )
# tab_label.bind("<Button-1>", selection)

reject_label=tk.Label(root,text="Enter Reject",anchor="w",bg="#FBFBFB",fg="#4C3822")
# reject_label.place(relx=.1,rely=.7,relheight=.05,relwidth=.15)
reject_label.config(font=labelfont)

reject_entry=tk.Entry(root,font=fontExample,highlightthickness=3)
reject_entry.config(highlightbackground = "#458BC6", highlightcolor="#458BC6")
# reject_entry.place(relx=.1,rely=.75)

enter_button=tk.Button(root,text="Enter",font=("Fixedsys",13),command=enter_and_print,bg="#458BC6")
enter_button.place(relx=.1,rely=.5)

# enter_button2=tk.Button(root,text="Enter",font=("Fixedsys",13),command=enter_and_print,bg="#458BC6")
# enter_button2.place(relx=.04,rely=.49)

clear_button=tk.Button(root,text="Clear",font=("Fixedsys",13),command=clearall,bg="#458BC6")
clear_button.place(relx=.04,rely=.25)

show_reject=tk.Button(root,text="Show REJECTS",font=("Fixedsys",13),command=show_rejtable,bg="#458BC6")
show_reject.place(relx=.9,rely=.055)

TRANSFER=tk.Button(root,text="TRANSFER",font=("Fixedsys",13),command=transfer,bg="#458BC6")
TRANSFER.place(relx=.8,rely=.055)

onlywin=tk.Button(root,text="WEIGHT IN",font=("Fixedsys",13),command=only_win,bg="#458BC6")
onlywin.place(relx=.25,rely=.055)

onlywout=tk.Button(root,text="WEIGHT OUT",font=("Fixedsys",13),command=only_wout,bg="#458BC6")
onlywout.place(relx=.35,rely=.055)

onlysale=tk.Button(root,text="SALE",font=("Fixedsys",13),command=only_sale,bg="#458BC6")
onlysale.place(relx=.45,rely=.055)

onlyreject=tk.Button(root,text="REJECT",font=("Fixedsys",13),command=only_reject,bg="#458BC6")
onlyreject.place(relx=.5,rely=.055)

wintext=tk.StringVar()
wintext.set("TOTAL IN \n "+str(round(sum(df["WEIGHT_IN"]),3)))
display_win=tk.Label(root,textvariable=wintext,anchor="w",bg="#FBFBFB",fg="#4C3822")
display_win.place(relx=.45,rely=.77,relheight=.05,relwidth=.15)
display_win.config(font=labelfont)

woutext=tk.StringVar()
woutext.set("TOTAL OUT \n "+str(round(sum(df["WEIGHT_OUT"]),3)))
display_wout=tk.Label(root,textvariable=woutext,anchor="w",bg="#FBFBFB",fg="#4C3822")
display_wout.place(relx=.55,rely=.77,relheight=.05,relwidth=.15)
display_wout.config(font=labelfont)

saletext=tk.StringVar()
saletext.set("TOTAL SALE \n "+str(round(sum(df["SALE"]),3)))
display_sale=tk.Label(root,textvariable=saletext,anchor="w",bg="#FBFBFB",fg="#4C3822")
display_sale.place(relx=.65,rely=.77,relheight=.05,relwidth=.15)
display_sale.config(font=labelfont)

openingtext=tk.StringVar()
openingtext.set("OPENING STOCK \n "+str(round(sum(df["OPENING"]),3)))
display_OPENING=tk.Label(root,textvariable=openingtext,anchor="w",bg="#FBFBFB",fg="#4C3822")
display_OPENING.place(relx=.3,rely=.77,relheight=.05,relwidth=.11)
display_OPENING.config(font=labelfont)

stocktext=tk.StringVar()
stocktext.set("CLOSING STOCK \n "+str(round(sum(df["STOCK"]),3)))
display_STOCK=tk.Label(root,textvariable=stocktext,anchor="w",bg="#FBFBFB",fg="#4C3822")
display_STOCK.place(relx=.75,rely=.77,relheight=.05,relwidth=.1)
display_STOCK.config(font=labelfont)

#print_button=tk.Button(root,text="Print",font=("Fixedsys",13),command=printsheet,bg="#458BC6")
#print_button.place(relx=.25,rely=.8)

#recalculate_button=tk.Button(root,text="Recalculate",font=("Fixedsys",13),command=recalculate,bg="#458BC6")
#recalculate_button.place(relx=.9,rely=.8)

#reject_button=tk.Button(root,text="Reject",font=("Fixedsys",13),command=reject,bg="#458BC6")
#   reject_button.place(relx=.04,rely=.446)

tk.mainloop()
