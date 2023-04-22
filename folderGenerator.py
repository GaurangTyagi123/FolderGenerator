from customtkinter import *
from tkinter import filedialog,messagebox
import pandas as pd
import re
import os

win = CTk()
set_appearance_mode("system")
set_default_color_theme("blue")
win.title("Excel Folder maker")
win.geometry("410x300")

font1=CTkFont(family='monospace',size=15,weight='bold'  ,slant='roman',underline=0,overstrike=0)
font2=CTkFont(family='tahoma'   ,size=15,weight='normal',slant='roman',underline=0,overstrike=0)
font4=CTkFont(family='monospace',size=15,weight='normal',slant='roman',underline=0,overstrike=0)

supported_file_types = (
    ("xlsx files",".xlsx"),
    ("all files","*.*")
)

filepath = ""
folderpath = ""
CFT = False
CBF = False

def findSheetNumberAndRange(columnRange):
    values = dict()
    values["sheet_name"] = False

    temp = columnRange
    search = re.search("!",temp,re.IGNORECASE)
    if(search):
        sheet_name= ""
        columnRange = columnRange.split(":")
        sheet_name = columnRange[0][:columnRange[0].find("!")]
        columnRange[0] = columnRange[0][columnRange[0].find("!")+1:]
        columnRange[1] = columnRange[1][columnRange[1].find("!")+1:]
        # processing for column name
        colTemp = re.findall("[A-Za-z]+",columnRange[0])[0]
        columnName =  ord(colTemp.upper())-65;        
        
        From = re.findall("\d+",columnRange[0])[0]
        to = re.findall("\d+",columnRange[1])[0]
        values["sheet_name"] = sheet_name
    else:
        colTemp = temp.split(":")[0]
        colTemp = re.findall("[A-Za-z]+",colTemp)[0]
        columnName =  ord(colTemp.upper())-65;        
        From = re.findall("\d+",temp.split(":")[0])[0]
        to = re.findall("\d+",temp.split(":")[1])[0]
    

    values["From"] = int(From)-2
    values["to"] = int(to)-2
    values["column"]  = int(columnName)
    
    return values

def folderGenerator(filename="",columnRange="",extraText=""):
    try:
        if(len(columnRange)!=0):
            values = findSheetNumberAndRange(columnRange)
        if(values['sheet_name']!=False):
            sheetName = values['sheet_name']
            data=pd.read_excel(filename,sheet_name=sheetName)
        else:
            data = pd.read_excel(filename)
        preview = open("preview.txt","w")
    except:
        return f"Error.. Could not open {filename}"
    
    columnName = data.columns[values["column"]]
    if(len(columnRange)!=0):
        if(values["to"]>=0 and values["From"]>=0):
            for folderName in data.loc[values["From"]:values["to"],columnName]:
                if(str(folderName).lower()!="nan"):
                    preview.write(folderName+extraText+"\n")
            return "Selected cells have been read"
    else:
        return "Error... column range is incorrect"
    

def fileOpener():
    global filepath,lab_file ,CTF,CBF
    filepath = filedialog.askopenfilename(title="Open Excel File",filetypes= supported_file_types)
    lab_file.configure(state = NORMAL)
    lab_file.delete(0,END)
    lab_file.insert(0,filepath.split("/")[-1])
    if not(filepath.endswith(".xlsx")):
        lab_file.configure(border_color="#c03737")
        CTF = False
    else:
        lab_file.configure(border_color="#2fa572")
        CTF = True
    lab_file.configure(state = DISABLED)

def folderOpener():
    global folderpath,CBF,CFT
    folderpath = filedialog.askdirectory() 
    if folderpath != "":
        CBF = True
        lab_base_dir.configure(state=NORMAL)
        lab_base_dir.delete(0,END)
        lab_base_dir.insert(0,"/".join(folderpath.split("/")[-2:]))
        lab_base_dir.configure(border_color="#2fa572")
        lab_base_dir.configure(state=DISABLED)
    else:
        CBF = True
        lab_base_dir.configure(state=NORMAL)
        lab_base_dir.delete(0,END)
        lab_base_dir.configure(border_color="#c03737")
        lab_base_dir.configure(state=DISABLED)



def getprev():
    fromcell = cell_from.get().strip()
    tocell = cell_to.get().strip()
    if (filepath !="") and (folderpath !="") and (filepath.endswith(".xlsx")) and (fromcell != "") and (tocell != ""):
        print("success")
        status = folderGenerator(filepath,fromcell+":"+tocell)
        messagebox.showinfo("status",status)
        os.startfile("preview.txt")
        messagebox.showinfo("Content Editor","Makes edits in the textfile to be reflected in the folders.")
        Execute.configure(state=NORMAL)

    else:
        messagebox.showerror("Input Error","Check for the inputs , there might be an error.")
        print("failiure")

def execute():
    
    f = open("preview.txt","r")
    exists = []
    for i in f:
        try:
            os.mkdir(folderpath+"/"+i.strip("\n"))
        except FileExistsError:
            exists.append(folderpath+"/"+i.strip("\n"))
            continue
    f.close()
    if(len(exists)!=0):messagebox.showinfo("status",f"these folders {exists} already exist....remaining folders have been created")

    fromcell = cell_from.get().strip()
    tocell = cell_to.get().strip()
    values = findSheetNumberAndRange(fromcell+":"+tocell)
    # opening
    if(values["sheet_name"]):
        data = pd.read_excel(filepath,sheet_name=values["sheet_name"])
        with pd.ExcelWriter(filepath,mode="a",if_sheet_exists="replace") as writer:
            columnName = data.columns[values['column']]
            data.loc[values['From']:values['to'],columnName] = data.loc[values['From']:values['to'],columnName].apply(lambda x:f'=HYPERLINK("{folderpath+"/"+str(x)}","{x}")')
            data.to_excel(writer,sheet_name=values["sheet_name"],index=False)
        
    else: 
        with pd.ExcelWriter(filepath,mode="a",if_sheet_exists="replace") as writer:
            data = pd.read_excel(filepath)
            columnName = data.columns[values['column']]
            data.loc[values['From']:values['to'],columnName] = data.loc[values['From']:values['to'],columnName].apply(lambda x:f'=HYPERLINK("{folderpath+"/"+str(x)}","{x}")')
            data.to_excel(writer,index=False)
    
    
    # storing
    os.remove("preview.txt")


lab_file = CTkEntry(win,height=25,width=200,font=font4,placeholder_text="Choose File",border_width=2,border_color="#C03737",state=DISABLED)
lab_file.place(x = 20,y = 20)

But = CTkButton(win,text="Open File",font=font1,width=165,command=fileOpener).place(x= 230,y = 20) 

lab_base_dir = CTkEntry(win,height=25,width=200,font=font4,placeholder_text="Choose File",border_width=2,border_color="#C03737",state=DISABLED)
lab_base_dir.place(x = 20,y = 60)

But_base_dir = CTkButton(win,text="Open Base Directory",font=font1,width=165,command=folderOpener).place(x = 230,y = 60)

lab_from = CTkLabel(win,text="from",font=font1).place(x = 20,y = 100)
lab_To = CTkLabel(win,text="To",font=font1).place(x = 20,y = 140)

cell_from = CTkEntry(win,height=25,width=75,font=font4,placeholder_text="from cell",border_width=2,border_color="#dddddd")
cell_from.place(x = 100,y = 100)
cell_to = CTkEntry(win,height=25,width=75,font=font4,placeholder_text="to cell",border_width=2,border_color="#dddddd")
cell_to.place(x = 100,y = 140)

get_preview = CTkButton(win,text="Get Preview",font=font1,width=370,command=getprev).place(x = 20,y = 180)

Execute = CTkButton(win,text="Execute",font=font1,width=370,fg_color="#c03737",state=DISABLED,command=execute)
Execute.place(x = 20,y = 220)


win.mainloop()