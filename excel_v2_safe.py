from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import tkinter.scrolledtext as tkst
import xlrd

root = Tk()

root.title("Charlie's Freelancer Helper  v0.2")

def quitdef():
        raise SystemExit

try:
        workbook = xlrd.open_workbook("hf_pp2.xls")
except:
        messagebox.showinfo("Missing File","Can't find workbook")
        quitdef()


        
        

worksheet = workbook.sheet_by_index(0)

total_rows = worksheet.nrows
total_cols = worksheet.ncols

locationValue = StringVar()
emailValue = StringVar()
phoneValue = StringVar()
dietValue = StringVar()
dayRateValue = StringVar()

noteValue = StringVar()

checkSoundVal = IntVar()
checkSoundOpVal = IntVar()
checkLXVal = IntVar()
checkLXOpVal = IntVar()
checkPowerVal = IntVar()
checkAVVal = IntVar()
checkRigVal = IntVar()
checkSetVal = IntVar()
checkCamVal = IntVar()
checkGenVal = IntVar()

finalList = []
mCellList = [17,18,21,22,19,23,24,25,26,20]

checkList = [checkLXVal.get(),checkLXOpVal.get(),checkSoundVal.get(),checkSoundOpVal.get(),checkPowerVal.get(),checkAVVal.get(),checkRigVal.get(),checkSetVal.get(),checkCamVal.get(),checkGenVal.get()]


searchName = StringVar()

#find all names:
#populate list
#find FULL names


def findnames():
        listbox.delete(0,END)
        for y in range (2, total_rows):

                #filter removes
                removeF = (worksheet.cell(y,4).value).lower()
                removeF = removeF.replace(" ","")
                if removeF != "yes" and removeF != "*":     
                        col = worksheet.cell(y,8)
                        fn = col.value
                        fn = fn.replace(" ","")

                        
                        if fn != '' and fn != 'name':
                            col2 = worksheet.cell(y,9)
                            ln = col2.value
                            ln = ln.replace(" ","")
                            value = (fn + " " + ln)
                            listbox.insert(END,value)                    

def recalldata(evt):
        checkLXVal.set(0)
        checkLXOpVal.set(0)
        checkSoundVal.set(0)
        checkSoundOpVal.set(0)
        checkPowerVal.set(0)
        checkAVVal.set(0)
        checkRigVal.set(0)
        checkSetVal.set(0)
        checkCamVal.set(0)
        checkGenVal.set(0)
        
        try:
                findname = (listbox.get(listbox.curselection()))
                #print(findname,' listbox')
                for y in range (0, total_rows):
                        fn = (worksheet.cell(y,8).value)
                        fn = fn.replace(" ","")
                        ln = (worksheet.cell(y,9).value)
                        ln = ln.replace(" ","")
                        fullname = fn + " " + ln
                        if fullname == findname:
                                locationValue.set(worksheet.cell(y,16).value)
                                emailValue.set(worksheet.cell(y,12).value)
                                phoneValue.set(worksheet.cell(y,13).value)
                                dayRateValue.set(worksheet.cell(y,14).value)
                                searchName.set(fullname)
                                
                                if (worksheet.cell(y,5).value) == "":
                                        dietValue.set("No Info")
                                else:
                                        dietValue.set(worksheet.cell(y,5).value)
                                notes.delete('1.0',END)
                                if worksheet.cell(y,33).value == "":
                                        notes.insert(END,"no notes on this freelancer...")
                                else:
                                        notes.insert(END,worksheet.cell(y,33).value)
                                for m in range(len(mCellList)):
                                        if worksheet.cell(y,mCellList[m]).value != "":
                                                if mCellList[m] == 17:
                                                        checkLXVal.set(1)
                                                if mCellList[m] == 18:
                                                        checkLXOpVal.set(1)
                                                if mCellList[m] == 21:
                                                        checkSoundVal.set(1)
                                                if mCellList[m] == 22:
                                                        checkSoundOpVal.set(1)
                                                if mCellList[m] == 19:
                                                        checkPowerVal.set(1)
                                                if mCellList[m] == 23:
                                                        checkAVVal.set(1)
                                                if mCellList[m] == 24:
                                                        checkRigVal.set(1)
                                                if mCellList[m] == 25:
                                                        checkSetVal.set(1)
                                                if mCellList[m] == 26:
                                                        checkCamVal.set(1)
                                                if mCellList[m] == 20:
                                                        checkGenVal.set(1)

        except:
                pass
                #print('unselected the listbox')

def searchForName():
        found = False
        listbox.selection_clear(0,END)
        searched = searchName.get()

        findnames()
        
        #check listbox for name
        for ind in range(0,listbox.size()):
                if searched.lower() == listbox.get(ind).lower():
                       # print('found') 
                        listbox.selection_set(ind)

                        #repopulate data fields
                        recalldata(0)
                        listbox.see(ind)
                        found = True
                        break
                
        if found == False:
                messagebox.showerror("Error", "Can't Find Name")
                               

def find():
        #reset data
        listbox.delete(0,END)

        list1 =[]
        list2 =[]
        list3 =[]
        list4 =[]
        list5 =[]
        list6 =[]
        list7 =[]
        list8 =[]
        list9 =[]
        list10=[]
        
        list1.clear()
        list2.clear()
        list3.clear()
        list4.clear()
        list5.clear()
        list6.clear()
        list7.clear()
        list8.clear()
        list9.clear()
        list10.clear()


        m1 = []
        m2 = []
        m3 = []
        m4 = []
        m5 = []
        m6 = []
        m7 = []
        m8 = []
        m9 = []
        m10 = []
        mainList = [m1,m2,m3,m4,m5,m6,m7,m8,m9,m10]

        disciplineLists=[]
        disciplineLists.clear()
        disciplineLists = [list1,list2,list3,list4,list5,list6,list7,list8,list9,list10]
        
        mCellList = []
        mCellList.clear()
        mCellList = [17,18,21,22,19,23,24,25,26,20]

        checkList = []
        checkList.clear()
        
        
        checkList = [checkLXVal.get(),checkLXOpVal.get(),checkSoundVal.get(),checkSoundOpVal.get(),checkPowerVal.get(),checkAVVal.get(),checkRigVal.get(),checkSetVal.get(),checkCamVal.get(),checkGenVal.get()]
        #print(checkList[0] + checkList[1] + checkList[2] + checkList[3])

        mCell = 0
        
        #check if only one checkbox presssed
        if (checkList[0] + checkList[1] + checkList[2] + checkList[3] + checkList[4] + checkList[5] + checkList[6] + checkList[7] + checkList[8] + checkList[9] < 2):

                #check which checkbox pressed
                for check in range (len(checkList)):
                        #print("check number",check)
                        #print("output",checkList[check])
                        if checkList[check] == 1:
                                #lx
                                if check == 0:
                                        mCell = 17
                                
                                        break
                                #lxOP
                                elif check == 1:
                                        mCell = 18
                                
                                        break
                                #sound
                                elif check == 2:
                                        mCell = 21
                                
                                        break
                                #soundOP
                                elif check ==  3:
                                        mCell = 22
                                
                                        break
                                #power
                                elif check ==  4:
                                        mCell = 19
                                
                                        break
                                #av
                                elif check ==  5:
                                        mCell = 23
                                
                                        break
                                #rigger
                                elif check ==  6:
                                        mCell = 24
                                
                                        break
                                
                                #set
                                elif check ==  7:
                                        mCell = 25
                                
                                        break
                                #camera
                                elif check ==  8:
                                        mCell = 26
                                
                                        break
                                
                                #general
                                elif check ==  8:
                                        mCell = 20
                                
                                        break
                                
                        
       #find first discipline
                #print('mcell',mCell)
                for y in range (3, total_rows):
                        col = (worksheet.cell(y,mCell).value)
                #if cell is empty then person available
                        if col != "":
                                addfullname(y)
                        
        #more than one option picked                        
        else:
                count = -1
                for listNumber in range (0,(len(disciplineLists))):
                        count += 1
                        for y in range (3, total_rows):
                                col = (worksheet.cell(y,mCellList[count]).value)
                #if cell is empty then person available
                                if col != "":
                                #make fullname
                                        fn = (worksheet.cell(y,8).value)
                                        fn = fn.replace(" ","")
                                        ln = (worksheet.cell(y,9).value)
                                        ln = ln.replace(" ","")
                                        fullname = fn + " " + ln
                                
                                #add to lists
                                        disciplineLists[count].append(fullname)

        #lists made! now to cross reference them.........
                                                      
                
                #find out what checklists are checked and add checked to new lists
                nCount = 0
                for x in range(len(checkList)):
                        if checkList[x] == 1:
                                mainList[nCount] = disciplineLists[x]
                                nCount +=1

                #find first list 
                for xx in range (len(mainList)):
                        if mainList[xx] != []:
                                firstList = mainList[xx]
                                del mainList[xx]
                                #print(firstList)
                                break
                #find second list
                for xy in range (len(mainList)):
                        if mainList[xy] != []:
                                secondList = mainList[xy]
                                del mainList[xy]
                                #print(secondList)
                                break
                        
                #find third list 
                for xxx in range (len(mainList)):
                        if mainList[xxx] != []:
                                thirdList = mainList[xxx]
                                del mainList[xxx]
                                #print(thirdList)
                                break
                        
                #find fourth list
                for xyy in range (len(mainList)):
                        if mainList[xyy] != []:
                                fourthList = mainList[xyy]
                                del mainList[xyy]
                                #print(fourthList)
                                break

                #find fifth list
                for xyx in range (len(mainList)):
                        if mainList[xyx] != []:
                                fifthList = mainList[xyx]
                                del mainList[xyx]
                                #print(fifthList)
                                break
                        
                for yy in range (len(mainList)):
                        if mainList[yy] != []:
                                sixthList = mainList[yy]
                                del mainList[yy]
                                #print(sixthList)
                                break

                for yyy in range (len(mainList)):
                        if mainList[yyy] != []:
                                seventhList = mainList[yyy]
                                del mainList[yyy]
                                #print(seventhList)
                                break

                for yyyx in range (len(mainList)):
                        if mainList[yyyx] != []:
                                eighthList = mainList[yyyx]
                                del mainList[yyyx]
                                #print(eighthList)
                                break

                for yyxx in range (len(mainList)):
                        if mainList[yyxx] != []:
                                ninthList = mainList[yyxx]
                                del mainList[yyxx]
                                #print(ninthList)
                                break
                        
                for xxxx in range (len(mainList)):
                        if mainList[xxxx] != []:
                                tenthList = mainList[xxxx]
                                del mainList[xxxx]
                                #print(tenthList)
                                break
                
                #match lists
                if nCount == 2:
                        for name in range(1000):
                                try:
                                        if firstList[name] in secondList:
                                                listbox.insert(END,firstList[name])
                                except:
                                        pass
                
                if nCount == 3:
                        for name in range(1000):
                                try:
                                        if (firstList[name] in secondList) and (firstList[name] in thirdList):
                                                listbox.insert(END,firstList[name])
                                except:
                                        pass

                if nCount == 4:
                        for name in range(1000):
                                try:
                                        if (firstList[name] in secondList) and (firstList[name] in thirdList) and (firstList[name] in fourthList):
                                                listbox.insert(END,firstList[name])
                                except:
                                        pass

                if nCount == 5:
                        for name in range(1000):
                                try:
                                        if (firstList[name] in secondList) and (firstList[name] in thirdList) and (firstList[name] in fourthList) and (firstList[name] in fifthList):
                                                listbox.insert(END,firstList[name])
                                except:
                                        pass

                if nCount == 6:
                        for name in range(1000):
                                try:
                                        if (firstList[name] in secondList) and (firstList[name] in thirdList) and (firstList[name] in fourthList) and (firstList[name] in fifthList) and (firstList[name] in sixthList):
                                                listbox.insert(END,firstList[name])
                                except:
                                        pass

                if nCount == 7:
                        for name in range(1000):
                                try:
                                        if (firstList[name] in secondList) and (firstList[name] in thirdList) and (firstList[name] in fourthList) and (firstList[name] in fifthList) and (firstList[name] in sixthList) and (firstList[name] in seventhList):
                                                listbox.insert(END,firstList[name])
                                except:
                                        pass

                if nCount == 8:
                        for name in range(1000):
                                try:
                                        if (firstList[name] in secondList) and (firstList[name] in thirdList) and (firstList[name] in fourthList) and (firstList[name] in fifthList) and (firstList[name] in sixthList) and (firstList[name] in seventhList) and (firstList[name] in eighthList):
                                                listbox.insert(END,firstList[name])
                                except:
                                        pass

                if nCount == 9:
                        for name in range(1000):
                                try:
                                        if (firstList[name] in secondList) and (firstList[name] in thirdList) and (firstList[name] in fourthList) and (firstList[name] in fifthList) and (firstList[name] in sixthList) and (firstList[name] in seventhList) and (firstList[name] in eighthList) and (firstList[name] in ninthList):
                                                listbox.insert(END,firstList[name])
                                except:
                                        pass
                                
                if nCount == 10:
                        for name in range(1000):
                                try:
                                        if (firstList[name] in secondList) and (firstList[name] in thirdList) and (firstList[name] in fourthList) and (firstList[name] in fifthList) and (firstList[name] in sixthList) and (firstList[name] in seventhList) and (firstList[name] in eighthList) and (firstList[name] in ninthList) and (firstList[name] in tenthList):
                                                listbox.insert(END,firstList[name])
                                except:
                                        pass
                
                
                
                
        
        
                        
                        
        
                
                
                        
                        #found match
                        
                        #add to final list
                
                        
                                
                
        
                
                        
                
def addfullname(y):
        fn = (worksheet.cell(y,8).value)
        fn = fn.replace(" ","")
        ln = (worksheet.cell(y,9).value)
        ln = ln.replace(" ","")
        fullname = fn + " " + ln
        listbox.insert(END,(fullname))
        

#GUI

searchButton = ttk.Button(root, width = 20, text = "   Search   ", command=searchForName).grid(column = 1, row = 2,sticky = E)
searchNameEntry = ttk.Entry(root,width = 40,textvariable = searchName).grid(column = 1, row = 2,sticky = W)
searchLabel = ttk.Label(root,text = " Full Name").grid(column = 0,row = 2,sticky = W)

locationLabel = ttk.Label(root, text = " Location :").grid(column = 0, row = 3, sticky = W)
locationEntry = ttk.Entry(root,width = 40, textvariable = locationValue).grid(column = 1,row = 3,sticky = W)

emailLabel = ttk.Label(root, text = "      Email :").grid(column= 0, row = 4,sticky = W)
emailEntry = ttk.Entry(root,width = 40,textvariable = emailValue).grid(column = 1, row = 4,sticky = W)

phoneLabel = ttk.Label(root, text = "    Phone :").grid(column= 0, row = 5,sticky = W)
phoneEntry = ttk.Entry(root,width = 40, textvariable = phoneValue).grid(column = 1, row = 5,sticky = W)

dayRateLabel = ttk.Label(root, text = "Day Rate :").grid(column= 0, row = 6,sticky = W)
dayRateEntry = ttk.Entry(root,width = 40, textvariable = dayRateValue).grid(column = 1, row = 6,sticky = W)

dietLabel = ttk.Label(root, text = "        Diet :").grid(column= 0, row = 7,sticky = W)
dietEntry = ttk.Entry(root ,width = 40, textvariable = dietValue).grid(column = 1, row = 7,sticky = W)

spacerLabel = ttk.Label(root, text = "").grid(column =0, row = 2, sticky = W,pady= 20)
spacerLabel2 = ttk.Label(root, text = "").grid(column =0, row = 8, sticky = W,pady= 10)

soundCheck = Checkbutton(root, text = "Sound", variable = checkSoundVal).grid(column = 0, row = 10,sticky = W)
soundOpCheck = Checkbutton(root, text = "Sound Op", variable = checkSoundOpVal).grid(column = 0,row =11,sticky = W)
genCheck = Checkbutton(root, text = "General Tech",variable = checkGenVal).grid(column = 0, row=12, sticky = W)
lxCheck = Checkbutton(root, text = "LX", variable = checkLXVal).grid(column = 1,row =10,sticky = W)
lxOpCheck = Checkbutton(root, text = "LX Op", variable = checkLXOpVal).grid(column = 1, row = 11,sticky = W)
powerCheck = Checkbutton(root, text = "Power", variable = checkPowerVal).grid(column = 1, row = 12,sticky = W)
AVCheck = Checkbutton(root, text = "AV      ", variable = checkAVVal).grid(column = 1, row = 10)
rigCheck = Checkbutton(root, text = "Rigger", variable = checkRigVal).grid(column = 1, row = 11)
setCheck = Checkbutton(root, text = "Set      ", variable = checkSetVal).grid(column = 1, row = 12)
camCheck = Checkbutton(root, text = "Camera Op",variable = checkCamVal).grid(column = 1,row=10, sticky = E)


updateButton = ttk.Button(root,width = 20,text = "Update List",command = find).grid(column = 1, row = 13,sticky = E)
disciplineLabel = ttk.Label(root, text = "Discipline :").grid(column = 0, row = 9,sticky = W)

notes = tkst.ScrolledText(root,height = 7, width = 25,borderwidth=3)
notes.grid(column = 1,row = 14)
notes.config(wrap= 'word',font = 'helvetica 20')


listbox = Listbox(root,selectmode=SINGLE,height = 14)
listbox.grid(column = 0,row = 14,pady = 50)

listbox.bind('<<ListboxSelect>>',recalldata)



findnames()

   
root.mainloop()
