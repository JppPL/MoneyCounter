#controller


import model as m
import view as v
     

def JoinFiles(f1,f2,f3,f4):
    #function is joining files into one xlsx file to further process them.
    r = m.JoinFiles(f1,f2,f3,f4)
    if r == "ok":
        v.JoinFilesOK()
    else:
        v.Error("Files not joined")   

def AssignCategoriesMB(file,sheet):
    #function runs trough each mbank sheet in xslx file (first two) and assingn categories to each transaction. 
    count = m.CountCategoriesMB(file,sheet)
    for i in range (count[0]+1,count[1]-4):
        r = m.AssignCategoriesMB(i, file,sheet)
        if r =="assigned":
            pass
        else:
            c = v.EnterCategories(r[0],r[1],r[2],r[3],r[4])
            k =""
            bank = "mb"
            while k != "ok" and k!="okDict":
                k = m.SaveCategory(i,c[0],c[1],c[2],file,sheet,bank)
                continue
            else:
                if k=="ok":
                    v.CategorySaved(i)
                else:
                    v.CategorySavedDict(i)

def AssignCategoriesBP(file,sheet):
    #function runs trough pkobp sheet in xslx file (third one) and assign categories to each transaction. 
    count = m.CountCategoriesBP(file,sheet)
    for i in range (count[0]+1,count[1]):
        r = m.AssignCategoriesBP(i, file,sheet)
        #print (r)
        if r == "assigned":
            pass
        else:
            c = v.EnterCategories(r[0],r[1],r[2],r[3],r[4])
            k =""
            bank = "bp"
            while k != "ok" and k !="okDict":
                k = m.SaveCategory(i,c[0],c[1],c[2],file,sheet,bank)
                continue
            else:
                if k=="ok":
                    v.CategorySaved(i)
                else:
                    v.CategorySavedDict(i)

def testAssignCategoriesBP():
    #this function tests assign cateogries bp
    AssignCategoriesBP("enter test file",'entertestsheet')

def AddSummaries(file):
    #this function generates set of summaries in xlsx file - summaries in each sheet and then summary sheet with all category data in one sheet.
    m.SheetSummaryMbank(file,'file1')
    m.SheetSummaryMbank(file,'file2')
    m.SheetSummaryPkoBP(file,'file3')
    m.FileSummary(file,'file1',1)
    m.FileSummary(file,'file2',2)
    m.FileSummary(file,'file3',3)
    m.SummarySummary(file)
    v.SummaryPrepared(file)

#here we run all functions in correct order, and handle errors.
def RunAll():
    try:
        f1,f2,f3,f4 = v.Start()
        JoinFiles(f1,f2,f3,f4)
        #f4="t2.xlsx"
        AssignCategoriesMB(f4,"file1")
        AssignCategoriesMB(f4,"file2")
        AssignCategoriesBP(f4,"file3")
        AddSummaries(f4)
    except Exception as e:
            v.Error(e)

def FixBP():
    try:
        print("Fixer PKO + Summary")
        f4 = input ("xlsx file to fix: ")
        AssignCategoriesBP(f4,"file3")
        AddSummaries(f4)
    except Exception as e:
            v.Error(e)

def FixMB2andBP():
    try:
        print("fixer only mbank2 + pko+ summary")
        f4 = input ("xlsx to fix : ")
        AssignCategoriesMB(f4,"file2")
        AssignCategoriesBP(f4,"file3")
        AddSummaries(f4)
    except Exception as e:
            v.Error(e)

def OnlySummaries(f4):
    try:
        #AssignCategoriesMB(f4,"file1")
        #ategoriesMB(f4,"file2")
        #AssignCategoriesBP(f4,"file3")
        AddSummaries(f4)
    except Exception as e:
            v.Error(e)

if __name__ == "__main__":
    # this won't be run when imported
    #OnlySummaries("sie2020.xlsx")
    RunAll()
    #FixBP()
    #FixMB2andBP()
    pass
