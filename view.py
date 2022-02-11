#view

def Start():
    #this function is run at the start of application to show first screen with file selection.
    print("Select files:")
    f1 = input("MBank 1 csv file: ")
    f2 = input("MBank 2 csv file: ")
    f3 = input("PKO csv file: ")
    f4 = input("Output file name (with .xlsx): ")
    return f1,f2,f3,f4
  
    
def JoinFilesOK():
    #this function run popup with info that files were joined into xlsx file but not checked.
    print("Files were joined correctly. Please be aware that files were not checked for their properness")
    #popup with ok
    pass

def Error(e):
    #this function run popup with error info
    print("Oops!", e.__class__,e, "occurred.")

def CategorySaved(i):
    #this function run popup with info that cateogry was saved for the transaction
    print("Category saved ",i)   
   
def CategorySavedDict(i):
    #this function run popup with info that cateogry was saved for the transaction and also added to dictionary
    print("Category saved and added to Dictionary",i)  

def EnterCategories(a,b,c,d,e):
    #this function run popup with possibility to enter category for transaction and to add this cateogry to dictionary
    u = ""
    while u !="ok":
        print("Date: "+ a )
        print("Amount: "+ b)
        print("Description: "+c+" "+d+" "+e )   
        cat = input("Category: ")
        m = CheckCategoryName(cat)
        if m ==False:
            print("Category not in dictionary")
            continue
        elif m==True:
            sav = input("Save category to dictionary? (y/n)")
            if sav == "y":
                sav = True
            else:
                sav = False
            return d,cat,sav                 


def SummaryPrepared(file):
    #this function run popup with info that summaries for sheets and file are finished
    print ("Calculation and summaries completed in "+ file)

def CheckCategoryName(category):
    #this function checks if category is in the list of categories. In the future list should be extenralized and made a config file.
    catdict = ReadCategoryFile()
    if category in catdict:
        return True
        #print("Category OK")
    else:
        return False
        #print("Wrong Category")  
import os

def ReadCategoryFile():
    #this function returns list with category names
    autofolder = str(os.path.dirname(__file__))
    file_to_open = autofolder + "/categories.csv"
    with open(file_to_open) as file:
        catdict = file.readlines()
        catdict = [line.rstrip() for line in catdict]
    return catdict

def TestReadCategoryFile():
    #this function if Readcategory was working fine with previouse set of records.
    catdict = ReadCategoryFile()
    #print(catdict)
    catdict1 = 0 #copy here content of your dict.csv as list
    if catdict == catdict1:
        print("dobre")
    else:  
        print("zle")

#Start()
#TestReadCategoryFile();