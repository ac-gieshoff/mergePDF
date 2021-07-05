# -*- coding: utf-8 -*-
"""
Created on Sat Jun 26 20:00:13 2021

@author: gies
"""

#load packages

from PyPDF4 import PdfFileWriter, PdfFileReader, PdfFileMerger
import glob
import os
from docx2pdf import convert
import shutil

#please adapt the path
directory = "your path"
pathlist= os.listdir(directory)
#define start folder
p=0
start = directory+ "\\" +pathlist[p]

#create temp folder
temp=directory+"\\temp"
os.makedirs(temp)

#create folder for files that cannot be processed automatically
manual=directory+"\\manual"
os.makedirs(manual)

#step 1: loop through all subdirectories and convert all to pdf.
#For some reason, convert fails to change the directory
#Let's do it 5 times after each other *sigh*

#1_original
subdir=directory+"\\"+pathlist[0]+"\\"
convert(subdir)
#2_jupe
subdir=directory+"\\"+pathlist[1]+"\\"
convert(subdir)
#3_kapm
subdir=directory+"\\"+pathlist[2]+"\\"
convert(subdir)
#4_badr
subdir=directory+"\\"+pathlist[3]+"\\"
convert(subdir)
#5_albm
subdir=directory+"\\"+pathlist[4]+"\\"
convert(subdir)
#6_albm
subdir=directory+"\\"+pathlist[5]+"\\"
convert(subdir)

    
#step two: split and save relevant page in temp
#list all files in start folder

filelist = glob.glob(start + '/*.pdf')
for file in filelist:
    #keep original name
    doc=file.split("\\")[-1]
    #original name without extension. Important for later batch upload
    original=doc.split(".")[0]
    
    p=0
     #at the beginning of each merging process, empty temp folder
    fileList=glob.glob(directory+"\\temp\\*")
    for ff in fileList:
        os.remove(ff)
    
    #get an identifier to search for files with similar filename
    doc=file.split("\\")[-1]
    identifier = doc.partition("(")[0]
    
    #now loop through each subdirectory
    for p in range(len(pathlist)):
        
        #create full path of subdirectory
        subdir=directory+ "\\" +pathlist[p]
        
        #file contains the complete path
        test = glob.glob(subdir + '/'+identifier+'*.pdf')
        if not test:
            print(f'{file} not in {subdir}. Moving to next file.')
        else:
            file = glob.glob(subdir + '/'+identifier+'*.pdf')[0]
            print(file)
            
            ##first part: define variables for later
            #document name
            doc=file.split("\\")[-1]
            #doc name without extension
            tempName=doc.split(".")[0]
            #teacher = subdirectory
            subdirId=file.split("\\")[-2]
            
            ##second part: split
            #read pdf in temp
            pdfName=glob.glob(subdir+"\\"+tempName+ ".pdf")[-1]        
            pdf=PdfFileReader(pdfName)
            numPage =pdf.getNumPages()
            #if the number of pages is not correct, save file in manual folder and move on in filelist
            if not numPage == 6:

                #simply copy file to manual folder
                outputName = manual+"\\"+f'{subdirId}{identifier}.pdf'
                shutil.copyfile(file, outputName)
                break
            else:
                    
            #split pdf, only keep first page
                pdfOut=PdfFileWriter()
                pdfOut.addPage(pdf.getPage(p))
                numPageout=pdfOut.getNumPages()
                print(f'new pdf has {numPageout} pages')
                        
                #write single page with subDirId and identifier in file name
                outputName = temp+"\\"+f'{subdirId}{identifier}.pdf'
                with open(outputName, 'wb') as out:
                    pdfOut.write(out)
          
                os.remove(pdfName)
                print("file done")
                p=p+1   
    
    ## merge pdfs in temp
    output = directory+"\\"+original +".pdf"
    pdfList = glob.glob(temp+"\\"+"*.pdf")
    if not pdfList:
        print("not in list")
        continue
    elif len(pdfList)==2:
        #copy all files in temp to manual
        for k in pdfList:
            outputName = manual+"\\"+f'{subdirId}{identifier}.pdf'
            shutil.copyfile(k, outputName)         
    else:
        print(pdfList)
        print(len(pdfList))
        
        merger = PdfFileMerger()
        for j in pdfList:      
            merger.append(open(j, 'rb'))
        with open(output, "wb") as out:
            merger.write(out)
    
    print(f'{identifier} done')  
          
print("stop")   
    
    