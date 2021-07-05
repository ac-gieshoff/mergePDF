# -*- coding: utf-8 -*-
"""
Created on Sat Jun 26 20:00:13 2021

@author: gies
"""
#This script converts docx-files to pdf, splits the new PDF in single pages, keeps one page and merges all pages containing the same identifier in the filename into a new PDF.
#In my example, the docx-files were saved in subdirectories and each subdiretory contained one file with the same identifier. The subdirectories were numbered to ensure the correct order of pages in the merged PDF.
#The identifier is a regular expression.
# The merged PDF should contain a certain number of pages. Any docx-file that contained more or less pages was moved to a "manual" folder for manual processing.

#load packages

from PyPDF4 import PdfFileWriter, PdfFileReader, PdfFileMerger
import glob
import os
from docx2pdf import convert
import shutil

#define the number of pages that allow save processing
requiredPages = 'some number'

#specifiy the path to the folder that contains your docx-files.
directory = "your path"
pathlist= os.listdir(directory)
#specify the folder that contains the first page of your merged PDF-file
p=0
start = directory+ "\\" +pathlist[p]

#create temp folder for temporary data.
temp=directory+"\\temp"
os.makedirs(temp)

#create folder for files that cannot be processed automatically
manual=directory+"\\manual"
os.makedirs(manual)

#The following lines convert all docx in each subdiretory to PDF.
#For some reason, convert fails to change the directory
#I recommend to do run the follwong lines up to line 55 seperatly.

#1_folder
subdir=directory+"\\"+pathlist[0]+"\\"
convert(subdir)
#2_folder
subdir=directory+"\\"+pathlist[1]+"\\"
convert(subdir)
#3_folder
subdir=directory+"\\"+pathlist[2]+"\\"
convert(subdir)
#4_folder
subdir=directory+"\\"+pathlist[3]+"\\"
convert(subdir)
#5_folder
subdir=directory+"\\"+pathlist[4]+"\\"
convert(subdir)
#6_folder
subdir=directory+"\\"+pathlist[5]+"\\"
convert(subdir)

    
#step two: split and save relevant page in temp
#list all files in start folder

filelist = glob.glob(start + '/*.pdf')
for file in filelist:
    #keep original name
    doc=file.split("\\")[-1]
    #keep the original name without extension
    original=doc.split(".")[0]
    
    p=0
     #at the beginning of each merging process, empty temp folder.
        #list all files in the directory
    fileList=glob.glob(directory+"\\temp\\*")
    #remove the files in that directory one by one
    for ff in fileList:
        os.remove(ff)
    
    #get an identifier to search for files with similar filename
    doc=file.split("\\")[-1]
    #keep all characters up to the character that defines the end of your identifying string
    identifier = doc.partition("pattern that defines your identifier")[0]
    
    #now loop through each subdirectory
    for p in range(len(pathlist)):
        
        #create full path of subdirectory
        subdir=directory+ "\\" +pathlist[p]
        
        #file contains the complete path
        test = glob.glob(subdir + '/'+identifier+'*.pdf')
        
        # if no file with the same identifier can be found in the subdirectory, move to next file in filelist.
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
            #subdirectory
            subdirId=file.split("\\")[-2]
            
            ##second part: split
            #read pdf in temp
            pdfName=glob.glob(subdir+"\\"+tempName+ ".pdf")[-1]        
            pdf=PdfFileReader(pdfName)
            numPage =pdf.getNumPages()
            #if the number of pages is not correct, save file in manual folder and move on in filelist
            if not numPage == requiredPages:

                #simply copy file to manual folder
                outputName = manual+"\\"+f'{subdirId}{identifier}.pdf'
                #shutil copies the file, it takes the filepath and the new filepath as arguments
                shutil.copyfile(file, outputName)
                break
            else:
                    
            #split pdf, only keep first page
            #open the connection to write a new PDF
                pdfOut=PdfFileWriter()
                #define the page you would like to keep
                pdfOut.addPage(pdf.getPage(p))
                #create a new PDF with only that page 
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
    
    #if the temp folder is empty, move to the next file in filelist
    if not pdfList:
        print("not in list")
        continue
        
    #if there is more than 1 file, copy all files to the manual folder
    elif len(pdfList)>1:
        #copy all files in temp to manual
        for k in pdfList:
            outputName = manual+"\\"+f'{subdirId}{identifier}.pdf'
     
    shutil.copyfile(k, outputName)    
    
    #otherwise merge all PDF in the temp folder to one PDF and save the new PDF on the top level of the directory
    else:
        print(pdfList)
        print(len(pdfList))
        #open the file merger
        merger = PdfFileMerger()
        #append all files in the temp folder
        for j in pdfList:      
            merger.append(open(j, 'rb'))
            #write a new PDF 
        with open(output, "wb") as out:
            merger.write(out)
    
    print(f'{identifier} done')  
          
print("stop")   
    
    
