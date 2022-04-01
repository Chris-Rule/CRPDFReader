import os
from io import StringIO

from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser

#Get current directory
WorkingDirectory = "C:\Python Sandbox\PyReadExcel\Larger Test"
#print("Current directory = ",WorkingDirectory)

#list files in current directory
OriginalList = os.listdir(WorkingDirectory)
#print("Original list =",OriginalList)

#clean up directory / target PDFs only
NewList = []
for i in OriginalList:
    if "~" in i or "txt" in i or "csv" in i:
        pass
    else:
        NewList.append(i)


#create blank output list
output = []

#FUNCTION to loop through text and find results
def find_text_function(TargetFilePath):
    #Get text from pdf document
    my_text_string = StringIO()
    with open(TargetFilePath, 'rb') as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, my_text_string, laparams=LAParams())
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)

    #Set working text
    WorkingText = my_text_string.getvalue()
    #Split block of text into strings
    split_list = list(WorkingText)
    #Format split list to remove unnecessary values
    formatted_list = [x for x in split_list if x != '\xa0' and x != '\x0c' and x != '\xad' and x != '\n']
    #print(formatted_list)
    rejoined_list = "".join(formatted_list)
    #print(rejoined_list)
    #Set up temporary list
    temp_list = ["",",","",",","","\n"]

    #FIND FIRST VALUE
    #loop through all values in rejoined list
    for count, i in enumerate(rejoined_list):
        #check if value has already been populated
        if temp_list[0] == '':
            #check for F
            if i == 'F':
                #check second character is digit
                if rejoined_list[count+1].isdigit():
                    temp_list[0] = rejoined_list[count:count+11]
                    #check if last four characters are digits
                    if temp_list[0][-1].isdigit() and temp_list[0][-2].isdigit() and temp_list[0][-3].isdigit() and temp_list[0][-4].isdigit():
                        pass
                    #check if last three characters are digits
                    elif temp_list[0][-2].isdigit() and temp_list[0][-3].isdigit() and temp_list[0][-4].isdigit():
                        temp_list[0] = temp_list[0][:-1]
                    elif temp_list[0][-3].isdigit() and temp_list[0][-4].isdigit():
                        temp_list[0] = temp_list[0][:-2] 
                    elif temp_list[0][-4].isdigit():
                        temp_list[0] = temp_list[0][:-3]
                    else:
                        temp_list[0] = "Error"                  
        else:
            break

    #FIND SECOND VALUE
    #check if index 2 is unpopulated
    if temp_list[2] == '':
        #find "Zone" in list
        SecondValueLocation = WorkingText.find("Zone")
        #assign result to temp list
        temp_list[2] = WorkingText[SecondValueLocation:SecondValueLocation + 9]

    #FIND THIRD VALUE
    #check for value
    if temp_list [4] == '':
        if "WSPverifiedonline" in rejoined_list :
            temp_list [4] = "WSP verified online"
        else:
            temp_list [4] = "Unverified"
    
    return temp_list

#Iterate through documents in target directory
for i in NewList:
    output.append(find_text_function(WorkingDirectory + "\\" + i))

#print output
#print(output)

#Format output string
output_string = []
for i in output:
    for item in i:
        output_string.append(item)

#Join output string
final_output = "".join(output_string)

#Export list to text file
TextOutput = final_output
with open(WorkingDirectory + "\\" + 'Text Output.csv', 'w+') as MyText:
    MyText.write(str(TextOutput))
    MyText.close()