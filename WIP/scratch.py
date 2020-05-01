from docx import Document
import os

from pptx import Presentation
import glob



document = Document()

document.add_paragraph("Hello.")

li = []

for number in range(0, 3):
    li.append('new' + str(number))


print(li)

#creating new sequential list of docs
for item in li:
    document.save(item + '.docx')

#iterate through list of slides in one sheet
#find any shapes that include AST # in question
#create word doc with that AST number
    #clone empty ecopy doc
#pop off beginning AST copy from that shape
#load word doc w copy from that shape
#save word doc w correct formatting for a doc
