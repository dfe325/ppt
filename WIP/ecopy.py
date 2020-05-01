from docx import Document
import os

from pptx import Presentation
import glob

ppt = Presentation('EARTHRISE.pptx')

shape_li = []

for slide in ppt.slides:
    for shape in slide.shapes:
        if shape.has_text_frame and "E7119" in shape.text:
            #print(shape.text)
            shape_li.append(shape.text)


document = Document()

document.add_paragraph(shape_li)



'''document.add_heading('Document Title', 0)

p = document.add_paragraph('A plain paragraph having some ')
p.add_run('bold').bold = True
p.add_run(' and some ')
p.add_run('italic.').italic = True
'''
document.save('ecopy.docx')
