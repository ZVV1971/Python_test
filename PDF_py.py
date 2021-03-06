﻿from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics 
from reportlab.pdfbase.ttfonts import TTFont 
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import A4
import argparse
import os

CONST_POSITION = 'Директор'  #Position of the signing person
CONST_FIO = 'А.В. Кимленко'  #Name and Initials of the signin person
CONST_PREFIX = '#'	     #Prefix character preceding the name of the file (to distinguish)		

parser = argparse.ArgumentParser(description='Gets PDFs and merges into them '
                                 '"copy valid" inscription '
                                 'with corresponding signatures')
parser.add_argument('-p', dest = 'Position',
                    help = 'Position of the signing person',
                    default = CONST_POSITION, type = str,
                    required = False)
parser.add_argument('-n', dest = 'Name',
                    default = CONST_FIO, type = str,
                    help = 'Family Name and initials of the signing person',
                    required = False)
parser.add_argument('-f', dest = 'FileNames',
                    help = 'PDF files to be processed',
                    nargs='+',
                    required = True)
parser.add_argument('-x', dest = 'Prefix',
                    default = CONST_PREFIX, type = str,
                    help = 'Prefix attached to the file name when processed',
                    required = False)

#this line is necessary to process arguments though
#it should be done in an automatic mode
args = parser.parse_args()

packet = io.BytesIO()

#Set the Arial font
pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
canvas = Canvas(packet, pagesize=A4)
canvas.setFont('Arial', 13)
#Prepare in-memory PDF with inscription
#to be merged into the existing PDFs
canvas.rotate(90)
canvas.drawString(20, -A4[0]+35, args.Name)
canvas.drawString(20, -A4[0]+50, '')
canvas.drawString(20, -A4[0]+65, args.Position)
canvas.drawString(20, -A4[0]+80, "Копия верна")
canvas.save()

# move to the beginning of the StringIO buffer
packet.seek(0)
new_pdf = PdfFileReader(packet)
# read your existing PDFs
for pdf_file in args.FileNames:
    if os.path.exists(pdf_file) and os.path.isfile(pdf_file):
        existing_pdf = PdfFileReader(open(pdf_file, "rb"))
        output = PdfFileWriter()
        # add the "watermark" (which is the new pdf) on the existing page
        for i in range(existing_pdf.getNumPages()):
            page = existing_pdf.getPage(i)
            #Only one page is assumed in an inscription
            page2 = new_pdf.getPage(0)
            #In case source PDF is larger than A4 (as sometimes happens with
            #TMZ certificates) scale it down/up to A4
            page.scaleTo(A4[0],A4[1])
            #Merge "watermark" into A4 source page
            page.mergePage(page2)
            #put to the output stream
            output.addPage(page)
        # finally, write "output" to a real file
        outputStream = open(args.Prefix + '_' + os.path.basename(pdf_file),
                            "wb")
        output.write(outputStream)
        outputStream.close()
        #Inform user about changes made
        print(pdf_file + '--->' + args.Prefix + '_' + os.path.basename(pdf_file))
    else:
        #Signal that wrong file has been fed up
        print(pdf_file, 'wrong file name or path')
