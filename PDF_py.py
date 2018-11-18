from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics 
from reportlab.pdfbase.ttfonts import TTFont 
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import A4

#Set the Arial font

packet = io.BytesIO()

pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))

canvas = Canvas(packet, pagesize=A4)
canvas.setFont('Arial', 13)
canvas.rotate(90)
canvas.drawString(20, -A4[0]+35, "Копия верна")
canvas.drawString(20, -A4[0]+50, "В.В. Захаренков")
canvas.drawString(20, -A4[0]+65, '')
canvas.drawString(20, -A4[0]+80, "Специалист по продажам")
#canvas.showPage()

canvas.save()

# move to the beginning of the StringIO buffer
packet.seek(0)
new_pdf = PdfFileReader(packet)
# read your existing PDF
existing_pdf = PdfFileReader(open("81801892.pdf", "rb"))
output = PdfFileWriter()
# add the "watermark" (which is the new pdf) on the existing page
for i in range(existing_pdf.getNumPages()):
    page = existing_pdf.getPage(i)
    page2 = new_pdf.getPage(0)
    page.mergePage(page2)
    output.addPage(page)
# finally, write "output" to a real file
outputStream = open("newpdf.pdf", "wb")
output.write(outputStream)
outputStream.close() 
