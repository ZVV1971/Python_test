from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

packet = io.BytesIO()

# create a new PDF with Reportlab
can = canvas.Canvas(packet, pagesize=letter)
can.drawString(100, 100, "Hello world")
can.save()

# move to the beginning of the StringIO buffer
packet.seek(0)
new_pdf = PdfFileReader(packet)
# read your existing PDF
existing_pdf = PdfFileReader(open("mypdf.pdf", "rb"))
output = PdfFileWriter()
# add the "watermark" (which is the new pdf) on the existing page
page = existing_pdf.getPage(0)
page2 = new_pdf.getPage(0)
page.mergePage(page2)
output.addPage(page)
# finally, write "output" to a real file
outputStream = open("newpdf.pdf", "wb")
output.write(outputStream)
outputStream.close() 