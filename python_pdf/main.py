from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape, A4

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont('Calibri', 'Calibri.ttf'))

existing_pdf = PdfFileReader(open("./aug-27-new.pdf", "rb"))

output = PdfFileWriter()

for pagenum in range(existing_pdf.getNumPages()):
#for pagenum in range(1,2):
    page = existing_pdf.getPage(pagenum)

    OrientationDegrees = page.get('/Rotate')
    print(OrientationDegrees)

    packet = io.BytesIO()
    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=A4)
    can.setFont("Times-Roman", 9)

    csv_title = "CSV-102469-STR-Rev 5 Attachment 3"
    
    if(OrientationDegrees == 0):
        can.rotate(-270)
        can.drawString(700, -570, csv_title)
        can.drawString(765, -582, "Page " + str(pagenum + 1) + " of 40")
    elif(OrientationDegrees == 90):
        can.rotate(-270)
        can.drawString(650, -565, csv_title)
        can.drawString(715, -577, "Page " + str(pagenum + 1) + " of 40")
    elif(OrientationDegrees == 270):
        can.rotate(-90)
        can.drawString(-210, 25, csv_title)
        can.drawString(-145, 13, "Page " + str(pagenum + 1) + " of 40")

        
    #can.drawString(-210, 15, "CSV-102469-STR_Rev 3 Attachment 4")
    can.save()

    #move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    # read your existing PDF

    # add the "watermark" (which is the new pdf) on the existing page
    new_page = new_pdf.getPage(0)
    page.mergePage(new_page)
    output.addPage(page)
# finally, write "output" to a real file
outputStream = open("shane_output.pdf", "wb")
output.write(outputStream)
outputStream.close()
