import pdfrw
from reportlab.pdfgen import canvas
from openpyxl import load_workbook
import sys

def create_overlay(name, order):
    """
    Create the data that will be overlayed on top
    of the form that we want to fill
    """
    c = canvas.Canvas('overlay.pdf')

    c.drawString(270, 465, name)

    if 'Classic Turkey Sandwich served with chips and cookie' in order:
        c.drawString(48, 325, 'X')
    elif 'Spinach Hummus Wrap served with side salad (Vegetarian)' in order:
        c.drawString(49, 262, 'X')
    elif 'Roast Beef Sandwich served with chips and cookie' in order:
        c.drawString(49, 205, 'X')
    elif 'Cobb Salad served with fruit and oat bar (Gluten Free)' in order:
        c.drawString(50, 145, 'X')
    else:
        return False
    c.save()
    return True

def merge_pdfs(form_pdf, overlay_pdf, output):
    """
    Merge the specified fillable form PDF with the
    overlay PDF and save the output
    """
    form = pdfrw.PdfReader(form_pdf)
    olay = pdfrw.PdfReader(overlay_pdf)

    for form_page, overlay_page in zip(form.pages, olay.pages):
        merge_obj = pdfrw.PageMerge()
        overlay = merge_obj.add(overlay_page)[0]
        pdfrw.PageMerge(form_page).add(overlay).render()

    writer = pdfrw.PdfWriter()
    writer.write(output, form)


if __name__ == '__main__':
    if len(sys.argv) < 2 or sys.argv[1] == '--help':
        print('Usage: python3 fill_pdfs.py [excel workbook] [sheet name] [pdf template]')
    else:
        wb = load_workbook(filename=sys.argv[1], read_only=True)
        ws = wb[sys.argv[2]]
        count = 0
        for row in ws.rows:
            row_cells = list(row)
            first_name = str(row_cells[0].value)
            last_name = str(row_cells[1].value)
            order = str(row_cells[2].value)
            if create_overlay(first_name+' '+last_name, order):
                merge_pdfs(sys.argv[3],
                       'overlay.pdf',
                       'orders/'+first_name+last_name+'.pdf')
                count += 1
        print("Finished filling " + str(count) + " order forms.")
