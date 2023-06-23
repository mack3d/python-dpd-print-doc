import os
import time
import xlrd
import re
from fpdf import FPDF,HTMLMixin
from printer import printFile

file_location = "C://Users//Maciej//Desktop//"

class MyFPDF(FPDF, HTMLMixin):
    pass

def searchfile():
    lista = os.listdir(path=file_location)
    test = 0
    for plik in lista:
        if re.search("^Kompensata.*xls$",plik):
            readXls(plik)
            test = 1
            break
    if test == 0:
        print("No files")
        time.sleep(3)

def readXls(plik):
    of = xlrd.open_workbook(file_location+plik)
    osh = of.sheet_by_index(0)
    num_rows = osh.nrows
    num_cols = osh.ncols
    data = []
    for row in range(num_rows):
        row_data = []
        for col in range(num_cols):
            cell_value = osh.cell_value(row, col)
            if cell_value != "":
                row_data.append(cell_value)
            if len(row_data) > 2:
                data.append(row_data)
                row_data = []
        if len(row_data) > 0:
            data.append(row_data)
    createPdf(plik, data);

def parseData(data):
    data[0] = (data[0][0], data[1][0])
    data[1] = (data[2][0], data[3][0], data[4][0])
    data.pop(2)
    data.pop(2)
    data.pop(2)
    for i in range(len(data)):
        row = data[i]
        if str(row[0]).find('Data księgowania') != -1:
            n_col_key = data[i+1]
            n_col_val = data[i+2]
            data[i+1] = n_col_val
            data[i+2] = n_col_key
            i = len(data)
    return data

def createPdf(plik, data):
    data = parseData(data)

    pdf = MyFPDF('P')
    pdf.add_page()
    pdf.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf', uni=True)
    pdf.set_font('DejaVu', '', 10)

    line_height = pdf.font_size * 1.2
    col_width = 55
    for row in data:
        if pdf.get_y() > 275:
            pdf.add_page()
        y = pdf.get_y()
        new_y = y
        x = pdf.get_x()

        for val_in_row in row:
            pdf.multi_cell(w = col_width, h = line_height, txt = str(val_in_row), border = 0, align='L')
            if new_y < pdf.get_y():
                new_y = pdf.get_y()-4
            x = x + col_width
            pdf.set_xy(x, y)

        pdf.set_y(new_y)
        pdf.ln(line_height)
    outputfile = file_location+plik+'.pdf'
    os.remove(file_location+plik)
    pdf.output(outputfile, 'F')
    printFile(outputfile)
    os.remove(outputfile)

if __name__ == '__main__':
    searchfile()