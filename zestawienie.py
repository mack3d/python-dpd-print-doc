import os
import time
import xlrd
import re
from fpdf import FPDF,HTMLMixin
import datetime
from printer import printFile

file_location = "C://Users//Maciej//Desktop//"

class MyFPDF(FPDF, HTMLMixin):
    pass

def searchfile():
    lista = os.listdir(path=file_location)
    test = 0
    for plik in lista:
        if re.search("^zestawienie_COD_202.*xls$",plik):
            print(plik)
            readXls(plik)
            test = 1
            break
    if test == 0:
        print("Nie ma pliku")
        time.sleep(3)

def readXls(plik):
    of = xlrd.open_workbook(file_location+plik)
    osh = of.sheet_by_index(0)
    row = 1
    wplaty = list()
    while True:
        if row < osh.nrows:
            wplaty.append([datetime.datetime(*xlrd.xldate_as_tuple(osh.cell_value(row,3), of.datemode)).strftime('%Y-%m-%d'),round(osh.cell_value(row,4),2),osh.cell_value(row,6),osh.cell_value(row,7),osh.cell_value(row,8),osh.cell_value(row,9),osh.cell_value(row,10),osh.cell_value(row,11),datetime.datetime(*xlrd.xldate_as_tuple(osh.cell_value(row,12), of.datemode)).strftime('%Y-%m-%d'),osh.cell_value(row,13)])
        else:
            break
        row+=1
    createpdf(plik,wplaty)

def createpdf(plik,wplaty):
    pdf = MyFPDF('L')
    pdf.add_page()
    pdf.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf', uni=True)
    pdf.set_font('DejaVu', '', 10)
    pdf.set_fill_color(220,220,220)
    wys = 6
    bg = True
    for wplata in wplaty:
        if bg == True:
            bg = False
        else:
            bg = True
        y = pdf.get_y()
        x = pdf.get_x()

        fpage = pdf.page_no()
        pdf.multi_cell(40,wys,str(wplata[0]),fill = bg)

        pdf.set_xy(x+20,y)
        if fpage < pdf.page_no():
            y = 10.00125
            pdf.set_xy(pdf.get_x(),y)
        pdf.multi_cell(25,wys,str(wplata[1]),align = 'C',fill = bg)

        pdf.set_xy(x+45,y)
        pdf.multi_cell(80,wys,wplata[2],align = "L",fill = bg)
        nexty = 0
        if pdf.get_y()-y > 7:
            nexty = pdf.get_y()

        pdf.set_xy(x+125,y)
        pdf.multi_cell(90,wys,str(wplata[3])+" "+wplata[4]+" "+wplata[5], align = "L",fill = bg)
        if pdf.get_y()-y > 7:
            nexty = pdf.get_y()

        pdf.set_xy(x+215,y)
        pdf.multi_cell(25,wys,wplata[6].split(";")[0],fill = bg)
        if pdf.get_y()-y > 7:
            nexty = pdf.get_y()

        pdf.set_xy(x+240,y)
        pdf.multi_cell(25,wys,wplata[6].split(";")[2],fill = bg)
        if pdf.get_y()-y > 7:
            nexty = pdf.get_y()
        
        if nexty != 0:
            pdf.set_y(nexty)
    
    pdf.set_font('DejaVu', '', 12)
    y = pdf.get_y()
    y += 10
    pdf.set_y(y)
    pdf.multi_cell(180,wys,"Zbiorczy przelew nr: "+str(wplaty[0][9])+" data: "+str(wplaty[0][8])+" na kwotę: "+str(wplaty[0][7]))

    outputfile = file_location+plik+'.pdf'
    os.remove(file_location+plik)
    pdf.output(outputfile, 'F')
    printFile(outputfile)
    os.remove(outputfile)

if __name__ == '__main__':
    searchfile()