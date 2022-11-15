import PyPDF2
import camelot
import ctypes
import numpy as np
from ctypes.util import find_library
import pandas as pd
from tabulate import tabulate
import pdf2image
import xlsxwriter
import pytesseract
from pytesseract import Output, TesseractError
import cv2
import PIL
import os
import shutil
import re
from PIL import Image
from numpy import asarray
from matplotlib import pyplot as plt

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
poppler_path = r"C:\Program Files (x86)\poppler-22.04.0\Library\bin"


def pdf2txt(pdf_path):
    file = open(pdf_path, 'rb')
    fileReader = PyPDF2.PdfFileReader(file)
    num_pages = fileReader.numPages
    text_list = []
    for i in range(0, num_pages):
        text = str(fileReader.getPage(i).extract_text()).replace("\n", "").split(" ")
        while ("" in text):
            text.remove("")
        text_list.append(text)
    return text_list

def pdf2tbl(pdf_path, page):
    a = camelot.read_pdf(pdf_path, pages=str(page))
    print(a[2].parsing_report)
    return a[0].df
    # print(a[0].df.to_excel("test.xlsx"))

def pdfimg2text(pdf_path_ocr, page):
    images = pdf2image.convert_from_path(pdf_path=pdf_path_ocr, poppler_path=poppler_path, first_page=page, last_page=page)
    img = asarray(images[0])
    img = cv2.resize(img, None, fx=1, fy=1)
    # plt.imshow(img, interpolation='nearest')
    # plt.show()
    ocr_dict = pytesseract.image_to_data(img, lang='eng', output_type=Output.DICT, config='--psm 4')
    ocr_text = ocr_dict['text']
    while ('' in ocr_text):
        ocr_text.remove('')
    while (' ' in ocr_text):
        ocr_text.remove(' ')
    return ocr_text

def folder2file(path):
    return os.listdir(path)

def files_sort(files, pdf_root_path):
    for i in files:
        page_list = pdf2txt(pdf_root_path+"\\"+str(i))
        print(page_list[0])
        if page_list[0] == []:
            shutil.copy(pdf_root_path + "\\" + str(i), pdf_root_path + "\\" + "OCR\\" + str(i))
        else:
            shutil.copy(pdf_root_path + "\\" + str(i), pdf_root_path + "\\" + "PDF\\" + str(i))

def main():
    try:
        wr = 0
        wc = 0
        # worksheet.write(0,1, "TTEST")
        # workbook.close()


        pdf_path = "BREGENEPD00017.pdf"
        pdf_path_ocr = "BREGENEPD000418.pdf"

        pdf_root_path = r"C:\Users\Ridheesh\Desktop\Firstplanit_Webscraper\BRE_EPD_Hub"
        pdf_OCR_path = r"C:\Users\Ridheesh\Desktop\Firstplanit_Webscraper\BRE_EPD_Hub\OCR"
        pdf_PDF_path = r"C:\Users\Ridheesh\Desktop\Firstplanit_Webscraper\BRE_EPD_Hub\PDF"

        files = folder2file(pdf_root_path)
        OCR_files = folder2file(pdf_OCR_path)
        PDF_files = folder2file(pdf_PDF_path)

        # text = pdfimg2text(pdf_path_ocr, 1)
        # print(text)
        # e_i = text.index('declaration') + 3
        # e_e = text.index('Company')
        # print(text[e_i:e_e])



        workbook = xlsxwriter.Workbook("EPD_Scrape.xlsx")
        worksheet = workbook.add_worksheet()
        for i in OCR_files:
            try:
                pages = list(pdf2txt(pdf_OCR_path + "\\" + str(i)))
                # page = pages[0]
                page = pdfimg2text(pdf_OCR_path + "\\" + str(i), 1)
                # e_i = page.index('No.:')+1
                # EPD_NO = page[e_i]
                # print(EPD_NO)
                # worksheet.write(wr, 0, EPD_NO)
                worksheet.write(wr, 0, str(i))

                e_i = page.index('provided')+2
                e_e = page.index('accordance')-2
                provided = page[e_i:e_e]
                print(provided)
                provided = ' '.join(provided)
                worksheet.write(wr, 1, str(provided))

                e_i = page.index('declaration')+3
                e_e = page.index('Company')
                name = page[e_i:e_e]
                print(name)
                name = ' '.join(name)
                worksheet.write(wr, 2, str(name))

                e_i = page.index('Address') + 1
                e_e = page.index('Signed') - 5
                address = page[e_i:e_e]
                print(address)
                address = ' '.join(address)
                worksheet.write(wr, 3, str(address))

                e_i = page.index('Signed') - 3
                e_e = page.index('Signed')
                date_I = page[e_i:e_e]
                print(date_I)
                date_I = ' '.join(date_I)
                worksheet.write(wr, 4, str(date_I))

                e_i = page.index('Expiry') - 7
                e_e = page.index('Expiry') - 4
                date_E = page[e_i:e_e]
                print(date_E)
                date_E = ' '.join(date_E)
                worksheet.write(wr, 5, str(date_E))

                tp = 1
                for p in pages:
                    # print(p)
                    if "site(s)" in p:
                        e_i = p.index('site(s)') + 1
                        e_e = p.index('Product:') - 1
                        site = p[e_i:e_e]
                        print(site)
                        site = ' '.join(site)
                        worksheet.write(wr, 6, str(site))

                    if "equiv." in p and "GWP" in p:
                        Table = pdf2tbl(pdf_OCR_path + "\\" + str(i), tp)
                        table = Table.values.tolist()
                        print(np.shape(table))
                        print(table)
                        for c, r in enumerate(table):
                            if "GWP" in r:
                                GWP_i = (c, r.index("GWP"))
                            if "A1-3" in r:
                                index2 = (c, r.index("A1-3"))
                            if "AP" in r:
                                AP_i = (c, r.index("AP"))
                            if "EP" in r:
                                EP_i = (c, r.index("EP"))

                        GWP = table[index2[0]][GWP_i[1]]
                        GWP_u = table[GWP_i[0] + 1][GWP_i[1]]

                        AP = table[index2[0]][AP_i[1]]
                        AP_u = table[AP_i[0] + 1][AP_i[1]]

                        EP = table[index2[0]][EP_i[1]]
                        EP_u = table[EP_i[0] + 1][EP_i[1]]

                        print(GWP, GWP_u, AP, AP_u, EP, EP_u)

                        GWP = ' '.join(GWP)
                        worksheet.write(wr, 7, str(GWP))

                        GWP_u = ' '.join(GWP_u)
                        worksheet.write(wr, 8, str(GWP_u))

                        AP = ' '.join(AP)
                        worksheet.write(wr, 9, str(AP))

                        AP_u = ' '.join(AP_u)
                        worksheet.write(wr, 10, str(AP_u))

                        EP = ' '.join(EP)
                        worksheet.write(wr, 11, str(EP))

                        EP_u = ' '.join(EP_u)
                        worksheet.write(wr, 12, str(EP_u))

                    elif "PERT" in p and "PENRT" in p:
                        Table = pdf2tbl(pdf_OCR_path + "\\" + str(i), tp)
                        table = Table.values.tolist()
                        print(np.shape(table))
                        print(table)
                        for c, r in enumerate(table):
                            if "PERT" in r:
                                PERT_i = (c, r.index("PERT"))
                            if "A1-3" in r:
                                index2 = (c, r.index("A1-3"))
                            if "PENRT" in r:
                                PENRT_i = (c, r.index("PENRT"))

                        PERT = table[index2[0]][PERT_i[1]]
                        PERT_u = table[PERT_i[0] + 1][PERT_i[1]]

                        PENRT = table[index2[0]][PENRT_i[1]]
                        PENRT_u = table[PENRT_i[0] + 1][PENRT_i[1]]

                        print("PAGE2: ", PERT, PERT_u, PENRT, PENRT_u)

                        PERT = ' '.join(PERT)
                        worksheet.write(wr, 13, str(PERT))

                        PERT_u = ' '.join(PERT_u)
                        worksheet.write(wr, 14, str(PERT_u))

                        PENRT = ' '.join(PENRT)
                        worksheet.write(wr, 15, str(PENRT))

                        PENRT_u = ' '.join(PENRT_u)
                        worksheet.write(wr, 16, str(PENRT_u))

                        wr += 1

                    tp += 1


            except Exception as e:
                    print(i)
                    print("ERROR: " + str(e))
                    continue

        workbook.close()

    except Exception as e:
        print("ERROR: " + str(e))

if __name__=="__main__":
    main()