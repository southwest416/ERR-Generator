import os
import sys

import pdfrw
import pandas as pd
import reportlab.lib.pagesizes
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject

from PyQt6.QtGui import *
from PyQt6.QtCore import *
from PyQt6.QtWidgets import *

# PDFRW CONSTANTS
ANNOT_KEY = '/Annots'
ANNOT_FIELD_KEY = '/T'
ANNOT_VAL_KEY = '/V'
ANNOT_RECT_KEY = '/Rect'
SUBTYPE_KEY = '/Subtype'
WIDGET_SUBTYPE_KEY = '/Widget'

# GLOBAL CONSTANTS
OUTPUT_DIRECTORY = 'Filled ERRs'
RESOURCE_DIRECTORY = 'resources'

WATERMARK_FILE_COVER_LETTER = RESOURCE_DIRECTORY + '\\coverwatermark.pdf'
WATERMARK_FILE_42 = RESOURCE_DIRECTORY + '\\42watermark.pdf'
WATERMARK_FILE_43 = RESOURCE_DIRECTORY + '\\43watermark.pdf'

# GLOBAL COUNT VARIABLE
num_err = 0


# PDFRW FUNCTIONS
# CREDIT: https://github.com/WestHealth/pdf-form-filler
def _text_form(annotation, value):
    # Improper way to replace "False" value with a 0
    # Should fix this eventually but probably won't
    if value == False:
        value = 0
    pdfstr = pdfrw.objects.pdfstring.PdfString.encode(str(value))
    annotation.update(pdfrw.PdfDict(V=pdfstr, AS=pdfstr))


def _checkbox(annotation, value, export=None):
    if export:
        export = '/' + export
    else:
        keys = annotation['/AP']['/N'].keys()
        if ['/Off'] in keys:
            keys.remove('/Off')
        export = keys[0]
    if value:
        annotation.update(pdfrw.PdfDict(V=export, AS=export))
    else:
        if '/V' in annotation:
            del annotation['/V']
        if '/AS' in annotation:
            del annotation['/AS']


def _radio_button(annotation, value):
    for each in annotation['/Kids']:
        # determine the export value of each kid
        keys = each['/AP']['/N'].keys()
        if ['/Off'] in keys:
            keys.remove('/Off')
        export = keys[0]

        if f'/{value}' == export:
            val_str = pdfrw.objects.pdfname.BasePdfName(f'/{value}')
        else:
            val_str = pdfrw.objects.pdfname.BasePdfName(f'/Off')
        each.update(pdfrw.PdfDict(AS=val_str))

    annotation.update(pdfrw.PdfDict(V=pdfrw.objects.pdfname.BasePdfName(f'/{value}')))


def _combobox(annotation, value):
    export = None
    for each in annotation['/Opt']:
        if each[1].to_unicode() == value:
            export = each[0].to_unicode()
    if export is None:
        raise KeyError(f"Export Value: {value} Not Found")
    pdfstr = pdfrw.objects.pdfstring.PdfString.encode(export)
    annotation.update(pdfrw.PdfDict(V=pdfstr, AS=pdfstr))


def _listbox(annotation, values):
    pdfstrs = []
    for value in values:
        export = None
        for each in annotation['/Opt']:
            if each[1].to_unicode() == value:
                export = each[0].to_unicode()
        if export is None:
            raise KeyError(f"Export Value: {value} Not Found")
        pdfstrs.append(pdfrw.objects.pdfstring.PdfString.encode(export))
    annotation.update(pdfrw.PdfDict(V=pdfstrs, AS=pdfstrs))


def _field_type(annotation):
    ft = annotation['/FT']
    ff = annotation['/Ff']

    if ft == '/Tx':
        return 'text'
    if ft == '/Ch':
        if ff and int(ff) & 1 << 17:  # test 18th bit
            return 'combo'
        else:
            return 'list'
    if ft == '/Btn':
        if ff and int(ff) & 1 << 15:  # test 16th bit
            return 'radio'
        else:
            return 'checkbox'


def _blank_page(w, h):
    blank = pdfrw.PageMerge()
    blank.mbox = [0, 0, w * 72, h * 72]
    blank = blank.render()
    return blank


def pdf_form_info(in_pdf):
    info = []
    for page in in_pdf.pages:
        annotations = page['/Annots']
        if annotations is None:
            continue
        for annotation in annotations:
            choices = None
            if annotation['/Subtype'] == '/Widget':
                if not annotation['/T']:
                    annotation = annotation['/Parent']
                key = annotation['/T'].to_unicode()
                ft = _field_type(annotation)
                value = annotation['/V']
                if ft == 'radio':
                    value = value[1:]
                    choices = []
                    for each in annotation['/Kids']:
                        keys = each['/AP']['/N'].keys()
                        if not keys[0][1:] in choices:
                            choices.append(keys[0][1:])
                elif ft == 'list' or ft == 'combo':
                    choices = [each[1].to_unicode() for each in annotation['/Opt']]
                    values = []
                    for each in annotation['/Opt']:
                        if each[0] in value:
                            values.append(each[1].to_unicode())
                    value = values
                else:
                    if value:
                        value = value.to_unicode()
                out = dict(name=key, type=ft)
                if value:
                    out['value'] = value
                if choices:
                    out['choices'] = choices
                info.append(out)
    return info


def fill_form(in_pdf, data, suffix=None):
    fillers = {'checkbox': _checkbox,
               'list': _listbox,
               'text': _text_form,
               'combo': _combobox,
               'radio': _radio_button}
    for page in in_pdf.pages:
        annotations = page['/Annots']
        if annotations is None:
            continue
        for annotation in annotations:

            if annotation['/Subtype'] == '/Widget':
                if not annotation['/T']:
                    annotation = annotation['/Parent']
                key = annotation['/T'].to_unicode()
                if key in data:
                    ft = _field_type(annotation)
                    fillers[ft](annotation, data[key])
                    if suffix:
                        new_T = pdfrw.objects.pdfstring.PdfString.encode(key + suffix)
                        annotation.update(pdfrw.PdfDict(T=new_T))
        in_pdf.Root.AcroForm.update(
            pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
    return in_pdf


def single_form_fill(in_file, data, out_file):
    pdf = pdfrw.PdfReader(in_file)
    out_pdf = fill_form(pdf, data)
    pdfrw.PdfWriter().write(out_file, out_pdf)


# PDFRW Safe Concat Function
# CREDIT: https://stackoverflow.com/questions/57008782/pypdf2-pdffilemerger-loosing-pdf-module-in-merged-file
def concatenate_pdfrw(pdf_files, output_filename):
    output = pdfrw.PdfWriter()
    num = 0
    output_acroform = None
    for pdf in pdf_files:
        input1 = pdfrw.PdfReader(pdf, verbose=False)
        output.addpages(input1.pages)
        if pdfrw.PdfName('AcroForm') in input1[pdfrw.PdfName('Root')].keys():  # Not all PDFs have an AcroForm node
            source_acroform = input1[pdfrw.PdfName('Root')][pdfrw.PdfName('AcroForm')]
            if pdfrw.PdfName('Fields') in source_acroform:
                output_formfields = source_acroform[pdfrw.PdfName('Fields')]
            else:
                output_formfields = []
            num2 = 0
            for form_field in output_formfields:
                key = pdfrw.PdfName('T')
                old_name = form_field[key].replace('(', '').replace(')', '')  # Field names are in the "(name)" format
                form_field[key] = 'FILE_{n}_FIELD_{m}_{on}'.format(n=num, m=num2, on=old_name)
                num2 += 1
            if output_acroform == None:
                # copy the first AcroForm node
                output_acroform = source_acroform
            else:
                for key in source_acroform.keys():
                    # Add new AcroForms keys if output_acroform already existing
                    if key not in output_acroform:
                        output_acroform[key] = source_acroform[key]
                # Add missing font entries in /DR node of source file
                if (pdfrw.PdfName('DR') in source_acroform.keys()) and (
                        pdfrw.PdfName('Font') in source_acroform[pdfrw.PdfName('DR')].keys()):
                    if pdfrw.PdfName('Font') not in output_acroform[pdfrw.PdfName('DR')].keys():
                        # if output_acroform is missing entirely the /Font node under an existing /DR, simply add it
                        output_acroform[pdfrw.PdfName('DR')][pdfrw.PdfName('Font')] = \
                            source_acroform[pdfrw.PdfName('DR')][
                                pdfrw.PdfName('Font')]
                    else:
                        # else add new fonts only
                        for font_key in source_acroform[pdfrw.PdfName('DR')][pdfrw.PdfName('Font')].keys():
                            if font_key not in output_acroform[pdfrw.PdfName('DR')][pdfrw.PdfName('Font')]:
                                output_acroform[pdfrw.PdfName('DR')][pdfrw.PdfName('Font')][font_key] = \
                                    source_acroform[pdfrw.PdfName('DR')][pdfrw.PdfName('Font')][font_key]
            if pdfrw.PdfName('Fields') not in output_acroform:
                output_acroform[pdfrw.PdfName('Fields')] = output_formfields
            else:
                # Add new fields
                output_acroform[pdfrw.PdfName('Fields')] += output_formfields
        num += 1
    output.trailer[pdfrw.PdfName('Root')][pdfrw.PdfName('AcroForm')] = output_acroform
    output.write(output_filename)


# PyPDF2 Set Need Appearances Function (Fixes fields not being visible when viewing generated PDF)
# CREDIT: https://stackoverflow.com/questions/58898542/update-a-fillable-pdf-using-pypdf2
def pypdf_set_need_appearances_writer(writer: PdfFileWriter):
    # See 12.7.2 and 7.7.2 for more info: http://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/PDF32000_2008.pdf
    try:
        catalog = writer._root_object
        # get the AcroForm tree
        if "/AcroForm" not in catalog:
            writer._root_object.update({
                NameObject("/AcroForm"): IndirectObject(len(writer._objects), 0, writer)
            })

        need_appearances = NameObject("/NeedAppearances")
        writer._root_object["/AcroForm"][need_appearances] = BooleanObject(True)
        # del writer._root_object["/AcroForm"]['NeedAppearances']
        return writer

    except Exception as e:
        print('set_need_appearances_writer() catch : ', repr(e))
        return writer


def get_signature():
    if os.path.isfile('signature.png'):
        return 'signature.png'
    elif os.path.isfile('signature.jpg'):
        return 'signature.jpg'
    return ''


def get_sup_signature():
    if os.path.isfile('supsignature.png'):
        return 'supsignature.png'
    elif os.path.isfile('supsignature.jpg'):
        return 'supsignature.jpg'
    return ''


# Creates watermark PDFs for drawn signatures to overlay before generating & signing packages
# This function is separate as it is only necessary to run once per program execution, not once per package
# CREDIT: https://stackoverflow.com/questions/2925484/place-image-over-pdf
def init_signatures(signature_path, sup_signature_path):
    # IF A SIGNATURE IS ATTACHED, CREATE WATERMARK FILES FOR EACH PAGE THAT NEEDS TO BE SIGNED
    # THESE FILES WILL LATER BE OVERLAID ON THE FILLED ERR DOCUMENTS
    if signature_path != '':
        # CREATE SIGNATURE WATERMARK FOR COVER LETTER PAGE
        canvas_cover = canvas.Canvas(WATERMARK_FILE_COVER_LETTER, pagesize=reportlab.lib.pagesizes.letter)
        canvas_cover.drawImage(signature_path, 72, 330, height=36, preserveAspectRatio=True, anchor='sw')
        canvas_cover.save()
        # CREATE SIGNATURE WATERMARK FOR 3330-42
        canvas_42 = canvas.Canvas(WATERMARK_FILE_42, pagesize=reportlab.lib.pagesizes.letter)
        canvas_42.drawImage(signature_path, 120, 489, width=180, height=24, preserveAspectRatio=True, anchor='sw')
        canvas_42.save()
        # CREATE SIGNATURE WATERMARK FOR 3330-43-1
        canvas_43 = canvas.Canvas(WATERMARK_FILE_43, pagesize=reportlab.lib.pagesizes.letter)
        canvas_43.drawImage(signature_path, 88, 92, width=210, height=24, preserveAspectRatio=True, anchor='sw')
        if sup_signature_path != '':
            canvas_43.drawImage(sup_signature_path, 378, 92, width=210, height=24, preserveAspectRatio=True,
                                anchor='sw')
        canvas_43.save()

    # IF ONLY A SUPERVISOR SIGNATURE IS ATTACHED, CREATE A WATERMARK FILE ONLY FOR THE 3330-43-1
    elif sup_signature_path != '':
        canvas_43 = canvas.Canvas(WATERMARK_FILE_43, pagesize=reportlab.lib.pagesizes.letter)
        canvas_43.drawImage(sup_signature_path, 378, 92, width=210, height=24, preserveAspectRatio=True,
                            anchor='sw')
        canvas_43.save()


# Overlays signature watermarks onto filled packages
# CREDIT: https://stackoverflow.com/questions/2925484/place-image-over-pdf
def insert_signatures(form_path, output_path):
    SIGNED_3330_43_1_PATH = '3330-43-1-signed.pdf'

    output_file = input_file = None
    watermark_cover = watermark_42 = watermark_43 = signed_3330_43_1_file = None

    watermark_cover_letter_exists = os.path.isfile(WATERMARK_FILE_COVER_LETTER)
    watermark_42_exists = os.path.isfile(WATERMARK_FILE_42)
    watermark_43_exists = os.path.isfile(WATERMARK_FILE_43)
    signed_3330_43_1_exists = os.path.isfile(SIGNED_3330_43_1_PATH)

    if watermark_cover_letter_exists or watermark_42_exists or watermark_43_exists or signed_3330_43_1_exists:
        output_file = PdfFileWriter()
        pypdf_set_need_appearances_writer(output_file)
        input_file = PdfFileReader(open(form_path, "rb"))

        if watermark_cover_letter_exists:
            watermark_cover = PdfFileReader(open(WATERMARK_FILE_COVER_LETTER, "rb"))

            cover_page = input_file.getPage(0)
            cover_page.mergePage(watermark_cover.getPage(0))
            output_file.addPage(cover_page)
        else:
            output_file.addPage(input_file.getPage(0))

        if watermark_42_exists:
            watermark_42 = PdfFileReader(open(WATERMARK_FILE_42, "rb"))

            page_42 = input_file.getPage(1)
            page_42.mergePage(watermark_42.getPage(0))
            output_file.addPage(page_42)
        else:
            output_file.addPage(input_file.getPage(1))

        # ADD FILLED 3330-42 PAGE 2 & 3330-43-1 PAGE 1 TO OUTPUT FILE
        output_file.addPage(input_file.getPage(2))
        output_file.addPage(input_file.getPage(3))

        if signed_3330_43_1_exists:
            signed_3330_43_1_file = PdfFileReader(open(SIGNED_3330_43_1_PATH, "rb"))
            output_file.addPage(signed_3330_43_1_file.getPage(0))
        elif watermark_43_exists:
            watermark_43 = PdfFileReader(open(WATERMARK_FILE_43, "rb"))

            page_43 = input_file.getPage(4)
            page_43.mergePage(watermark_43.getPage(0))
            output_file.addPage(page_43)
        else:
            output_file.addPage(input_file.getPage(4))

        with open(output_path, "wb") as outputStream:
            output_file.write(outputStream)

        input_file.stream.close()
        if watermark_cover is not None:
            watermark_cover.stream.close()
        if watermark_42 is not None:
            watermark_42.stream.close()
        if watermark_43 is not None:
            watermark_43.stream.close()
        if signed_3330_43_1_file is not None:
            signed_3330_43_1_file.stream.close()


    # # IF A SIGNATURE WATERMARK FILE IS FOUND, OVERLAY WATERMARK FILES FOR ALL 3 SIGNABLE PAGES ONTO PACKAGE
    # if os.path.isfile(WATERMARK_FILE_COVER_LETTER):
    #     # CREATE INPUT READER & OUTPUT WRITER
    #     output_file = PdfFileWriter()
    #     pypdf_set_need_appearances_writer(output_file)
    #     input_file = PdfFileReader(open(form_path, "rb"))
    #
    #     # CREATE WATERMARK FILE READERS
    #     watermark_cover = PdfFileReader(open(WATERMARK_FILE_COVER_LETTER, "rb"))
    #     watermark_42 = PdfFileReader(open(WATERMARK_FILE_42, "rb"))
    #     watermark_43 = PdfFileReader(open(WATERMARK_FILE_43, "rb"))
    #     signed_3330_43_1_file = None
    #
    #     # GET COVER LETTER PAGE FROM FILLED PACKAGE, OVERLAY WATERMARK FILE, ADD TO OUTPUT FILE
    #     cover_page = input_file.getPage(0)
    #     cover_page.mergePage(watermark_cover.getPage(0))
    #     output_file.addPage(cover_page)
    #
    #     # GET 3330-42 PAGE FROM FILLED PACKAGE, OVERLAY WATERMARK FILE, ADD TO OUTPUT FILE
    #     page_42 = input_file.getPage(1)
    #     page_42.mergePage(watermark_42.getPage(0))
    #     output_file.addPage(page_42)
    #
    #     # ADD FILLED 3330-42 PAGE 2 & 3330-43-1 PAGE 1 TO OUTPUT FILE
    #     output_file.addPage(input_file.getPage(2))
    #     output_file.addPage(input_file.getPage(3))
    #
    #     # CHECK IF FULL SIGNED 3330_43_1 PAGE 2 EXISTS
    #     # IF SO, INSERT HERE, OTHERWISE, DRAW SIGNATURES
    #     if os.path.isfile(SIGNED_3330_43_1_PATH):
    #         # GET SIGNED 3330-43-1 PAGE AND ADD TO OUTPUT FILE
    #         signed_3330_43_1_file = PdfFileReader(open(SIGNED_3330_43_1_PATH, "rb"))
    #         output_file.addPage(signed_3330_43_1_file.getPage(0))
    #     else:
    #         # GET 3330-43-1 PAGE FROM FILLED PACKAGE, OVERLAY WATERMARK FILE, ADD TO OUTPUT FILE
    #         page_43 = input_file.getPage(4)
    #         page_43.mergePage(watermark_43.getPage(0))
    #         output_file.addPage(page_43)
    #
    #     # WRITE ALL PAGES TO OUTPUT_FILE
    #     with open(output_path, "wb") as outputStream:
    #         output_file.write(outputStream)
    #
    #     # CLOSE FILE STREAMS
    #     input_file.stream.close()
    #     watermark_cover.stream.close()
    #     watermark_42.stream.close()
    #     watermark_43.stream.close()
    #     if signed_3330_43_1_file is not None:
    #         signed_3330_43_1_file.stream.close()
    #
    # # IF ONLY A WATERMARK FILE FOR 3330-43-1 IS FOUND, OVERLAY ONLY WATERMARK FILE FOR 3330-43-1 ON PACKAGE
    # elif os.path.isfile(WATERMARK_FILE_43):
    #     # CREATE INPUT READER & OUTPUT WRITER
    #     output_file = PdfFileWriter()
    #     pypdf_set_need_appearances_writer(output_file)
    #     input_file = PdfFileReader(open(form_path, "rb"))
    #
    #     # CREATE WATERMARK FILE READERS
    #     watermark_43 = PdfFileReader(open(WATERMARK_FILE_43, "rb"))
    #     signed_3330_43_1_file = None
    #
    #     # ADD COVER LETTER, 3330-42, AND FIRST PAGE OF 3330-43-1
    #     for i in range(4):
    #         output_file.addPage(input_file.getPage(i))
    #
    #     # CHECK IF FULL SIGNED 3330_43_1 PAGE 2 EXISTS
    #     # IF SO, INSERT HERE, OTHERWISE, DRAW SIGNATURES
    #     if os.path.isfile(SIGNED_3330_43_1_PATH):
    #         # GET SIGNED 3330-43-1 PAGE AND ADD TO OUTPUT FILE
    #         signed_3330_43_1_file = PdfFileReader(open(SIGNED_3330_43_1_PATH, "rb"))
    #         output_file.addPage(signed_3330_43_1_file.getPage(0))
    #     else:
    #         # GET 3330-43-1 PAGE FROM FILLED PACKAGE, OVERLAY WATERMARK FILE, ADD TO OUTPUT FILE
    #         page_43 = input_file.getPage(4)
    #         page_43.mergePage(watermark_43.getPage(0))
    #         output_file.addPage(page_43)
    #
    #     # WRITE ALL PAGES TO OUTPUT_FILE
    #     with open(output_path, "wb") as outputStream:
    #         output_file.write(outputStream)
    #
    #     # CLOSE FILE STREAMS
    #     input_file.stream.close()
    #     watermark_43.stream.close()
    #     if signed_3330_43_1_file is not None:
    #         signed_3330_43_1_file.stream.close()
    #
    # # IF NO WATERMARK FILES FOUND BUT A SIGNED 3330-43-1 PAGE IS, REPLACE THAT PAGE
    # elif os.path.isfile(SIGNED_3330_43_1_PATH):
    #     # CREATE INPUT READER & OUTPUT WRITER
    #     output_file = PdfFileWriter()
    #     pypdf_set_need_appearances_writer(output_file)
    #     input_file = PdfFileReader(open(form_path, "rb"))
    #     signed_3330_43_1_file = PdfFileReader(open(SIGNED_3330_43_1_PATH, "rb"))
    #
    #     # ADD COVER LETTER, 3330-42, AND FIRST PAGE OF 3330-43-1
    #     for i in range(4):
    #         output_file.addPage(input_file.getPage(i))
    #
    #     # GET SIGNED 3330-43-1 PAGE AND ADD TO OUTPUT FILE
    #     output_file.addPage(signed_3330_43_1_file.getPage(0))
    #
    #     # WRITE ALL PAGES TO OUTPUT_FILE
    #     with open(output_path, "wb") as outputStream:
    #         output_file.write(outputStream)
    #
    #     # CLOSE FILE STREAMS
    #     input_file.stream.close()
    #     signed_3330_43_1_file.stream.close()


def clean_files():
    if os.path.isfile(RESOURCE_DIRECTORY + '\\42watermark.pdf'):
        os.remove(RESOURCE_DIRECTORY + '\\42watermark.pdf')
    if os.path.isfile(RESOURCE_DIRECTORY + '\\43watermark.pdf'):
        os.remove(RESOURCE_DIRECTORY + '\\43watermark.pdf')
    if os.path.isfile(RESOURCE_DIRECTORY + '\\coverwatermark.pdf'):
        os.remove(RESOURCE_DIRECTORY + '\\coverwatermark.pdf')


def generate_err(resume_path="resume.pdf", performance_path="performance.pdf", signed_43_1_path="3330-43-1-signed.pdf",
                 signature_img_path=get_signature(), sup_signature_img_path=get_sup_signature(), progress_tracker=None):
    # CONSTANTS & INITIALIZATIONS
    EMPTY_PDF_PATH = RESOURCE_DIRECTORY + '\\CoverLetter+3330-42+3330-43combined.pdf'
    DATA_SPREADSHEET_PATH = '1. Personal Information.xlsx'
    init_signatures(signature_img_path, sup_signature_img_path)

    # CREATES EXCELFILE OBJECT TO PARSE BACKEND DATA INTO USABLE DICT
    data_xls = pd.ExcelFile(DATA_SPREADSHEET_PATH, engine="openpyxl")
    data_backend = data_xls.parse("Backend", header=None, index_col=0, usecols="A,B").fillna('').to_dict()

    # CREATES OUTPUT, BUILD, RESOURCE DIRECTORIES IF NOT EXISTS
    if not os.path.exists(OUTPUT_DIRECTORY):
        os.mkdir(OUTPUT_DIRECTORY)
    if not os.path.exists(RESOURCE_DIRECTORY):
        os.mkdir(RESOURCE_DIRECTORY)

    # COUNT NUMBER OF ERR'S
    for i in range(1, 21):
        if data_backend[1].get("Facility" + str(i)):
            num_err = i

    # ITERATES THROUGH EACH ERR
    for i in range(1, num_err + 1):
        # CHECKS IF FACILITY WAS SELECTED, SKIPS IF NOT
        if data_backend[1].get("Facility" + str(i)):

            # CREATES DICT FOR FACILITY i CONTAINING ALL PDF KEYS TO FILL
            data = data_xls.parse("PDFKeys" + str(i), header=None, index_col=0).fillna('').to_dict()

            # OUTPUT PATH CONSTANTS
            FILLED_PDF_PATH = OUTPUT_DIRECTORY + "\\" + str(data[1].get("Facility")) + "(temp).pdf"
            SIGNED_PDF_PATH = OUTPUT_DIRECTORY + "\\" + str(data[1].get("Facility")) + "(signed).pdf"
            FINAL_OUTPUT_PATH = OUTPUT_DIRECTORY + "\\" + str(data[1].get("Facility")) + ".pdf"

            # FILL FORM AND INSERT SIGNATURES IF PRESENT
            single_form_fill(EMPTY_PDF_PATH, data[1], FILLED_PDF_PATH)
            insert_signatures(FILLED_PDF_PATH, SIGNED_PDF_PATH)

            # CREATES LIST OF FILES TO CONCATENATE INTO ONE PACKAGE
            # IF A SIGNED 3330-43-1 PAGE IS FOUND, INSERT INTO PACKAGE
            if os.path.isfile(SIGNED_PDF_PATH):
                concat_paths = [SIGNED_PDF_PATH]
            else:
                concat_paths = [FILLED_PDF_PATH]

            # CHECK IF RESUME OR PMAS IS ATTACHED, APPEND TO END
            if os.path.isfile(resume_path):
                concat_paths.append(resume_path)
            if os.path.isfile(performance_path):
                concat_paths.append(performance_path)

            # CHECKS IF PERFORMANCE PLAN OR RESUME IS ATTACHED
            # IF SO, CONCATENATES INTO ONE PACKAGE
            # IF FILES NOT FOUND, RENAMES TEMPORARY OUTPUT FILE TO FINAL OUTPUT FILE
            if concat_paths.__len__() > 1:
                concatenate_pdfrw(concat_paths, FINAL_OUTPUT_PATH)
                if os.path.isfile(FILLED_PDF_PATH):
                    os.remove(FILLED_PDF_PATH)
                if os.path.isfile(SIGNED_PDF_PATH):
                    os.remove(SIGNED_PDF_PATH)
            else:
                if os.path.isfile(FINAL_OUTPUT_PATH):
                    os.remove(FINAL_OUTPUT_PATH)
                if os.path.isfile(SIGNED_PDF_PATH):
                    os.rename(SIGNED_PDF_PATH, FINAL_OUTPUT_PATH)
                else:
                    os.rename(FILLED_PDF_PATH, FINAL_OUTPUT_PATH)
            if progress_tracker is not None:
                progress_tracker.emit((i / num_err) * 100)
            print("Processed: " + str(data[1].get("Facility")))

    clean_files()


# Creates a worker object for the ERR Generator
# CREDIT: https://realpython.com/python-pyqt-qthread/
class ERRWorker(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(int)

    def __init__(self, resume_path="", performance_path="", signed_43_1_path="", signature_img_path="",
                 sup_signature_img_path=""):
        self.resume_path = resume_path
        self.performance_path = performance_path
        self.signed_43_1_path = signed_43_1_path
        self.signature_img_path = signature_img_path
        self.sup_signature_img_path = sup_signature_img_path

        super(ERRWorker, self).__init__()

    def run(self):
        generate_err(self.resume_path, self.performance_path, self.signed_43_1_path, self.signature_img_path,
                     self.sup_signature_img_path, self.progress)
        self.finished.emit()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle("ERR Generator")

        self.performancelabel = QLabel("Select a performance plan PDF: ")
        self.resumelabel = QLabel("Select a resume PDF: ")
        self.signaturelabel = QLabel("Select signature image: ")
        self.supsignaturelabel = QLabel("Select supervisor's signature image: ")
        self.signed431label = QLabel("Select signed 3330-43-1 page 2 PDF: ")
        self.label = QLabel("Click to run")

        self.performancebutton = QPushButton("Open File")
        self.performancebutton.clicked.connect(self.open_performance_pdf)
        self.resumebutton = QPushButton("Open File")
        self.resumebutton.clicked.connect(self.open_resume_pdf)
        self.signaturebutton = QPushButton("Open File")
        self.signaturebutton.clicked.connect(self.open_signature_image)
        self.supsignaturebutton = QPushButton("Open File")
        self.supsignaturebutton.clicked.connect(self.open_supsignature_image)
        self.signed431button = QPushButton("Open File")
        self.signed431button.clicked.connect(self.open_431_pdf)

        self.performancelineedit = QLineEdit()
        self.resumelineedit = QLineEdit()
        self.signaturelineedit = QLineEdit()
        self.supsignaturelineedit = QLineEdit()
        self.signed431lineedit = QLineEdit()

        self.button = QPushButton("Run")
        self.button.clicked.connect(self.run_err_generator)

        layout = QGridLayout()
        layout.addWidget(self.performancelabel, 0, 0)
        layout.addWidget(self.resumelabel, 2, 0)
        layout.addWidget(self.signaturelabel, 4, 0)
        layout.addWidget(self.supsignaturelabel, 6, 0)
        layout.addWidget(self.signed431label, 8, 0)

        layout.addWidget(self.performancelineedit, 1, 0)
        layout.addWidget(self.resumelineedit, 3, 0)
        layout.addWidget(self.signaturelineedit, 5, 0)
        layout.addWidget(self.supsignaturelineedit, 7, 0)
        layout.addWidget(self.signed431lineedit, 9, 0)

        layout.addWidget(self.performancebutton, 1, 1)
        layout.addWidget(self.resumebutton, 3, 1)
        layout.addWidget(self.signaturebutton, 5, 1)
        layout.addWidget(self.supsignaturebutton, 7, 1)
        layout.addWidget(self.signed431button, 9, 1)

        layout.addWidget(self.label, 10, 0)
        layout.addWidget(self.button, 10, 1)

        container = QWidget()
        container.setLayout(layout)

        self.setCentralWidget(container)

    def open_performance_pdf(self):
        performance_path = QFileDialog.getOpenFileName(self, "Browse", "", "PDF Files (*.pdf)")

        if performance_path:
            self.performancelineedit.setText(performance_path[0])

    def open_resume_pdf(self):
        resume_path = QFileDialog.getOpenFileName(self, "Browse", "", "PDF Files (*.pdf)")

        if resume_path:
            self.resumelineedit.setText(resume_path[0])

    def open_431_pdf(self):
        path_3330_43_1 = QFileDialog.getOpenFileName(self, "Browse", "", "PDF Files (*.pdf)")

        if path_3330_43_1:
            self.signed431lineedit.setText(path_3330_43_1[0])

    def open_signature_image(self):
        signature_path = QFileDialog.getOpenFileName(self, "Browse", "", "Image Files (*.png, *.jpg)")

        if signature_path:
            self.signaturelineedit.setText(signature_path[0])

    def open_supsignature_image(self):
        sup_signature_path = QFileDialog.getOpenFileName(self, "Browse", "", "Image Files (*.png, *.jpg)")

        if sup_signature_path:
            self.supsignaturelineedit.setText(sup_signature_path[0])

    def report_progress(self):
        pass

    def run_err_generator(self):
        self.thread = QThread()
        self.worker = ERRWorker(self.resumelineedit.text(), self.performancelineedit.text(),
                                self.signed431lineedit.text(), self.signaturelineedit.text(),
                                self.supsignaturelineedit.text())

        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.progress.connect(self.report_progress)

        self.thread.start()

        self.button.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.button.setEnabled(True)
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    app.exec()
