import os
import sys
import traceback

import pdfrw
import pandas as pd
import reportlab.lib.pagesizes
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject
from gui.MainWindow import Ui_MainWindow
from gui.Disclaimer import Ui_DisclaimerDialog

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
EMPTY_COVER_42_43_PDF_PATH = RESOURCE_DIRECTORY + '\\CoverLetter+3330-42+3330-43combined.pdf'
EMPTY_43_PDF_PATH = RESOURCE_DIRECTORY + '\\3330-43-1.pdf'
DATA_SPREADSHEET_PATH = '1. Personal Information.xlsx'

PERFORMANCE_PATH_ = 'performance.pdf'
RESUME_PATH_ = 'resume.pdf'
SIGNED_3330_43_1_PATH_ = '3330-43-1-signed.pdf'

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


# Creates "watermark" PDFs that are empty PDFs with signatures drawn that will later be overlaid on filled pages
# This function is run early as it is only necessary to run once per program execution, not once per package
# CREDIT: https://stackoverflow.com/questions/2925484/place-image-over-pdf
def _generate_signature_watermark_files(signature_path):
    if signature_path != '':
        canvas_cover = canvas.Canvas(WATERMARK_FILE_COVER_LETTER, pagesize=reportlab.lib.pagesizes.letter)
        canvas_cover.drawImage(signature_path, 72, 330, height=36, preserveAspectRatio=True, anchor='sw')
        canvas_cover.save()

        canvas_42 = canvas.Canvas(WATERMARK_FILE_42, pagesize=reportlab.lib.pagesizes.letter)
        canvas_42.drawImage(signature_path, 120, 489, width=180, height=24, preserveAspectRatio=True, anchor='sw')
        canvas_42.save()

        canvas_43 = canvas.Canvas(WATERMARK_FILE_43, pagesize=reportlab.lib.pagesizes.letter)
        canvas_43.drawImage(signature_path, 88, 92, width=210, height=24, preserveAspectRatio=True, anchor='sw')
        canvas_43.save()


# Overlays signature watermark files onto each page or inserts signed page in place of generated page as necessary
# CREDIT: https://stackoverflow.com/questions/2925484/place-image-over-pdf
def _insert_signatures(form_path, output_path, signed_3330_43_1_path):
    watermark_cover_letter_exists = os.path.isfile(WATERMARK_FILE_COVER_LETTER)
    watermark_42_exists = os.path.isfile(WATERMARK_FILE_42)
    watermark_43_exists = os.path.isfile(WATERMARK_FILE_43)
    signed_3330_43_1_exists = os.path.isfile(signed_3330_43_1_path)

    if watermark_cover_letter_exists or watermark_42_exists or watermark_43_exists or signed_3330_43_1_exists:
        output_file = PdfFileWriter()
        pypdf_set_need_appearances_writer(output_file)
        input_file = PdfFileReader(open(form_path, "rb"))

        # Since we don't know which files/image combos the user may use,
        # we declare the PdfFileReader variables here so we can close any used streams at the end of the function
        watermark_cover = watermark_42 = watermark_43 = signed_3330_43_1_file = None

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

        # No signatures on pages 2 and 3
        output_file.addPage(input_file.getPage(2))
        output_file.addPage(input_file.getPage(3))

        # We prefer a full signed 3330-43-1 page 2 over a generated one
        if signed_3330_43_1_exists:
            signed_3330_43_1_file = PdfFileReader(open(signed_3330_43_1_path, "rb"))
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


def _clean_files():
    if os.path.isfile(RESOURCE_DIRECTORY + '\\42watermark.pdf'):
        os.remove(RESOURCE_DIRECTORY + '\\42watermark.pdf')
    if os.path.isfile(RESOURCE_DIRECTORY + '\\43watermark.pdf'):
        os.remove(RESOURCE_DIRECTORY + '\\43watermark.pdf')
    if os.path.isfile(RESOURCE_DIRECTORY + '\\coverwatermark.pdf'):
        os.remove(RESOURCE_DIRECTORY + '\\coverwatermark.pdf')


# Creates a worker object for the ERR Generator that will later be attached to a separate thread from the UI
# CREDIT: https://realpython.com/python-pyqt-qthread/
class ERRWorker(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(int)
    status = pyqtSignal(str)

    def __init__(self, resume_path, performance_path, signed_43_1_path, signature_img_path):

        super(ERRWorker, self).__init__()

        # Initialization errors cannot be printed during initialization because we don't connect the status signal to
        # the ui's progress_output slot until after initialization is complete. Therefore, we store all errors in
        # this init_errors list and print the list when we call ERRWorker.run().
        self.init_errors = []

        # PATH ERROR CHECKING
        # If the path provided is blank, invalid, or a directory, revert to defaults
        # If the path provided is filled, but invalid or a directory, display an error
        # If the path provided is filled and valid, use it
        # We use this approach for each path to avoid repeating code, but readability is not great
        if resume_path == '' or os.path.isdir(resume_path) or not os.path.isfile(resume_path):
            self.resume_path = RESUME_PATH_
        if resume_path != '' and (os.path.isdir(resume_path) or not os.path.isfile(resume_path)):
            self.init_errors.append("ERROR: The path specified for Resume: " + resume_path
                                    + " is invalid and will not be used. Reverting to default...")
        else:
            self.resume_path = resume_path

        if performance_path == '' or os.path.isdir(performance_path) or not os.path.isfile(performance_path):
            self.performance_path = PERFORMANCE_PATH_
        if performance_path != '' and (os.path.isdir(performance_path) or not os.path.isfile(performance_path)):
            self.init_errors.append("ERROR: The path specified for Performance Plan: " + performance_path
                                    + " is invalid and will not be used. Reverting to default...")
        else:
            self.performance_path = performance_path

        if signed_43_1_path == '' or os.path.isdir(signed_43_1_path) or not os.path.isfile(signed_43_1_path):
            self.signed_43_1_path = SIGNED_3330_43_1_PATH_
        if signed_43_1_path != '' and (os.path.isdir(signed_43_1_path) or not os.path.isfile(signed_43_1_path)):
            self.init_errors.append("ERROR: The path specified for Signed 3330-43-1 page 2: " + signed_43_1_path
                                    + " is invalid and will not be used. Reverting to default...")
        else:
            self.signed_43_1_path = signed_43_1_path

        if signature_img_path == '' or os.path.isdir(signature_img_path) or not os.path.isfile(signature_img_path):
            if os.path.isfile('signature.png'):
                self.signature_img_path = 'signature.png'
            elif os.path.isfile('signature.jpg'):
                self.signature_img_path = 'signature.jpg'
            else:
                self.signature_img_path = ''
        if signature_img_path != '' and (os.path.isdir(signature_img_path) or not os.path.isfile(signature_img_path)):
            self.init_errors.append("ERROR: The path specified for Signature Image: " + signature_img_path
                                    + " is invalid and will not be used. Reverting to default...")
        else:
            self.signature_img_path = signature_img_path

    def print_init_errors(self):
        for error_string in self.init_errors:
            self.status.emit(error_string)
            sys.stderr.write(error_string + "\n")

    # This was originally a static function but now part of ERRWorker class
    # This was moved to allow the function to emit a pyqtSignal(str) that can be received by MainWindow.print_status
    # slot and output to a console in the UI.
    def generate_err(self, resume_path, performance_path, signed_43_1_path, signature_img_path):

        _generate_signature_watermark_files(signature_img_path)

        if os.path.isfile(DATA_SPREADSHEET_PATH):
            data_xls = pd.ExcelFile(DATA_SPREADSHEET_PATH, engine="openpyxl")
            data_backend = data_xls.parse("Backend", header=None, index_col=0, usecols="A,B").fillna('').to_dict()

            if not os.path.exists(OUTPUT_DIRECTORY):
                os.mkdir(OUTPUT_DIRECTORY)
            if not os.path.exists(RESOURCE_DIRECTORY):
                os.mkdir(RESOURCE_DIRECTORY)

            # Excel document supports maximum 20 ERRs at once
            num_err = 0
            for i in range(1, 21):
                if data_backend[1].get("Facility" + str(i)):
                    num_err = i
            if data_backend[1].get("USAJOBS"):
                num_err += 1

            if num_err == 0:
                sys.stderr.write(
                    "ERROR: No desired facilities found. Please verify you have filled out 1. Personal Information.xlsx\n")
                self.status.emit(
                    "ERROR: No desired facilities found. Please verify you have filled out 1. Personal Information.xlsx")

            for i in range(1, num_err + 1):
                if data_backend[1].get("Facility" + str(i)):

                    data = data_xls.parse("PDFKeys" + str(i), header=None, index_col=0).fillna('').to_dict()

                    # Each path represents a "step" in the generation, package will be filled first, then signed, then finalized
                    FILLED_PDF_PATH = OUTPUT_DIRECTORY + "\\" + str(data[1].get("Facility")) + "(temp).pdf"
                    SIGNED_PDF_PATH = OUTPUT_DIRECTORY + "\\" + str(data[1].get("Facility")) + "(signed).pdf"
                    FINAL_OUTPUT_PATH = OUTPUT_DIRECTORY + "\\" + str(data[1].get("Facility")) + ".pdf"

                    single_form_fill(EMPTY_COVER_42_43_PDF_PATH, data[1], FILLED_PDF_PATH)
                    _insert_signatures(FILLED_PDF_PATH, SIGNED_PDF_PATH, signed_43_1_path)

                    # concat_paths represents the structure of an ERR package
                    # We start with the filled/signed portion including Cover letter, 3330-42, and 3330-43-1
                    # Then we attach a resume, then a performance plan, in that order
                    if os.path.isfile(SIGNED_PDF_PATH):
                        concat_paths = [SIGNED_PDF_PATH]
                    else:
                        concat_paths = [FILLED_PDF_PATH]

                    if os.path.isfile(resume_path):
                        concat_paths.append(resume_path)
                    if os.path.isfile(performance_path):
                        concat_paths.append(performance_path)

                    # Here we concatenate everything in concat_paths into a single file
                    # Since we previously used temporary placeholder files (temp) and (signed).pdf, we remove them if present
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
                            if os.path.isfile(FILLED_PDF_PATH):
                                os.remove(FILLED_PDF_PATH)
                        else:
                            os.rename(FILLED_PDF_PATH, FINAL_OUTPUT_PATH)

                    self.status.emit("Processed: " + str(data[1].get("Facility")))
                    self.progress.emit(int(100 * (i / num_err)))
                    print("Processed: " + str(data[1].get("Facility")))

            if data_backend[1].get("USAJOBS"):
                data = data_xls.parse("PDFKeysUSAJOBS", header=None, index_col=0).fillna('').to_dict()

                # Each path represents a "step" in the generation, package will be filled first, then finalized
                FILLED_PDF_PATH = "Filled USAJOBS 3330-43-1\\" + str(data[1].get("Facility")) + "(temp).pdf"
                FINAL_OUTPUT_PATH = "Filled USAJOBS 3330-43-1\\" + str(data[1].get("Facility")) + ".pdf"
                if not os.path.isdir("Filled USAJOBS 3330-43-1"):
                    os.mkdir("Filled USAJOBS 3330-43-1")

                single_form_fill(EMPTY_43_PDF_PATH, data[1], FILLED_PDF_PATH)

                # Reuses code from _insert_signatures function above
                # Requires abstraction
                watermark_43 = signed_3330_43_1_file = None
                if os.path.isfile(signed_43_1_path) or os.path.isfile(WATERMARK_FILE_43):
                    output_file = PdfFileWriter()
                    pypdf_set_need_appearances_writer(output_file)
                    input_file = PdfFileReader(open(FILLED_PDF_PATH, "rb"))

                    output_file.addPage(input_file.getPage(0))

                    if os.path.isfile(signed_43_1_path):
                        signed_3330_43_1_file = PdfFileReader(open(signed_43_1_path, "rb"))
                        output_file.addPage(signed_3330_43_1_file.getPage(0))
                    elif os.path.isfile(WATERMARK_FILE_43):
                        watermark_43 = PdfFileReader(open(WATERMARK_FILE_43, "rb"))

                        page_43 = input_file.getPage(1)
                        page_43.mergePage(watermark_43.getPage(0))
                        output_file.addPage(page_43)

                    with open(FINAL_OUTPUT_PATH, "wb") as outputStream:
                        output_file.write(outputStream)

                    input_file.stream.close()
                    if signed_3330_43_1_file is not None:
                        signed_3330_43_1_file.stream.close()
                    elif watermark_43 is not None:
                        watermark_43.stream.close()

                if os.path.isfile(FINAL_OUTPUT_PATH):
                    os.remove(FILLED_PDF_PATH)
                else:
                    os.rename(FILLED_PDF_PATH, FINAL_OUTPUT_PATH)

                self.status.emit("Processed: USAJOBS Announcement " + str(data[1].get("Vacancy Number")) + " for "
                                 + str(data[1].get("Facility")))
                self.progress.emit(100)
                print("Processed: USAJOBS Announcement " + str(data[1].get("Vacancy Number")) + " for "
                      + str(data[1].get("Facility")))

            _clean_files()
        else:
            sys.stderr.write("ERROR: 1. Personal Information.xlsx not found! Aborting!\n")
            self.status.emit("ERROR: 1. Personal Information.xlsx not found! Aborting!")

    def run(self):
        # Initialization errors cannot be printed during initialization because we don't connect the status signal to
        # the ui's progress_output slot until after initialization is complete. Therefore, we allow initialization to
        # complete, log all errors, then output them just prior to runtime
        self.print_init_errors()

        try:
            self.generate_err(self.resume_path, self.performance_path, self.signed_43_1_path, self.signature_img_path)
        except Exception as e:
            self.status.emit(str(e))
            traceback.print_exc()

        self.finished.emit()


class MainWindow(QMainWindow, Ui_MainWindow):
    terms_accepted = False

    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)

        self.run_button.clicked.connect(self.run_err_generator)

        self.resume_browse.clicked.connect(self.open_resume_pdf)
        self.performance_browse.clicked.connect(self.open_performance_pdf)
        self.signature_browse.clicked.connect(self.open_signature_image)
        self.signed_43_1_browse.clicked.connect(self.open_43_1_pdf)

    def open_performance_pdf(self):
        performance_path = QFileDialog.getOpenFileName(self, "Select Performance Plan PDF", "",
                                                       "PDF Files (*.pdf);;All Files (*)")
        self.performance_line_edit.setText(performance_path[0])

    def open_resume_pdf(self):
        resume_path = QFileDialog.getOpenFileName(self, "Select Resume PDF", "", "PDF Files (*.pdf);;All Files (*)")
        self.resume_line_edit.setText(resume_path[0])

    def open_43_1_pdf(self):
        # We declare the variable here as None because the user may choose to not accept the terms in DisclaimerDialog.
        # This results in a "referencing local variable before declaration" error when we try to setText at the end
        path_3330_43_1 = None
        if not self.terms_accepted:
            disclaimer_dialog = DisclaimerDialog()
            disclaimer_dialog.buttonBox.accepted.connect(self.set_terms_accepted_true)
            disclaimer_dialog.buttonBox.rejected.connect(self.set_terms_accepted_false)
            disclaimer_dialog.exec()

        if self.terms_accepted:
            path_3330_43_1 = QFileDialog.getOpenFileName(self, "Select Signed 3330-43-1 PDF", "",
                                                         "PDF Files (*.pdf);;All Files (*)")

        if path_3330_43_1 is not None:
            self.signed_43_1_line_edit.setText(path_3330_43_1[0])

    def open_signature_image(self):
        signature_path = QFileDialog.getOpenFileName(self, "Select Signature Image", "",
                                                     "Image Files (*.png *.jpg *.jpeg *.gif);;All Files (*)")

        self.signature_line_edit.setText(signature_path[0])

    def report_progress(self, progress):
        self.progress_bar.setValue(progress)

    def print_status(self, string):
        self.progress_output.append(string)

    def set_terms_accepted_true(self):
        self.terms_accepted = True

    def set_terms_accepted_false(self):
        self.terms_accepted = False

    def run_err_generator(self):
        self.thread = QThread()
        self.worker = ERRWorker(self.resume_line_edit.text(), self.performance_line_edit.text(),
                                self.signed_43_1_line_edit.text(), self.signature_line_edit.text())

        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.progress.connect(self.report_progress)
        self.worker.status.connect(self.print_status)

        self.thread.start()

        self.run_button.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.run_button.setEnabled(True)
        )


class DisclaimerDialog(QDialog, Ui_DisclaimerDialog):
    def __init__(self):
        super(DisclaimerDialog, self).__init__()
        self.setupUi(self)

        self.ok_button = self.buttonBox.button(QDialogButtonBox.StandardButton.Ok)
        self.ok_button.setEnabled(False)
        self.checkBox.toggled.connect(self.ok_button.setEnabled)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    app.exec()
