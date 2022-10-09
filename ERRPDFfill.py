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
from collections import defaultdict

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
# CREDIT: https://stackoverflow.com/questions/2925484/place-image-over-pdf
def _generate_signature_watermark_files(signature_path):

    # If one watermark file exists, all three must exist, so we check if cover letter watermark exists
    if not os.path.isfile(WATERMARK_FILE_COVER_LETTER):
        canvas_cover = canvas.Canvas(WATERMARK_FILE_COVER_LETTER, pagesize=reportlab.lib.pagesizes.letter)
        canvas_cover.drawImage(signature_path, 72, 330, height=36, preserveAspectRatio=True, anchor='sw')
        canvas_cover.save()

        canvas_42 = canvas.Canvas(WATERMARK_FILE_42, pagesize=reportlab.lib.pagesizes.letter)
        canvas_42.drawImage(signature_path, 120, 489, width=180, height=24, preserveAspectRatio=True, anchor='sw')
        canvas_42.save()

        canvas_43 = canvas.Canvas(WATERMARK_FILE_43, pagesize=reportlab.lib.pagesizes.letter)
        canvas_43.drawImage(signature_path, 88, 92, width=210, height=24, preserveAspectRatio=True, anchor='sw')
        canvas_43.save()


# Takes a filled out ERR package and overlays/attaches signatures and appends resume & performance plan
# Returns nothing, all operations are conducted on the files specified in form_path and output_path
def sign_and_append_documents(form_path, output_path, signature_img_path=None, signed_3330_43_1_path=None,
                              resume_path=None, performance_path=None, usajobs_announcement=False):

    output_file_writer = PdfFileWriter()
    pypdf_set_need_appearances_writer(output_file_writer)
    input_file_reader = PdfFileReader(open(form_path, "rb"))

    if os.path.isfile(signature_img_path):
        _generate_signature_watermark_files(signature_img_path)

    if usajobs_announcement:

        # We prefer a signed 3330-43-1 page 2, so we use that first if it exists
        if os.path.isfile(signed_3330_43_1_path):
            output_file_writer.addPage(input_file_reader.getPage(0))

            signed_3330_43_1_reader = PdfFileReader(open(signed_3330_43_1_path, "rb"))
            output_file_writer.addPage(signed_3330_43_1_reader.getPage(0))

            with open(output_path, "wb") as output_stream:
                output_file_writer.write(output_stream)

            input_file_reader.stream.close()
            signed_3330_43_1_reader.stream.close()

            os.remove(form_path)

        # If the signed 3330-43-1 page 2 is not present, we fall back to a signature image
        elif os.path.isfile(signature_img_path):
            output_file_writer.addPage(input_file_reader.getPage(0))

            watermark_43_reader = PdfFileReader(open(WATERMARK_FILE_43, "rb"))
            page_43 = input_file_reader.getPage(1)
            page_43.mergePage(watermark_43_reader.getPage(0))
            output_file_writer.addPage(page_43)

            with open(output_path, "wb") as output_stream:
                output_file_writer.write(output_stream)

            input_file_reader.stream.close()
            watermark_43_reader.stream.close()

            os.remove(form_path)

        else:
            input_file_reader.stream.close()

            if os.path.isfile(output_path):
                os.remove(output_path)
            os.rename(form_path, output_path)

    else:
        temp_path = form_path[0:-4] + " (signed).pdf"

        # Even if a signed 3330-43-1 is present, we still can sign other pages with a signature image
        # So we start with the signature image here instead of the 3330-43-1
        if os.path.isfile(signature_img_path):

            # Declared early so we can close the streams later if necessary
            watermark_43_reader = None
            signed_3330_43_1_reader = None

            watermark_cover_letter_reader = PdfFileReader(open(WATERMARK_FILE_COVER_LETTER, "rb"))
            page_cover_letter = input_file_reader.getPage(0)
            page_cover_letter.mergePage(watermark_cover_letter_reader.getPage(0))
            output_file_writer.addPage(page_cover_letter)

            watermark_42_reader = PdfFileReader(open(WATERMARK_FILE_42, "rb"))
            page_42 = input_file_reader.getPage(1)
            page_42.mergePage(watermark_42_reader.getPage(0))
            output_file_writer.addPage(page_42)

            # Page 2 and 3 are not signable pages, so we add them here, as is
            output_file_writer.addPage(input_file_reader.getPage(2))
            output_file_writer.addPage(input_file_reader.getPage(3))

            # Prefer a signed 3330-43-1 page 2, fall back to signature image since we know it exists
            if os.path.isfile(signed_3330_43_1_path):
                signed_3330_43_1_reader = PdfFileReader(open(signed_3330_43_1_path, "rb"))
                output_file_writer.addPage(signed_3330_43_1_reader.getPage(0))
            else:
                watermark_43_reader = PdfFileReader(open(WATERMARK_FILE_43, "rb"))
                page_43 = input_file_reader.getPage(4)
                page_43.mergePage(watermark_43_reader.getPage(0))
                output_file_writer.addPage(page_43)

            with open(temp_path, "wb") as output_stream:
                output_file_writer.write(output_stream)

            input_file_reader.stream.close()
            watermark_cover_letter_reader.stream.close()
            watermark_42_reader.stream.close()
            if signed_3330_43_1_reader is not None:
                signed_3330_43_1_reader.stream.close()
            if watermark_43_reader is not None:
                watermark_43_reader.stream.close()

        # If no signature image is present, we may still have a 3330-43-1 page 2
        elif os.path.isfile(signed_3330_43_1_path):
            for i in range(4):
                output_file_writer.addPage(input_file_reader.getPage(i))

            signed_3330_43_1_reader = PdfFileReader(open(signed_3330_43_1_path, "rb"))
            output_file_writer.addPage(signed_3330_43_1_reader.getPage(0))

            with open(temp_path, "wb") as output_stream:
                output_file_writer.write(output_stream)

            input_file_reader.stream.close()
            signed_3330_43_1_reader.stream.close()

        else:
            # We stick with temp_path var for simplicity's sake. If we don't have any signatures, we will just use the
            # original filled ERR package for the next steps
            temp_path = form_path
            input_file_reader.stream.close()

        file_paths_to_concatenate = [temp_path]
        if os.path.isfile(resume_path):
            file_paths_to_concatenate.append(resume_path)
        if os.path.isfile(performance_path):
            file_paths_to_concatenate.append(performance_path)

        if file_paths_to_concatenate.__len__() > 1:
            # This function overwrites any existing file
            concatenate_pdfrw(file_paths_to_concatenate, output_path)
        else:
            # os.rename does not overwrite existing files, so we remove the output file if already present
            if os.path.isfile(output_path):
                os.remove(output_path)
            os.rename(temp_path, output_path)

        # Checking for leftovers
        if os.path.isfile(form_path):
            os.remove(form_path)
        if os.path.isfile(temp_path):
            os.remove(temp_path)


# Removes watermark files prior to exiting the program since they are of no use to the user
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

    # This was originally a module level function but now part of ERRWorker class
    # This was moved to allow the function to emit a pyqtSignal(str) that can be received by MainWindow.print_status
    # slot and output to a console in the UI.
    def generate_err(self, resume_path, performance_path, signed_43_1_path, signature_img_path):

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

            # dict to track completed ERRs to allow for multiple ERRs for different roles to the same facility
            err_facilities_completed = defaultdict(list)

            # If a USAJOBS application is being generated,
            # we don't need to warn the user they haven't selected an ERR facility
            if num_err == 0 and not data_backend[1].get("USAJOBS"):
                sys.stderr.write(
                    "ERROR: No desired facilities found. Please verify you have filled out 1. Personal Information.xlsx\n")
                self.status.emit(
                    "ERROR: No desired facilities found. Please verify you have filled out 1. Personal Information.xlsx")

            for i in range(1, num_err + 1):
                if data_backend[1].get("Facility" + str(i)):

                    data = data_xls.parse("PDFKeys" + str(i), header=None, index_col=0).fillna('').to_dict()

                    current_facility_id = str(data[1].get("Facility"))
                    filled_pdf_path = OUTPUT_DIRECTORY + "\\" + current_facility_id + "(temp).pdf"
                    final_output_path = OUTPUT_DIRECTORY + "\\" + current_facility_id + ".pdf"

                    # Rename duplicate ERRs to follow [location] - [role].pdf structure (eg. "DFW - CPC.pdf")
                    # This structure allows us to generate multiple ERRs to one facility for different roles
                    # While maintaining the ability to overwrite old ERRs and showing the user which ERR is which
                    if current_facility_id in err_facilities_completed:
                        if os.path.isfile(final_output_path):
                            renamed_file = final_output_path[0:(len(OUTPUT_DIRECTORY) + 4)] + " - " +\
                                      str(err_facilities_completed[current_facility_id][0]) + ".pdf"
                            if os.path.isfile(renamed_file):
                                os.remove(renamed_file)
                            os.rename(final_output_path, renamed_file)

                        final_output_path = OUTPUT_DIRECTORY + "\\" + current_facility_id + " - " + \
                            str(data[1].get("DesiredRole")).upper() + ".pdf"

                    single_form_fill(EMPTY_COVER_42_43_PDF_PATH, data[1], filled_pdf_path)
                    sign_and_append_documents(filled_pdf_path, final_output_path, signature_img_path,
                                              signed_43_1_path, resume_path, performance_path)

                    err_facilities_completed[current_facility_id].append(str(data[1].get("DesiredRole")).upper())
                    self.status.emit("Processed: " + current_facility_id)
                    self.progress.emit(int(100 * (i / num_err)))
                    print("Processed: " + current_facility_id)

            if data_backend[1].get("USAJOBS"):
                data = data_xls.parse("PDFKeysUSAJOBS", header=None, index_col=0).fillna('').to_dict()

                filled_pdf_path = "Filled USAJOBS 3330-43-1\\" + str(data[1].get("Facility")) + "(temp).pdf"
                final_output_path = "Filled USAJOBS 3330-43-1\\" + str(data[1].get("Facility")) + ".pdf"
                if not os.path.isdir("Filled USAJOBS 3330-43-1"):
                    os.mkdir("Filled USAJOBS 3330-43-1")

                single_form_fill(EMPTY_43_PDF_PATH, data[1], filled_pdf_path)
                sign_and_append_documents(filled_pdf_path, final_output_path, signature_img_path, signed_43_1_path,
                                          resume_path, performance_path, True)

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

        _clean_files()

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
