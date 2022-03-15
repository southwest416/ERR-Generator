import os

import pdfrw
import pandas as pd
import reportlab.lib.pagesizes
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject

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


# PDFRW FUNCTIONS
# CREDIT: https://github.com/WestHealth/pdf-form-filler
def _text_form(annotation, value):
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


# Creates signature watermark PDFs to overlay before generating & signing packages
# CREDIT: https://stackoverflow.com/questions/2925484/place-image-over-pdf
def init_signatures(signature_path, sup_signature_path):
    # CONSTANTS & INITIALIZATIONS
    WATERMARK_FILE_COVER_LETTER = RESOURCE_DIRECTORY + '\\coverwatermark.pdf'
    WATERMARK_FILE_42 = RESOURCE_DIRECTORY + '\\42watermark.pdf'
    WATERMARK_FILE_43 = RESOURCE_DIRECTORY + '\\43watermark.pdf'

    # IF A SIGNATURE IS ATTACHED, CREATE WATERMARK FILES FOR EACH PAGE THAT NEEDS TO BE SIGNED
    # THESE FILES WILL LATER BE OVERLAID ON THE FILLED ERR DOCUMENTS
    if signature_path != '':
        # CREATE SIGNATURE WATERMARK FOR COVER LETTER PAGE
        canvas_cover = canvas.Canvas(WATERMARK_FILE_COVER_LETTER, pagesize=reportlab.lib.pagesizes.letter)
        canvas_cover.drawImage(signature_path, 0, 330, height=36, preserveAspectRatio=True)
        canvas_cover.save()
        # CREATE SIGNATURE WATERMARK FOR 3330-42
        canvas_42 = canvas.Canvas(WATERMARK_FILE_42, pagesize=reportlab.lib.pagesizes.letter)
        canvas_42.drawImage(signature_path, 8, 489, height=24, preserveAspectRatio=True)
        canvas_42.save()
        # CREATE SIGNATURE WATERMARK FOR 3330-43-1
        canvas_43 = canvas.Canvas(WATERMARK_FILE_43, pagesize=reportlab.lib.pagesizes.letter)
        canvas_43.drawImage(signature_path, -24, 92, height=24, preserveAspectRatio=True)
        if sup_signature_path != '':
            canvas_43.drawImage(sup_signature_path, 264, 92, height=24, preserveAspectRatio=True)
        canvas_43.save()

    # IF ONLY A SUPERVISOR SIGNATURE IS ATTACHED, CREATE A WATERMARK FILE ONLY FOR THE 3330-43-1
    if sup_signature_path != '':
        if not os.path.isfile(WATERMARK_FILE_43):
            canvas_43 = canvas.Canvas(WATERMARK_FILE_43, pagesize=reportlab.lib.pagesizes.letter)
            canvas_43.drawImage(sup_signature_path, 264, 92, height=24, preserveAspectRatio=True)
            canvas_43.save()


# Overlays signature watermarks onto filled packages
# CREDIT: https://stackoverflow.com/questions/2925484/place-image-over-pdf
def draw_signatures(form_path, output_path):
    # CONSTANTS & INITIALIZATIONS
    WATERMARK_FILE_COVER_LETTER = RESOURCE_DIRECTORY + '\\coverwatermark.pdf'
    WATERMARK_FILE_42 = RESOURCE_DIRECTORY + '\\42watermark.pdf'
    WATERMARK_FILE_43 = RESOURCE_DIRECTORY + '\\43watermark.pdf'

    # IF A SIGNATURE WATERMARK FILE IS FOUND, OVERLAY WATERMARK FILES FOR ALL 3 SIGNABLE PAGES ONTO PACKAGE
    if os.path.isfile(WATERMARK_FILE_COVER_LETTER):
        # CREATE INPUT READER & OUTPUT WRITER
        output_file = PdfFileWriter()
        pypdf_set_need_appearances_writer(output_file)
        input_file = PdfFileReader(open(form_path, "rb"))

        # GET COVER LETTER PAGE FROM FILLED PACKAGE, OVERLAY WATERMARK FILE, ADD TO OUTPUT FILE
        cover_page = input_file.getPage(0)
        watermark_cover = PdfFileReader(open(WATERMARK_FILE_COVER_LETTER, "rb"))
        cover_page.mergePage(watermark_cover.getPage(0))
        output_file.addPage(cover_page)

        # GET 3330-42 PAGE FROM FILLED PACKAGE, OVERLAY WATERMARK FILE, ADD TO OUTPUT FILE
        page_42 = input_file.getPage(1)
        watermark_42 = PdfFileReader(open(WATERMARK_FILE_42, "rb"))
        page_42.mergePage(watermark_42.getPage(0))
        output_file.addPage(page_42)

        # ADD FILLED 3330-42 PAGE 2 & 3330-43-1 PAGE 1 TO OUTPUT FILE
        output_file.addPage(input_file.getPage(2))
        output_file.addPage(input_file.getPage(3))

        # GET 3330-43-1 PAGE FROM FILLED PACKAGE, OVERLAY WATERMARK FILE, ADD TO OUTPUT FILE
        page_43 = input_file.getPage(4)
        watermark_43 = PdfFileReader(open(WATERMARK_FILE_43, "rb"))
        page_43.mergePage(watermark_43.getPage(0))
        output_file.addPage(page_43)

        # WRITE ALL PAGES TO OUTPUT_FILE
        with open(output_path, "wb") as outputStream:
            output_file.write(outputStream)

        # CLOSE FILE STREAMS
        input_file.stream.close()
        watermark_cover.stream.close()
        watermark_42.stream.close()
        watermark_43.stream.close()

    # IF ONLY A WATERMARK FILE FOR 3330-43-1 IS FOUND, OVERLAY ONLY WATERMARK FILE FOR 3330-43-1 ON PACKAGE
    elif os.path.isfile(WATERMARK_FILE_43):
        # CREATE INPUT READER & OUTPUT WRITER
        output_file = PdfFileWriter()
        pypdf_set_need_appearances_writer(output_file)
        input_file = PdfFileReader(open(form_path, "rb"))

        for i in range(4):
            output_file.addPage(input_file.getPage(i))

        # GET 3330-43-1 PAGE FROM FILLED PACKAGE, OVERLAY WATERMARK FILE, ADD TO OUTPUT FILE
        page_43 = input_file.getPage(4)
        watermark_43 = PdfFileReader(open(WATERMARK_FILE_43, "rb"))
        page_43.mergePage(watermark_43.getPage(0))
        output_file.addPage(page_43)

        # WRITE ALL PAGES TO OUTPUT_FILE
        with open(output_path, "wb") as outputStream:
            output_file.write(outputStream)

        # CLOSE FILE STREAMS
        input_file.stream.close()
        watermark_43.stream.close()


def clean_files():
    if os.path.isfile(RESOURCE_DIRECTORY + '\\42watermark.pdf'):
        os.remove(RESOURCE_DIRECTORY + '\\42watermark.pdf')
    if os.path.isfile(RESOURCE_DIRECTORY + '\\43watermark.pdf'):
        os.remove(RESOURCE_DIRECTORY + '\\43watermark.pdf')
    if os.path.isfile(RESOURCE_DIRECTORY + '\\coverwatermark.pdf'):
        os.remove(RESOURCE_DIRECTORY + '\\coverwatermark.pdf')


if __name__ == "__main__":
    # CONSTANTS
    EMPTY_PDF_PATH = RESOURCE_DIRECTORY + '\\CoverLetter+3330-42+3330-43combined.pdf'
    DATA_SPREADSHEET_PATH = '1. Personal Information.xlsx'
    RESUME_PATH = 'resume.pdf'
    PERFORMANCE_PATH = 'performance.pdf'
    SIGNATURE_IMAGE_PATH = get_signature()
    SUP_SIGNATURE_IMAGE_PATH = get_sup_signature()
    init_signatures(SIGNATURE_IMAGE_PATH, SUP_SIGNATURE_IMAGE_PATH)

    # CREATES EXCELFILE OBJECT TO PARSE BACKEND DATA INTO USABLE DICT
    data_xls = pd.ExcelFile(DATA_SPREADSHEET_PATH, engine="openpyxl")
    data_backend = data_xls.parse("Backend", header=None, index_col=0, usecols="A,B").fillna('').to_dict()

    # CREATES OUTPUT, BUILD, RESOURCE DIRECTORIES IF NOT EXISTS
    if not os.path.exists(OUTPUT_DIRECTORY):
        os.mkdir(OUTPUT_DIRECTORY)
    if not os.path.exists(RESOURCE_DIRECTORY):
        os.mkdir(RESOURCE_DIRECTORY)

    # ITERATES THROUGH EACH ERR
    for i in range(1, 21):
        # CHECKS IF FACILITY WAS SELECTED, SKIPS IF NOT
        if data_backend[1].get("Facility" + str(i)):

            # CREATES DICT FOR FACILITY i CONTAINING ALL PDF KEYS TO FILL
            data = data_xls.parse("PDFKeys" + str(i), header=None, index_col=0).fillna('').to_dict()

            # OUTPUT PATH CONSTANTS
            FILLED_PDF_PATH = OUTPUT_DIRECTORY + "\\" + str(data[1].get("Facility")) + "(temp).pdf"
            SIGNED_PDF_PATH = OUTPUT_DIRECTORY + "\\" + str(data[1].get("Facility")) + "(signed).pdf"
            FINAL_OUTPUT_PATH = OUTPUT_DIRECTORY + "\\" + str(data[1].get("Facility")) + ".pdf"

            # FILL FORM AND DRAW SIGNATURES IF PRESENT
            single_form_fill(EMPTY_PDF_PATH, data[1], FILLED_PDF_PATH)
            draw_signatures(FILLED_PDF_PATH, SIGNED_PDF_PATH)

            # CREATES LIST OF FILES TO CONCATENATE INTO ONE PACKAGE
            if os.path.isfile(SIGNED_PDF_PATH):
                concat_paths = [SIGNED_PDF_PATH]
            else:
                concat_paths = [FILLED_PDF_PATH]

            # CHECK IF RESUME OR PMAS IS ATTACHED, APPEND TO END
            if os.path.isfile(RESUME_PATH):
                concat_paths.append(RESUME_PATH)
            if os.path.isfile(PERFORMANCE_PATH):
                concat_paths.append(PERFORMANCE_PATH)

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

            print("Processed: " + str(data[1].get("Facility")))

    clean_files()
    input("Press enter to exit!")
