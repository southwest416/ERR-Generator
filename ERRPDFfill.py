import os

import pdfrw
import pandas as pd

# PDFRW CONSTANTS
ANNOT_KEY = '/Annots'
ANNOT_FIELD_KEY = '/T'
ANNOT_VAL_KEY = '/V'
ANNOT_RECT_KEY = '/Rect'
SUBTYPE_KEY = '/Subtype'
WIDGET_SUBTYPE_KEY = '/Widget'


# def fill_pdf(input_path, output_path, data_dict):
#     template_pdf = pdfrw.PdfReader(input_path)
#     for page in template_pdf.pages:
#         annotations = page[ANNOT_KEY]
#         if annotations is not None:
#             for annotation in annotations:
#                 if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
#                     if not annotation[ANNOT_FIELD_KEY]:
#                         annotation=annotation['/Parent']
#                     if annotation[ANNOT_FIELD_KEY]:
#                         key = annotation[ANNOT_FIELD_KEY].to_unicode()
#                         if key in data_dict.keys():
#                             if type(data_dict[key]) == bool:
#                                 if data_dict[key] == True:
#                                     annotation.update(pdfrw.PdfDict(V=pdfrw.objects.pdfname.BasePdfName('/Yes')))
#                                     annotation.update(pdfrw.PdfDict(AS=pdfrw.PdfName('Yes')))
#                             else:
#                                 print(pdfrw.PdfString.encode(str(data_dict[key])))
#                                 annotation.update(
#                                     pdfrw.PdfDict(V='{}'.format(data_dict[key]))
#                                     #pdfrw.PdfDict(V=pdfrw.PdfString.encode(str(data_dict[key])))
#                                 )
#                                 annotation.update(pdfrw.PdfDict(AP=''))
#     template_pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
#     pdfrw.PdfWriter().write(output_path, template_pdf)


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


# PDFRW Concat Function
# CREDIT: https://www.blog.pythonlibrary.org/2018/06/06/creating-and-manipulating-pdfs-with-pdfrw/
def concatenate(paths, output):
    writer = pdfrw.PdfWriter()

    for path in paths:
        reader = pdfrw.PdfReader(path)
        writer.addpages(reader.pages)

    writer.write(output)


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


if __name__ == "__main__":
    # CONSTANTS
    pdf_template = 'CoverLetter+3330-42+3330-43combined.pdf'
    data_spreadsheet_path = '1. Personal Information.xlsx'
    resume_path = 'resume.pdf'
    performance_path = 'performance.pdf'
    output_directory = 'Filled ERRs'

    # GENERATES EMPTY DICTS FOR EACH FACILITY
    data_fac1, data_fac2, data_fac3, data_fac4, data_fac5 = ({},) * 5

    # CREATES EXCELFILE OBJECT TO PARSE BACKEND DATA INTO USABLE DICT
    data_xls = pd.ExcelFile(data_spreadsheet_path, engine="openpyxl")
    data_backend = data_xls.parse("Backend", header=None, index_col=0, usecols="A,B").fillna('').to_dict()

    # CREATES OUTPUT DIRECTORY IF NOT EXISTS
    if not os.path.exists(output_directory):
        os.mkdir(output_directory)

    # ITERATES THROUGH EACH ERR
    for i in range(1, 6):
        # CHECKS IF FACILITY WAS SELECTED, SKIPS IF NOT
        if data_backend[1].get("Facility" + str(i)):

            # CREATES DICT FOR FACILITY i CONTAINING ALL PDF KEYS TO FILL
            data = data_xls.parse("PDFKeys" + str(i), header=None, index_col=0).fillna('').to_dict()

            # OUTPUT PATH CONSTANTS
            output_path = output_directory + "\\" + str(data[1].get("Facility")) + "(temp).pdf"
            final_output_path = output_directory + "\\" + str(data[1].get("Facility")) + ".pdf"

            # RUNS SINGLE FORM FILL WITH PDF KEYS & GENERATES TEMPORARY OUTPUT FILE
            single_form_fill(pdf_template, data[1], output_path)

            # CREATES LIST OF FILES TO CONCATENATE INTO ONE PACKAGE
            # IF FILES NOT FOUND, SKIPS
            concat_paths = [output_path]
            if os.path.isfile(resume_path):
                concat_paths.append(resume_path)
            if os.path.isfile(performance_path):
                concat_paths.append(performance_path)

            # CHECKS IF PERFORMANCE PLAN OR RESUME IS ATTACHED
            # IF SO, CONCATENATES INTO ONE PACKAGE
            # IF FILES NOT FOUND, RENAMES TEMPORARY OUTPUT FILE TO FINAL OUTPUT FILE
            if concat_paths.__len__() > 1:
                concatenate_pdfrw(concat_paths, final_output_path)
                os.remove(output_path)
            else:
                if os.path.isfile(final_output_path):
                    os.remove(final_output_path)
                os.rename(output_path, final_output_path)

            print("Processed: " + str(data[1].get("Facility")))

    input("Press enter to exit!")
