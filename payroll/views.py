from django.shortcuts import redirect, render
from django.views.decorators.csrf import csrf_exempt

from docx import Document
from docx.shared import Pt
import os
import subprocess
from django.http import HttpResponse
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL

import datetime
import io
import zipfile

import math
from PyPDF2 import PdfReader, PdfWriter

# Create your views here.

@csrf_exempt
def index(request):
    if request.method == 'POST':
        gross_salary, fedWithholding, ss, medicare, fica_deduction = payroll_calculator(int(request.POST['anual']), int(request.POST['period']))
        # return generate_pdf(request, gross_salary, fedWithholding, ss, medicare, fica_deduction)

        return generate_2do_pdf(request, gross_salary, fedWithholding, ss, medicare, fica_deduction)
    return render(request, 'index.html')

def payroll_calculator(anual, period):
    gross_salary = round_up(anual / period)
    fedWithholding  = round_up(gross_salary * get_tax_rate(anual))
    ss = round_up(gross_salary * 0.062)
    medicare = round_up(gross_salary * 0.0145)
    fica_deduction = round_up(fedWithholding + ss + medicare)
    return gross_salary, fedWithholding, ss, medicare, fica_deduction

def generate_pdf(request, gross_salary, fedWithholding, ss, medicare, fica_deduction):
    start_date = datetime.datetime(2023, 12, 21)
    start_period = get_pay_date_correct(datetime.datetime.strptime(request.POST['start_period'], '%Y-%m-%d'))
    end_period = get_pay_date_correct(datetime.datetime.strptime(request.POST['end_period'], '%Y-%m-%d'))
    number_payments = (end_period - start_period).days // 14
    period = int(request.POST['period'])
    check_id = request.POST['check_id']

    temp_docx_paths = []
    temp_pdf_paths = []
    final_pdf_paths = []

    try:
        for i in range(number_payments):
            start_period += datetime.timedelta(days=14)
            if period == 26:
                payment_number = (start_period - start_date).days // 14
            else:
                payment_number = (start_period - start_date).days // 7

            temp_docx_path = f'temp_modified_{i}.docx'
            temp_pdf_path = f'temp_output_{i}.pdf'
            
            # Cargar el archivo .docx base
            doc = Document('base.docx')
            
            # Modificar el archivo .docx (por ejemplo, agregar un nombre)
            replacements = {
                '<<nombre>>': f"{request.POST['name']} {request.POST['last_name']}",
                '<<client_address>>': request.POST['client_address'],
                '<<company>>': request.POST['company'],
                '<<city_state>>': request.POST['city_state'],
                '<<address_co>>': request.POST['address_co'],
                '<<check_id>>': str(check_id),
                '<<fecha>>': start_period.strftime('%m/%d/%Y'),
                '<<pay_date>>': (start_period - datetime.timedelta(days=14)).strftime('%m/%d/%Y'),
                '<<netpaytext>>': number_to_words(round(gross_salary - fica_deduction)),
                '<<decimal>>': get_decimal_part(gross_salary - fica_deduction),
                '<<ssn_digits>>': request.POST['ssn_digits'],
                '<<netpay>>': format_number(round_up(gross_salary - fica_deduction)),
                '<<dependents>>': request.POST['dependents'],
                '<<salary>>': str(format_number(gross_salary)),
                '<<fed>>': str(format_number(fedWithholding)),
                '<<ss>>': str(format_number(ss)),
                '<<mc>>': str(format_number(medicare)),
                '<<totalt>>': str(format_number(round_up(fedWithholding + ss + medicare))),
                # Years to Date
                '<<salaryytd>>': str(format_number(round_up(gross_salary * payment_number))),
                '<<fedytd>>': str(format_number(round_up(fedWithholding * payment_number))),
                '<<ssytd>>': str(format_number(round_up(ss * payment_number))),
                '<<mcytd>>': str(format_number(round_up(medicare * payment_number))),
                '<<totaltytd>>': str(format_number(round_up((fedWithholding + ss + medicare) * payment_number))),
            }

            def change_font_size(run, size):
                run.font.size = Pt(size)

            def justify_paragraph(paragraph, alignment):
                paragraph.alignment = alignment

            no_font_size_changes = ['<<nombre>>', '<<fecha>>', '<<netpay>>', '<<dependents>>', '<<pay_date>>']

            # Modificar los párrafos
            for paragraph in doc.paragraphs:
                for key, value in replacements.items():
                    if key in paragraph.text:
                        paragraph.text = paragraph.text.replace(key, value)

            # Modificar las tablas
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for key, value in replacements.items():
                            if key in cell.text:
                                cell.text = cell.text.replace(key, value)
                                for paragraph in cell.paragraphs:
                                    for run in paragraph.runs:
                                        if key not in no_font_size_changes:
                                            change_font_size(run, 7)
                                            justify_paragraph(paragraph, WD_PARAGRAPH_ALIGNMENT.RIGHT)
                                        elif key == '<<netpay>>':
                                            change_font_size(run, 9)
                                            run.bold = True
                                        else:
                                            change_font_size(run, 8)
                                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            # Definir el directorio de salida para los PDFs
            output_dir = os.path.join('media', 'pdfs')

            # Crear el directorio si no existe
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # Guardar el documento modificado temporalmente
            doc.save(temp_docx_path)
            temp_docx_paths.append(temp_docx_path)  # Agregar a la lista de documentos temporales

            # Convertir el documento .docx modificado a PDF usando LibreOffice
            result = subprocess.run(
                ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, temp_docx_path],
                capture_output=True,
                text=True
            )

            if result.returncode != 0:
                print(f"Error al convertir a PDF: {result.stderr}")
                return HttpResponse(f"Error durante la conversión a PDF. {result.stderr}", status=500)

            # Comprobar el nombre del archivo PDF creado
            temp_pdf_path = os.path.join(output_dir, f'temp_modified_{i}.pdf')
            if not os.path.exists(temp_pdf_path):
                print(f"Error: {temp_pdf_path} no se ha creado.")
                return HttpResponse(f"Error durante la conversión a PDF.", status=500)

            # Obtener el nombre del archivo PDF desde los parámetros de la solicitud
            pdf_name = f"{request.POST['name']}{request.POST['last_name']}_{start_period.strftime('%m%d%Y')}_{round_up(gross_salary)}{'BiWeekly' if request.POST['period'] == '26' else 'Weekly' if request.POST['period'] == '52' else ''}.pdf"

            # Renombrar el archivo PDF
            final_pdf_path = os.path.join(output_dir, pdf_name)
            os.rename(temp_pdf_path, final_pdf_path)
            final_pdf_paths.append(final_pdf_path)
            
            check_id = int(check_id) + 13

        # Crear el archivo ZIP en memoria
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            for pdf_path in final_pdf_paths:  # Cambiar pdf_files a final_pdf_paths
                zip_file.write(pdf_path, os.path.basename(pdf_path))

        # Preparar la respuesta HTTP con el archivo ZIP
        zip_buffer.seek(0)
        response = HttpResponse(zip_buffer, content_type='application/zip')
        response['Content-Disposition'] = 'attachment; filename="payroll_pdfs.zip"'
        
    finally:
        for path in temp_docx_paths + temp_pdf_paths + final_pdf_paths:
            if os.path.exists(path):
                os.remove(path)

    return response




def get_tax_rate(income):
    tax_brackets = [
        (0, 10275, 0.10),
        (10276, 41775, 0.12),
        (41776, 89075, 0.22),
        (89076, 170050, 0.24),
        (170051, 215950, 0.32), 
        (215951, 539900, 0.35),
        (539901, float('inf'), 0.37)
    ]
    
    for lower, upper, rate in tax_brackets:
        if lower <= income <= upper:
            return rate
        
def format_number(number):
    return "{:,.2f}".format(number)

def round_up(number, decimals=2):
    factor = 10 ** decimals
    return math.ceil(number * factor) / factor

def number_to_words(num):
    # Definimos palabras para números del 0 al 19 y de 20 en 20 hasta 90
    num_words = {
        0: "ZERO", 1: "ONE", 2: "TWO", 3: "THREE", 4: "FOUR", 5: "FIVE",
        6: "SIX", 7: "SEVEN", 8: "EIGHT", 9: "NINE", 10: "TEN",
        11: "ELEVEN", 12: "TWELVE", 13: "THIRTEEN", 14: "FOURTEEN", 15: "FIFTEEN",
        16: "SIXTEEN", 17: "SEVENTEEN", 18: "EIGHTEEN", 19: "NINETEEN",
        20: "TWENTY", 30: "THIRTY", 40: "FORTY", 50: "FIFTY",
        60: "SIXTY", 70: "SEVENTY", 80: "EIGHTY", 90: "NINETY"
    }

    # Definimos palabras para las centenas
    hundreds_words = {
        100: "HUNDRED", 1000: "THOUSAND"
    }

    if num < 20:
        return num_words[num]
    elif num < 100:
        tens = num // 10 * 10
        ones = num % 10
        if ones == 0:
            return num_words[tens]
        else:
            return num_words[tens] + " " + num_words[ones]
    elif num < 1000:
        hundreds = num // 100
        remainder = num % 100
        if remainder == 0:
            return num_words[hundreds] + " " + hundreds_words[100]
        else:
            return num_words[hundreds] + " " + hundreds_words[100] + " AND " + number_to_words(remainder)
    elif num < 10000:
        thousands = num // 1000
        remainder = num % 1000
        if remainder == 0:
            return num_words[thousands] + " " + hundreds_words[1000]
        else:
            return num_words[thousands] + " " + hundreds_words[1000] + " " + number_to_words(remainder)
    else:
        return "Number out of range (must be between 0 and 9999)"
    
def get_decimal_part(number):
    # Convert the number to string if it's not already
    number_str = str(number)
    # Split into integer and decimal parts
    parts = number_str.split('.')
    if len(parts) > 1:
        # Return only the first two digits of the decimal part
        return parts[1][:2].ljust(2, '0')  # Ensure it always returns at least two digits
    else:
        return "00"  # Return '00' if there is no decimal part
    
def get_pay_date_correct(pay_date):
    start_date = datetime.datetime(2023, 12, 22)
    if (pay_date - datetime.datetime(2024, 1, 5)).days % 14 == 0:
        return pay_date
    else:
        rounded = (pay_date - start_date).days // 14
        mult = rounded * 14
        date_object = start_date + datetime.timedelta(days=mult)
        return date_object


def generate_2do_pdf(request, gross_salary, fedWithholding, ss, medicare, fica_deduction):
    start_date = datetime.datetime(2023, 12, 21)
    start_period = get_pay_date_correct(datetime.datetime.strptime(request.POST['start_period'], '%Y-%m-%d'))
    end_period = get_pay_date_correct(datetime.datetime.strptime(request.POST['end_period'], '%Y-%m-%d'))
    number_payments = (end_period - start_period).days // 14
    period = int(request.POST['period'])
    check_id = request.POST['check_id']

    temp_docx_paths = []
    temp_pdf_paths = []
    final_pdf_paths = []

    try:
        # LOOP para generar múltiples PDFs
        for i in range(number_payments):
            start_period += datetime.timedelta(days=14)
            
            temp_docx_path = f'temp_modified_{i}.docx'
            temp_pdf_path = f'temp_output_{i}.pdf'
            
            # Cargar el archivo .docx base
            doc = Document('base.docx')
            
            # Modificar el archivo .docx
            replacements = {
                '<<nombre>>': f"{request.POST['name']} {request.POST['last_name']}",
                '<<client_address>>': request.POST['client_address'],
                '<<company>>': request.POST['company'],
                '<<city_state>>': request.POST['city_state'],
                '<<address_co>>': request.POST['address_co'],
                '<<check_id>>': str(check_id),
                '<<fecha>>': start_period.strftime('%m/%d/%Y'),
                '<<pay_date>>': (start_period - datetime.timedelta(days=14)).strftime('%m/%d/%Y'),
                '<<netpaytext>>': number_to_words(round(gross_salary - fica_deduction)),
                '<<decimal>>': get_decimal_part(gross_salary - fica_deduction),
                '<<ssn_digits>>': request.POST['ssn_digits'],
                '<<netpay>>': format_number(round_up(gross_salary - fica_deduction)),
                '<<dependents>>': request.POST['dependents'],
                '<<salary>>': str(format_number(gross_salary)),
                '<<fed>>': str(format_number(fedWithholding)),
                '<<ss>>': str(format_number(ss)),
                '<<mc>>': str(format_number(medicare)),
                '<<totalt>>': str(format_number(round_up(fedWithholding + ss + medicare))),
                # Years to Date
                '<<salaryytd>>': str(format_number(round_up(gross_salary * payment_number))),
                '<<fedytd>>': str(format_number(round_up(fedWithholding * payment_number))),
                '<<ssytd>>': str(format_number(round_up(ss * payment_number))),
                '<<mcytd>>': str(format_number(round_up(medicare * payment_number))),
                '<<totaltytd>>': str(format_number(round_up((fedWithholding + ss + medicare) * payment_number))),
            }

            # Modificar los párrafos
            for paragraph in doc.paragraphs:
                for key, value in replacements.items():
                    if key in paragraph.text:
                        paragraph.text = paragraph.text.replace(key, value)

            # Modificar las tablas
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for key, value in replacements.items():
                            if key in cell.text:
                                cell.text = cell.text.replace(key, value)
                                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            # Definir el directorio de salida para los PDFs
            output_dir = os.path.join('media', 'pdfs')

            # Crear el directorio si no existe
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # Guardar el documento modificado temporalmente
            doc.save(temp_docx_path)
            temp_docx_paths.append(temp_docx_path)

            # Convertir el documento .docx modificado a PDF usando LibreOffice
            result = subprocess.run(
                ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, temp_docx_path],
                capture_output=True,
                text=True
            )

            if result.returncode != 0:
                print(f"Error al convertir a PDF: {result.stderr}")
                return HttpResponse(f"Error durante la conversión a PDF. {result.stderr}", status=500)

            # Comprobar el nombre del archivo PDF creado
            temp_pdf_path = os.path.join(output_dir, f'temp_modified_{i}.pdf')
            if not os.path.exists(temp_pdf_path):
                print(f"Error: {temp_pdf_path} no se ha creado.")
                return HttpResponse(f"Error durante la conversión a PDF.", status=500)

            # Obtener el nombre del archivo PDF
            pdf_name = f"{request.POST['name']}{request.POST['last_name']}_{start_period.strftime('%m%d%Y')}.pdf"

            # Renombrar el archivo PDF
            final_pdf_path = os.path.join(output_dir, pdf_name)
            os.rename(temp_pdf_path, final_pdf_path)
            final_pdf_paths.append(final_pdf_path)

        # Crear el archivo ZIP en memoria
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            for pdf_path in final_pdf_paths:
                zip_file.write(pdf_path, os.path.basename(pdf_path))

        # Preparar la respuesta HTTP con el archivo ZIP
        zip_buffer.seek(0)
        response = HttpResponse(zip_buffer, content_type='application/zip')
        response['Content-Disposition'] = 'attachment; filename="payroll_pdfs.zip"'
        
    finally:
        for path in temp_docx_paths + temp_pdf_paths + final_pdf_paths:
            if os.path.exists(path):
                os.remove(path)

    return response