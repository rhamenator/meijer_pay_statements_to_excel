import os
import re
from tkinter import messagebox
import fitz # PyMuPDF
import tqdm as tq
from tqdm import tqdm
from export_data import save_to_csv, save_to_excel, save_to_json, save_to_xml
from file_utils import pick_input_file, pick_output_file
from main_utils import is_gui_mode
from main_utils import print_message

def main_logic(input_file, output_file, text_box = None):
    if input_file is None:
        input_file = pick_input_file(input_file)
    if not input_file:
        print_message("No input file selected.")
        return False
    if output_file is None:
        output_file = pick_output_file(input_file, output_file)
    if not output_file:
        print_message("No output file selected.")
        return False

    print_message(f"Analyzing...\nInput file: {input_file}\nOutput file: {output_file}")
    file_path = input_file
    output_data = extract_structured_data(file_path, text_box)
    
    print_message(f"Creating JSON file from\nInput file: {input_file}")
    json_output_path = output_file.replace('.xlsx', '.json')
    save_to_json(output_data, json_output_path)
    
    print_message(f"Creating XML file from\nInput file: {input_file}")
    xml_output_path = output_file.replace('.xlsx', '.xml')
    save_to_xml(output_data, xml_output_path)
    
    print_message(f"Creating CSV file from\nInput file: {input_file}")
    csv_output_path = output_file.replace('.xlsx', '.csv')
    save_to_csv(output_data, csv_output_path)
    
    print_message(f"Creating Excel workbook\n{output_file}")
    excel_output_path = output_file
    success = save_to_excel(output_data, excel_output_path)
    if is_gui_mode():
        messagebox.showinfo('Information', 'All finished!')
        print_message('Close window to exit.')
    # else:
    #     print('Press any key to exit...')
    #     event = keyboard.read_event()
    return success

def extract_structured_data(file_path, text_widget = None):
    document = fitz.open(file_path)
    total_pages = len(document)
    output_data = []
    infile = os.path.basename(file_path)
    page_num = 0
    for page in tqdm(document,f'Reading and analyzing file {infile}', total_pages):
        page_num += 1
        page_text = page.get_text("blocks")
        output_data.append(analyze_page_text(page_text, page_num))
    return output_data
 
def analyze_page_text(page_text, page_num):
    data = {
        "Identification": {},
        "Pay Period Info": {},
        "Earnings": {},
        "Employee Taxes": {},
        "Post Tax Deductions": {},
        "Pre Tax Deductions": {},
        "Employer Paid Benefits": {},
        "Subject or Taxable Wages": {},
        "Allowances": {},
        "Absence Plans": {},
        "Payment Information": {}
    }
    
    lines = [block[4] for block in page_text]
    
    # Process Pay Period Info
    if len(lines) > 4:
        values = lines[3].split('\n')
        identification_headers = ["Name", "Company", "Employee ID"]
        identification_values = values[0:2]
        data["Identification"].update(dict(zip(identification_headers, identification_values)))
        
        pay_period_info_headers = ["Pay Period Begin", "Pay Period End", "Check Date", "Check Number"]
        pay_period_info_values = values[3:6]
        data["Pay Period Info"].update(dict(zip(pay_period_info_headers, pay_period_info_values)))
        
        pay_period_begin = values[3]
        pay_period_end = values[4]
    
    # Process other sections
    section_names = ["Earnings", "Employee Taxes", "Pre Tax Deductions", "Post Tax Deductions", "Subject or Taxable Wages", "Employer Paid Benefits", "Absence Plans", "Marital Status", "Payment Information"]
    
    for line in lines[7:]:
        if line.strip() in section_names:
            other_sections = [s for s in section_names if s != line.strip()]
            if line.strip() == "Earnings":
                process_earnings_table(lines[lines.index(line):], data["Earnings"], page_num + 1, pay_period_begin, pay_period_end, other_sections)
            elif line.strip() in ["Employee Taxes", "Pre Tax Deductions", "Post Tax Deductions"]:
                process_deductions(lines[lines.index(line):], data[line.strip()], page_num + 1, other_sections)
            elif line.strip() == "Subject or Taxable Wages":
                process_table_subject_taxable_wages(lines[lines.index(line):], data["Subject or Taxable Wages"], page_num + 1, other_sections)
            elif line.strip() == "Employer Paid Benefits":
                process_employer_paid_benefits(lines[lines.index(line):], data["Employer Paid Benefits"], page_num + 1, other_sections)
            elif "Absence Plans" in line:
                process_absence_plans(lines[lines.index(line):], data["Absence Plans"], page_num + 1, other_sections)
            elif "Marital Status" in line:
                process_allowances(lines[lines.index(line):], data["Allowances"], page_num + 1, other_sections)
    
    return data


def process_earnings_table(lines, target_dict, page_num, pay_period_begin, pay_period_end, section_names):
    dictionary_name = ''
    ending_tags = ['Total']
    ending_tags.extend(section_names)
    for line_no, line in enumerate(lines):
        values = line.split('\n')
        if line_no in [0,1]:
            continue
        elif values[0] in ending_tags:
            break
        first_col_updated = ''
        dates = ''
        hours = ''
        rate = 0
        amount = 0
        ytd_hours = 0
        ytd_amount = 0
        if line_no == 0:
            dictionary_name = line.strip('\n')
        if len(values) >= 2:
            description = values[0]
            date_regex = r'^(.*?)\b(\d{2}/\d{2}/\d{4})\b'
            match = re.match(date_regex, description)
            if "Total" in values[0]:
                description = (((values[0]).strip()).replace(':','')).replace(dictionary_name,'')
                dates = pay_period_begin + " - " + pay_period_end
                amount = safe_float_conversion(values[1].replace(",", ""))
                ytd_amount = safe_float_conversion(values[2].replace(",", ""))
                target_dict[description] = {'Dates': dates, 'Amount': amount, 'YTD Amount': ytd_amount}
                continue
            elif match:
                dates = description[-23:]
                description_text = description.replace(dates,"")
                first_col_updated = description_text.strip()
                if len(values) < 8:
                    values.insert(1,dates)
                else:
                    values[1]=dates
                if values[2] == "0" and (values[3]).isspace() and (values[4]).isspace():
                    continue
                else: hours = safe_float_conversion(values[2])
                rate = safe_float_conversion(values[3])
                amount_text = values[4].replace(",", "")
                if amount_text.count(".") > 1:
                    amount = safe_float_conversion(amount_text[:amount_text.find(".") + 3])
                    ytd_hours = safe_float_conversion(amount_text.replace((values[4].replace(",", "")).replace(amount_text,""),""))
                    ytd_amount = safe_float_conversion(values[5].replace(",", ""))                
                else:
                    amount = safe_float_conversion(amount_text)
                    ytd_hours = safe_float_conversion(values[5].replace(",", ""))
                    ytd_amount = safe_float_conversion(values[6].replace(",", ""))
            else:
                first_col_updated = description
                if len(values) < 8 and not values[1].strip():
                    values.insert(1,pay_period_begin + " - " + pay_period_end)
                dates = (values[1]).strip()
                
                if len((values[2]).split()) > 1:
                    split_values = (values[2]).split()
                    hours = safe_float_conversion(split_values[0])
                    values.insert(3, split_values[1])
                else:
                    hours = safe_float_conversion(values[2])
                rate = safe_float_conversion(values[3])
                amount_text = values[4].replace(",", "")
                if amount_text.count(".") > 1:
                    amount = safe_float_conversion(amount_text[:amount_text.find(".") + 3])
                    ytd_hours_text = amount_text.replace((values[4].replace(",", "")).replace(amount_text,""),"")
                    if " " in ytd_hours_text:
                        ytd_hours = safe_float_conversion(ytd_hours_text[ytd_hours_text.find(" ")+1:])
                    else:
                        ytd_hours = safe_float_conversion((amount_text[amount_text.find(".") + 3:]).strip())
                    ytd_amount = safe_float_conversion(values[5].replace(",", ""))                
                else:
                    amount = safe_float_conversion(amount_text)
                    if first_col_updated == "PRC Hours Balance":
                        ytd_hours = 0
                        ytd_amount = safe_float_conversion(values[5].replace(",", ""))
                    else:
                        if len(values) < 8:
                            ytd_hours = safe_float_conversion(values[4].replace(",", ""))
                            ytd_amount = safe_float_conversion(values[5].replace(",", ""))
                        else: 
                            ytd_hours = safe_float_conversion(values[5].replace(",", ""))
                            ytd_amount = safe_float_conversion(values[6].replace(",", ""))

            if first_col_updated == "Night Premium" and rate:
                first_col_updated = "Night Premium"

            if first_col_updated in target_dict:
                existing_data = target_dict[first_col_updated]
                existing_data['Hours'] += hours
                existing_data['Amount'] += amount
                existing_data['YTD Hours'] += ytd_hours
                existing_data['YTD Amount'] += ytd_amount
            else:
                target_dict[first_col_updated] = {'Dates': dates, 'Hours': hours, 'Rate': rate, 'Amount': amount, 'YTD Hours': ytd_hours, 'YTD Amount': ytd_amount}

            if not dates.strip():
                target_dict[first_col_updated]['Amount'] = 0
        
        else: break
        

        
def process_deductions(lines, target_dict, page_num, other_sections):
    headers = ''
    dictionary_name = ''
    ending_tags = ['Total']
    ending_tags.extend(other_sections)
    for line_no, line in enumerate(lines):
        if line_no == 0:
            dictionary_name = line.strip('\n')
        elif line_no == 1:
            headers = line.split('\n')
            continue
        else:
            values = line.split('\n')
            description = values[0].strip()
            if description.endswith('Total:'):
                description = 'Total'
            if len(values) > 2:
                values[1] = safe_float_conversion(values[1].replace(",", ""))
                values[2] = safe_float_conversion(values[2].replace(",", ""))
                section_data = dict(zip(headers[1:-1], values[1:-1]))
                target_dict[description] = {k: v for k, v in section_data.items()}
            if description == 'Total' or description in ending_tags: break


def process_table_subject_taxable_wages(lines, target_dict, page_num, other_sections):
    headers = ''
    dictionary_name = ''
    ending_tags = ['Total', 'Federal']
    ending_tags.extend(other_sections)
    for line_no, line in enumerate(lines):
        values = line.split('\n')
        if ((values[0]).replace('\n','')).strip() in ending_tags:
            break
        elif line_no == 0:
            dictionary_name = values[0]
            continue
        elif line_no == 1:
            headers = values
            continue
        else:
            for item in range(1, len(values[1:])):
                values[item] = safe_float_conversion(values[item].replace(',',''))
            section_data = dict(zip(headers[1:-1], values[1:-1]))
            description = values[0]
            target_dict[description] = {k: v for k, v in section_data.items()}

def process_employer_paid_benefits(lines, target_dict, page_num, other_sections):
    headers = ''
    dictionary_name = ''
    ending_tags = ['Total']
    ending_tags.extend(other_sections)
    for line_no, line in enumerate(lines):
        if line_no == 0:
            dictionary_name = line.strip('\n')
        elif line_no == 1:
            headers = line.split('\n')
            continue
        else:
            values = line.split('\n')
            description = values[0]
            if len(values) > 2:
                values[1] = safe_float_conversion(values[1].replace(",", ""))
                values[2] = safe_float_conversion(values[2].replace(",", ""))
                section_data = dict(zip(headers[1:-1], values[1:-1]))
                if 'Total' in description: description = 'Total'
                target_dict[description] = {k: v for k, v in section_data.items()}
            if description in ending_tags: break


def process_absence_plans(lines, target_dict, page_num, other_sections):
    dictionary_name = ''
    ending_tags = ['Total', 'Federal']
    ending_tags.extend(other_sections)
    for line_no, line in enumerate(lines):
        this_line = line.split('\n')
        if line_no == 0:
            dictionary_name = this_line[0]
            continue
        elif line_no == 1:
            headers = this_line
            continue
        elif this_line[0] in ending_tags:
            break
        else:
            values = line.split('\n')
            for item in range(1, len(values[1:])):
                values[item] = safe_float_conversion(values[item].replace(',',''))
            section_data = dict(zip(headers[1:-1], values[1:-1]))
            description = headers[0]
            target_dict[description] = {k: v for k, v in section_data.items()}

def process_allowances(lines, target_dict, page_num, other_sections):
    ending_tags = ['Total', 'Payment Information']
    ending_tags.extend(other_sections)
    for line_no, line in enumerate(lines):
        values = line.split('\n')
        if line_no == 0:
            headers = ['Description', 'Federal', 'State']
        description = (values[0]).strip()
        if description == ending_tags[1]:
            break
        for item in range(1, len(values[1:])):
            values[item] = safe_int_conversion(values[item].replace(',',''))
        section_data = dict(zip(headers[1:], values[1:]))
        target_dict[description] = {k: v for k, v in section_data.items()}
        if description in ending_tags: break

def process_payment_information(lines, target_dict, page_num, other_sections):
    ending_tags = ['Total', 'Federal']
    ending_tags.extend(other_sections)
    headers = []
    values=[]
    description = ''
    headers = ''
    dictionary_name = ''
    for line_no, line in enumerate(lines):
        this_line = line.split('\n')
        if line_no == 0:
            dictionary_name = this_line[0]
            continue
        elif line_no == 1:
            headers = this_line
        else:
            values = line.split('\n')
            description = values[0]
            if description in ending_tags: break
            else:
                headers[0] = 'Name'
                if (values[-2]).endswith('USD'):
                    amountfield = (values[-2]).split()
                    values[-2] = safe_float_conversion(amountfield[0])
                    values[-1] = safe_float_conversion(amountfield[0])
                    values.append(amountfield[1])
                    headers[-1] = 'Currency'
                else:
                    amountfield = (values[-1]).split()
                    values[-1] = amountfield[0]
                    values.append(amountfield[1])
                    headers[-1] = 'Currency'
                section_data = dict(zip(headers, values))
                target_dict['Bank'] = {k: v for k, v in section_data.items()}
        if description in ending_tags: break

def safe_float_conversion(value):
    try:
        return float(value.replace(',', ''))
    except ValueError:
        return 0.0
    
def safe_int_conversion(value):
    try:
        if type(value) == int: 
            return value
        if type(value) == float:
            return int(value)
        if type(value) == str:
            result = 0 if value.isspace() else int(value)
            return result
        if type(value) == bool:
            result = 1 if value == True else 0
            return result
        return int(value)
    except ValueError:
        return 0
