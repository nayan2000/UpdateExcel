#!/usr/bin/python3

import openpyxl

def getStringInput(input_message):
    while True:
        string = input(input_message)
        if string:
            break
    return string


def getNumericInput(input_message):
    while(True):
        number = input(input_message)
        try:
            number = int(number)
            break
        except Exception as ValueError:
            pass
    
    return number

def get_input():
    product_name = getStringInput('Please enter product name: ')
    hsn_code = getNumericInput('Please enter HSN Code of product: ')
    rate = getNumericInput('Please enter rate of product: ')
    gst_rate = getNumericInput('Please enter the GST Rate(without % sign): ')
    
    return product_name, hsn_code, rate, gst_rate

def update_excel(product_name, hsn_code, rate, gst_rate):
    file_path = r'c:\Users\personal\Desktop\BILLING SOFTWARE.xlsx'
    book = openpyxl.load_workbook(file_path)
    sheetName = 'Itemdetail'
    sheet = book[sheetName]
    sheet.append((product_name, hsn_code, rate, gst_rate))
    book.save(file_path)
    

#Take input  
product_name, hsn_code, rate, gst_rate = get_input()
#Update these values to excel
update_excel(product_name, hsn_code, rate, gst_rate)
