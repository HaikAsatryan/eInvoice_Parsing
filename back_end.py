import pandas as pd
import xml.etree.ElementTree as et


def backend(xlsx_path, xml_path):
    try:
        tree = et.parse(xml_path)

        root = tree.getroot()

        issue_date = []
        delivery_date = []
        invoice_serial = []
        invoice_number = []
        tax_code = []
        counterpary_name = []
        invoice_amount = []
        bank_account = []

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}SubmissionDate"):
            issue_date.append(elm.text)

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}Series"):
            invoice_serial.append(elm.text)

        if invoice_serial[0] == 'Բ':
            for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}DeliveryDate"):
                delivery_date.append(elm.text)
        else:
            for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}SupplyDate"):
                delivery_date.append(elm.text)

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}Number"):
            invoice_number.append(elm.text)

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}SupplierInfo"):
            elm = elm.find(".//{http://www.taxservice.am/tp3/invoice/definitions}TIN")
            tax_code.append(elm.text)

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}SupplierInfo"):
            elm = elm.find(".//{http://www.taxservice.am/tp3/invoice/definitions}Name")
            counterpary_name.append(elm.text)

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}Total"):
            elm = elm.find(".//{http://www.taxservice.am/tp3/invoice/definitions}TotalPrice")
            invoice_amount.append(elm.text)

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}SupplierInfo"):
            try:
                elm = elm.find(".//{http://www.taxservice.am/tp3/invoice/definitions}BankAccount")
                elm = elm.find(".//{http://www.taxservice.am/tp3/invoice/definitions}BankAccountNumber")
                bank_account.append(elm.text)
            except:
                bank_account.append("N/A")

        output = {
            'issue_date': issue_date,
            'invoice_serial': invoice_serial,
            'invoice_number': invoice_number,
            'delivery_date': delivery_date,
            'Հարկ վճարողի հաշվառման համարը (ՀՎՀՀ)': tax_code,
            'Իրավաբանական անձի անվանումը կամ ֆիզիկական անձի (անհատ ձեռնարկատեր) անունը, ազգանունը': counterpary_name,
            'Արժեքը': invoice_amount,
            'Հաշվեհամար': bank_account
        }

        pd.set_option('display.max_columns', None)
        df = pd.DataFrame(output)

        df['Սերիա և համար'] = df['invoice_serial'] + df['invoice_number'].astype(str)
        df['Դուրս գրման ա/թ'] = df['issue_date'].str[:10]
        df['Դուրս գրման ա/թ'] = pd.to_datetime(df['Դուրս գրման ա/թ']).dt.strftime('%d/%m/%Y')
        df['Առաքման (տեղափոխման) ա/թ'] = df['delivery_date'].str[:10]
        df['Առաքման (տեղափոխման) ա/թ'] = pd.to_datetime(df['Առաքման (տեղափոխման) ա/թ']).dt.strftime('%d/%m/%Y')

        df = df[[
            'Դուրս գրման ա/թ',
            'Առաքման (տեղափոխման) ա/թ',
            'Սերիա և համար',
            'Հարկ վճարողի հաշվառման համարը (ՀՎՀՀ)',
            'Իրավաբանական անձի անվանումը կամ ֆիզիկական անձի (անհատ ձեռնարկատեր) անունը, ազգանունը',
            'Արժեքը',
            'Հաշվեհամար'
        ]]

        writer = pd.ExcelWriter(xlsx_path)
        df.to_excel(writer, sheet_name='Exported data')
        writer.save()

    except Exception as xml_error:
        print(xml_error)
