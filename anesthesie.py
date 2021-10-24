#!/usr/bin/env python3
import openpyxl
import os
import pandas as pd

home = os.path.expanduser("~")
work_folder = os.path.join(home, 'OneDrive/Business/smartvets/Specialist Bouvien/financieel-dev')
product_file = 'Materialen.xlsx'
anesthesia_table = 'Anesthesie Tabel.xlsx'
export_file = 'ProductList.xlsx'
os.chdir(work_folder)
print(os.getcwd())

df_anesthesia = pd.read_excel(anesthesia_table, 'table')

def get_ml(weight, product):
    # skip rows with mg/ml and prices
    value = 0
    lookup_table = df_anesthesia.drop([0,1])
    for index, row in lookup_table.iterrows():
        if row['weight'] <= weight:
            value = row[product]
    return(value)

def get_price(product):
    price_table = df_anesthesia.iloc[1]
    return(price_table[product])


workbook = openpyxl.load_workbook('Patientadministratie.xlsx')
worksheet = workbook['anestesie']

col_customer_code   =  0
col_weight          =  1
col_dexmedetomidine =  2
col_midazolam       =  3
col_ketamine        =  4
col_methadon        =  5
col_buprenorfine    =  6
col_carprofen       =  7
col_meloxicam       =  8
col_propofol        =  9
col_sedation        = 10
col_analgesia       = 11
col_nsaid           = 12
col_induction       = 13

price_injection     = 13.40

for row in worksheet.iter_rows():
    customer_code        = row[col_customer_code].value
    weight               = row[col_weight].value or 0
    inj_dexmedetomidine  = row[col_dexmedetomidine].value or 0
    inj_midazolam        = row[col_midazolam].value or 0
    inj_ketamine         = row[col_ketamine].value or 0
    inj_methadon         = row[col_methadon].value or 0
    inj_buprenorfine     = row[col_buprenorfine].value or 0
    inj_carprofen        = row[col_carprofen].value or 0
    inj_meloxicam        = row[col_meloxicam].value or 0
    inj_propofol         = row[col_propofol].value or 0
    sedation             = row[col_sedation].value
    analgesia            = row[col_analgesia].value
    nsaid                = row[col_nsaid].value
    induction            = row[col_induction].value

    if customer_code is not None and sedation is None:
        sedation = inj_dexmedetomidine * get_price('dexmedetomidine') * get_ml(weight, 'dexmedetomidine') + \
                   inj_midazolam * get_price('midazolam') * get_ml(weight, 'midazolam') + \
                   inj_ketamine * get_price('ketamine') * get_ml(weight, 'ketamine') + \
                   max(inj_dexmedetomidine, inj_midazolam, inj_ketamine) * price_injection
        print(sedation)
        row[col_sedation].value = round(sedation, 2)

    if customer_code is not None and analgesia is None:
        analgesia = inj_buprenorfine * get_price('buprenorfine') * get_ml(weight, 'buprenorfine') + \
                    inj_buprenorfine * price_injection                   

        print(analgesia)
        row[col_analgesia].value = round(analgesia, 2)
        
    if customer_code is not None and nsaid is None:
        nsaid = inj_carprofen * get_price('carprofen') * get_ml(weight, 'carprofen') + \
                inj_meloxicam * get_price('meloxicam') * get_ml(weight, 'meloxicam') + \
                max(inj_carprofen, inj_meloxicam) * price_injection
        print(nsaid)
        row[col_nsaid].value = round(nsaid, 2)

    if customer_code is not None and induction is None:
        induction = inj_propofol * get_price('propofol') * get_ml(weight, 'propofol') + \
                    inj_propofol * price_injection

        print(induction)
        row[col_induction].value = round(induction, 2)

    workbook.save('Patientadministratie.xlsx')


