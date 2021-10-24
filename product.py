import pandas as pd
import numpy as np
#from datetime import date, datetime
import glob
import os
import shutil
from datetime import date
import datetime

def set_columns(workbook, worksheet, df):
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width

home = os.path.expanduser("~")
work_folder = os.path.join(home, 'OneDrive/Business/smartvets/Specialist Bouvien/financieel')
product_file = 'Materialen.xlsx'
export_file = 'ProductList.xlsx'
os.chdir(work_folder)
print(os.getcwd())

def create_product_list:
    # create export dataframe for services
    print('services')
    df_services = pd.read_excel(product_file, 'services')
    df_services = df_services.rename(columns = {'productnaam':'omschrijving', 'inkoopprijs':'inkoopprijs excl btw', 'verkoopprijs excl (**)':'verkoopprijs excl btw'})
    df_packages = df_services[ df_services['verkoopprijs excl btw'].notna() ]
    df_services.loc[df_services['btw'] == 9, 'btw'] = 'LAAG'
    df_services.loc[df_services['btw'] == 21, 'btw'] = 'HOOG'
    df_services = df_services[['code', 'omschrijving', 'inkoopprijs excl btw', 'verkoopprijs excl btw', 'btw', 'groep', 'tegenrekening']]

    # create export dataframe for packages of reusables
    print('packages')
    df_packages = pd.read_excel(product_file, 'packages')
    df_packages = df_packages.rename(columns = {'productnaam':'omschrijving', 'inkoopprijs':'inkoopprijs excl btw', 'verkoopprijs per gebruik (**)':'verkoopprijs excl btw'})
    df_packages = df_packages[ df_packages['verkoopprijs excl btw'].notna() ]
    df_packages.loc[df_packages['btw'] == 9, 'btw'] = 'LAAG'
    df_packages.loc[df_packages['btw'] == 21, 'btw'] = 'HOOG'
    df_packages = df_packages[['code', 'omschrijving', 'inkoopprijs excl btw', 'verkoopprijs excl btw', 'btw', 'groep', 'tegenrekening']]

    # create export dataframe for reusables
    print('reusables')
    df_reusables = pd.read_excel(product_file, 'reusables')
    df_reusables = df_reusables.rename(columns = {'productnaam':'omschrijving', 'inkoopprijs':'inkoopprijs excl btw', 'verkoopprijs per gebruik (**)':'verkoopprijs excl btw'})
    df_reusables = df_reusables[ df_reusables['verkoopprijs excl btw'].notna() ]
    df_reusables.loc[df_reusables['btw'] == 9, 'btw'] = 'LAAG'
    df_reusables.loc[df_reusables['btw'] == 21, 'btw'] = 'HOOG'
    df_reusables = df_reusables[['code', 'omschrijving', 'inkoopprijs excl btw', 'verkoopprijs excl btw', 'btw', 'groep', 'tegenrekening']]

    # create export dataframe for disposables
    print('disposables')
    df_disposables = pd.read_excel(product_file, 'disposables')
    df_disposables = df_disposables.rename(columns = {'productnaam':'omschrijving', 'inkoopprijs per stuk':'inkoopprijs excl btw', 'verkoopprijs per stuk (**)':'verkoopprijs excl btw'})
    df_disposables = df_disposables[ df_disposables['verkoopprijs excl btw'].notna() ]
    df_disposables.loc[df_disposables['btw'] == 9, 'btw'] = 'LAAG'
    df_disposables.loc[df_disposables['btw'] == 21, 'btw'] = 'HOOG'
    df_disposables = df_disposables[['code', 'omschrijving', 'inkoopprijs excl btw', 'verkoopprijs excl btw', 'btw', 'groep', 'tegenrekening']]

    # create export dataframe for medication
    print('medication')
    df_medication = pd.read_excel(product_file, 'medication')
    df_medication = df_medication.rename(columns = {'productnaam':'omschrijving', 'inkoopprijs per unit':'inkoopprijs excl btw', 'verkoopprijs per unit (**)':'verkoopprijs excl btw'})
    df_medication = df_medication[ df_medication['verkoopprijs excl btw'].notna() ]
    df_medication.loc[df_medication['btw'] == 9, 'btw'] = 'LAAG'
    df_medication.loc[df_medication['btw'] == 21, 'btw'] = 'HOOG'
    df_medication = df_medication[['code', 'omschrijving', 'inkoopprijs excl btw', 'verkoopprijs excl btw', 'btw', 'groep', 'tegenrekening']]

    # create export dataframe for lab
    print('lab')
    df_lab = pd.read_excel(product_file, 'lab')
    df_lab = df_lab.rename(columns = {'productnaam':'omschrijving', 'inkoopprijs':'inkoopprijs excl btw', 'verkoopprijs excl (**)':'verkoopprijs excl btw'})
    df_lab = df_lab[ df_lab['verkoopprijs excl btw'].notna() ]
    df_lab.loc[df_lab['btw'] == 9, 'btw'] = 'LAAG'
    df_lab.loc[df_lab['btw'] == 21, 'btw'] = 'HOOG'
    df_lab = df_lab[['code', 'omschrijving', 'inkoopprijs excl btw', 'verkoopprijs excl btw', 'btw', 'groep', 'tegenrekening']]

    df_products = pd.concat([df_packages, df_services, df_reusables, df_disposables, df_medication, df_lab], sort=False)
    with pd.ExcelWriter(export_file, date_format='dd-mm-yyyy', datetime_format='dd-mm-yyyy', engine='xlsxwriter') as writer:
        df_products.to_excel(writer, sheet_name='products', index=False)
        workbook = writer.book
        worksheet = writer.sheets['products']
        set_columns(workbook, worksheet, df_services)

        writer.save()

def update_pathway(pathway):
    