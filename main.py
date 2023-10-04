import pandas as pd
from pyxlsb import open_workbook
import datetime

from datetime import datetime, timedelta
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to convert Excel serial date to datetime and format it as "dd.mm.yyyy"
def excel_serial_date_to_datetime(serial_date):
    if isinstance(serial_date, float):
        epoch_start = datetime(1900, 1, 1)
        delta = timedelta(days=int(serial_date) - 2)
        return epoch_start + delta
    return None

# Function to correct date columns in DataFrame
def correct_column_date(df, column_name):
    df[column_name] = df[column_name].apply(excel_serial_date_to_datetime)
    df[column_name] = df[column_name].dt.strftime('%d.%m.%Y')
    return df

# Function to process and save DataFrame as CSV
def process_and_save_csv(df, name, type):
    logging.info(f"Processing {name} - {type}")
    df.to_csv(f"{name}_{type}.csv", sep=';', decimal=',', encoding = "utf-8-sig", index=False, mode='a')
    logging.info(f"{name}_{type}.csv saved successfully")

if __name__ == "__main__":
    xlsb_file = 'input.xlsb'
    workbook = open_workbook(xlsb_file)
    sheet_names = workbook.sheets

    for sheet_name in sheet_names:

        try:
            int(sheet_name)
        except Exception as error:
            print("An exception occurred:", type(error).__name__, "–", error)
            continue

        with workbook.get_sheet(sheet_name) as sheet:
            data = [[item.v for item in row] for row in sheet.rows()]
            df = pd.DataFrame(data[1:], columns=data[0])

            row_index = df[df['PROJEKTNUMMER'] == 'Mittelabruf:'].index[0]

            # Process data for Einnahme.csv
            einnahme = df[row_index+1:]
            # einnahme['BUCHUNGSDATUM_ZUKUNFT'] = ""
            einnahme.columns = df.iloc[row_index+1]
            einnahme = einnahme.iloc[1:].dropna(subset=['BELEGNUMMER'], how='all')
            einnahme = correct_column_date(einnahme, "BELEGDATUM")
            einnahme = correct_column_date(einnahme, "BUCHUNGSDATUM")
            einnahme = correct_column_date(einnahme, "ZAHLUNGSDATUM")
            einnahme['PROJEKTNUMMER'] = einnahme['PROJEKTNUMMER'].astype(int)
            einnahme['ZAHLUNGSWEISE'] = einnahme['ZAHLUNGSWEISE'].astype(int)
            einnahme['BETRAG'] = einnahme['BETRAG'].apply(lambda x: x.replace(',', '.') if isinstance(x, str) else x).astype(float)
            einnahme = einnahme.loc[:, einnahme.columns.notna()]


            process_and_save_csv(einnahme, sheet_name, "_einnahme")

            # Process data for Ausgabe.csv
            ausgabe = df[:row_index].dropna(subset=['BELEGNUMMER'], how='all')
            ausgabe = ausgabe[ausgabe['BETRAG'] != 0.0]
            ausgabe['BUCHUNGSDATUM_ZUKUNFT'] = ""
            #print(ausgabe[["BELEGDATUM"]].to_string(index=False))

            ausgabe = correct_column_date(ausgabe, "BELEGDATUM")

            ausgabe = correct_column_date(ausgabe, "BUCHUNGSDATUM")
            ausgabe = correct_column_date(ausgabe, "ZAHLUNGSDATUM")
            ausgabe = ausgabe.dropna(thresh=ausgabe.shape[1] - 6)
            ausgabe['PROJEKTNUMMER'] = ausgabe['PROJEKTNUMMER'].astype(int)
            ausgabe['ZAHLUNGSWEISE'] = ausgabe['ZAHLUNGSWEISE'].astype(int)
            ausgabe['BETRAG'] = ausgabe['BETRAG'].apply(lambda x: x.replace(',', '.') if isinstance(x, str) else x).astype(float)
            ausgabe['ANTEIL'] = ausgabe['ANTEIL'].apply(lambda x: x.replace(',', '.') if isinstance(x, str) else x).astype(float)
            process_and_save_csv(ausgabe, sheet_name, "_ausgabe")

    logging.info("CSV files created successfully.")
