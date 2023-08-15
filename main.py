import pandas as pd
from pyxlsb import open_workbook
from datetime import datetime, timedelta

# Function to convert Excel serial date to datetime and format it as "dd.mm.yyyy"


def excel_serial_date_to_datetime(serial_date):
    if isinstance(serial_date, float):
        epoch_start = datetime(1900, 1, 1)
        # Subtract 2 days for Excel's date system
        delta = timedelta(days=int(serial_date) - 2)
        return epoch_start + delta
    return None

# Function to correct date columns in DataFrame


def correct_column_date(df, column_name):
    df[str(column_name)] = df[str(column_name)].apply(
        excel_serial_date_to_datetime)
    df[str(column_name)] = df[str(column_name)].dt.strftime('%d.%m.%Y')
    return df

# Function to create and save CSV file


def create_csv(df, name, type):
    print(''.join([name, str(type), ".csv"]))
    df.to_csv(''.join([name, str(type), ".csv"]),
              sep=';', decimal=',', index=False)
    return None


if __name__ == "__main__":
    xlsb_file = 'input.xlsb'
    workbook = open_workbook(xlsb_file)

    # Get the list of sheet names in the workbook
    sheet_names = workbook.sheets

    # Iterate through each sheet
    for sheet_name in sheet_names:
        with workbook.get_sheet(sheet_name) as sheet:
            data = []

            # Extract data from the sheet
            for row in sheet.rows():
                data.append([item.v for item in row])

            # Convert the extracted data to a DataFrame
            df = pd.DataFrame(data[1:], columns=data[0])

            # Find the index of the row with 'Mittelabruf:'
            row_index = df[df['PROJEKTNUMMER'] == 'Mittelabruf:'].index[0]

            # Process data for Einnahme.csv
            einnahme = df[row_index+1:]
            new_header = df.iloc[row_index+1]
            einnahme.columns = new_header
            einnahme = einnahme.reset_index(drop=True)
            einnahme = einnahme[1:]
            einnahme = einnahme.dropna(subset=['BELEGNUMMER'], how='all')
            einnahme = correct_column_date(einnahme, "BELEGDATUM")
            einnahme = correct_column_date(einnahme, "BUCHUNGSDATUM")
            einnahme = correct_column_date(einnahme, "ZAHLUNGSDATUM")
            einnahme['PROJEKTNUMMER'] = einnahme['PROJEKTNUMMER'].astype(int)
            einnahme['ZAHLUNGSWEISE'] = einnahme['ZAHLUNGSWEISE'].astype(int)
            einnahme['BETRAG'] = einnahme['BETRAG'].apply(
                lambda x: x.replace(',', '.') if isinstance(x, str) else x)
            einnahme['BETRAG'] = einnahme['BETRAG'].astype(float)
            create_csv(einnahme, sheet_name, "_einnahme")

            # Process data for Ausgabe.csv
            ausgabe = df[:row_index]
            ausgabe = ausgabe.dropna(subset=['BELEGNUMMER'], how='all')
            ausgabe = ausgabe[ausgabe['BETRAG'] != 0.0]
            ausgabe = correct_column_date(ausgabe, "BELEGDATUM")
            ausgabe = correct_column_date(ausgabe, "BUCHUNGSDATUM")
            ausgabe = correct_column_date(ausgabe, "ZAHLUNGSDATUM")
            ausgabe = ausgabe.dropna(thresh=ausgabe.shape[1] - 6)
            ausgabe['PROJEKTNUMMER'] = ausgabe['PROJEKTNUMMER'].astype(int)
            ausgabe['ZAHLUNGSWEISE'] = ausgabe['ZAHLUNGSWEISE'].astype(int)
            ausgabe['BETRAG'] = ausgabe['BETRAG'].apply(
                lambda x: x.replace(',', '.') if isinstance(x, str) else x)
            ausgabe['BETRAG'] = ausgabe['BETRAG'].astype(float)
            ausgabe['ANTEIL'] = ausgabe['ANTEIL'].apply(
                lambda x: x.replace(',', '.') if isinstance(x, str) else x)
            ausgabe['ANTEIL'] = ausgabe['ANTEIL'].astype(float)
            create_csv(ausgabe, sheet_name, "_ausgabe")

    print("CSV files created successfully.")
