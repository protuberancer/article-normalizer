import tabula
import pandas as pd
import xlrd


# Функция для преобразования значений
def decode_cp1251(value):
    if isinstance(value, str):
        try:
            decoded_value = value.encode('latin1').decode('cp1251')
            return decoded_value
        except UnicodeEncodeError:
            return value
    else:
        return value


def pdf_to_df_list(pdf_file_path):
    # Read PDF file
    tables = tabula.read_pdf(pdf_file_path, pages='all', lattice=True, pandas_options={'header': None})

    # Декодирование значений внутри каждого DataFrame
    decoded_tables = [df.apply(lambda x: x.map(decode_cp1251) if x.dtype == 'object' else x) for df in tables]

    return decoded_tables


def list_to_excel(tables, excel_file_path):
    # Write each table to a separate sheet in the Excel file
    with pd.ExcelWriter(excel_file_path) as writer:
        for i, table in enumerate(tables):
            table.to_excel(writer, sheet_name=f'Лист{i+1}', header=False, index=False)


def pdf_reader(pdf_file_path):
    raw_tables = pdf_to_df_list(pdf_file_path)
    tables = []
    for df in raw_tables:
        if not df.empty and not df.isnull().all().all():
            df.dropna(axis=0, how='all', inplace=True)
            df = df.apply(lambda x: x.map(lambda x: x.replace('\r', ' ') if isinstance(x, str) else x))
            df = df.apply(lambda x: x.map(lambda x: x.replace('\n', ' ') if isinstance(x, str) else x))
            tables.append(df)
    return tables
