import pandas as pd
from openpyxl import Workbook


def load_csv(filename):
    with open(filename) as f:
        num_cols = max(len(line.split(',')) for line in f)  # get max number of columns in file
        f.seek(0)
        return pd.read_csv(f, names=range(num_cols), skiprows=[0])


def format_percent_and_headings(writer, sheet_name):
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # Convert compliance to %
    percent_format = workbook.add_format({"num_format": "0.00%"})
    worksheet.set_column(6, 6, None, percent_format)

    # Add headings
    fill_format = workbook.add_format({"bg_color": "yellow"})

    worksheet.write_string("A1", "HOST NAME", cell_format=fill_format)
    worksheet.write_string("B1", "TOTAL PORTS", cell_format=fill_format)
    worksheet.write_string("C1", "XGE PORTS", cell_format=fill_format)
    worksheet.write_string("D1", "COMPLIANT PORTS", cell_format=fill_format)
    worksheet.write_string("E1", "NON-COMPLIANT PORTS", cell_format=fill_format)
    worksheet.write_string("F1", "NON-COMPLIANT OPEN PORTS", cell_format=fill_format)
    worksheet.write_string("G1", "COMPLIANCE %", cell_format=fill_format)
    worksheet.write_string("H1", "PHYSICAL LOCATION", cell_format=fill_format)
    worksheet.write_string("I1", "COUNTRY", cell_format=fill_format)


def generate_sheet1(writer, df):
    """
    Copies the raw CSV data into the first sheet of the Excel Workbook
    :param writer:
    :param df:
    :return:
    """
    df[6] = df[6].str.replace("%", "")  # remove the % symbol on compliance column
    df[6] = (df[6].astype(float)) / 100  # convert to float & divide by 100

    # write to the first sheet, i.e. Audit Data - Start at row 1, insert headings at row 0 later
    df.to_excel(writer, index=False, header=False, sheet_name='Audit Data', startrow=1)

    # Add required formatting
    format_percent_and_headings(wr, "Audit Data")
    return df


def generate_sheet2(writer, df):
    # write to the second sheet, i.e. Non-Compliance
    df.to_excel(writer, index=False, header=False, sheet_name='Non-Compliance')


def generate_sheet3(writer, df):
    """
    Writes the 8 labelled columns into the third worksheet ("Compliance")
    :param writer:
    :param df:
    :return:
    """
    # drop all columns after column 9 ("country" column)
    df.drop(df.iloc[:, 9:], inplace=True, axis=1)
    df.to_excel(writer, index=False, header=False, sheet_name='Compliance', startrow=1)

    # format percentages and add column headers
    format_percent_and_headings(wr, "Compliance")


def generate_sheet4(writer, df):
    pass


if __name__ == '__main__':
    dataframe = load_csv("Data/huawei.csv")
    wr = pd.ExcelWriter("output.xlsx", engine="xlsxwriter")

    # Generate the worksheets
    raw_df = generate_sheet1(wr, dataframe)

    generate_sheet2(wr, raw_df)
    generate_sheet3(wr, dataframe)
    generate_sheet4(wr, dataframe)

    wr.close()
