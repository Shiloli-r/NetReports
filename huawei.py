import pandas as pd
from openpyxl import Workbook

headers = ['HOST NAME', 'TOTAL PORTS', 'XGE PORTS', 'COMPLIANT PORTS', 'NON-COMPLIANT MAB PORTS', 'NON-COMPLIANT OPEN '
                                                                                                  'PORTS',
           'COMPLIANCE %', 'PHYSICAL LOCATION', 'COUNTRY']


def load_csv(filename):
    with open(filename) as f:
        num_cols = max(len(line.split(',')) for line in f)  # get max number of columns in file
        f.seek(0)
        return pd.read_csv(f, names=range(num_cols), skiprows=[0])


def generate_sheet1(writer, df):
    df[6] = df[6].str.replace("%", "")  # remove the % symbol on compliance column
    df[6] = (df[6].astype(float)) / 100  # convert to float & divide by 100

    # write to the first sheet, i.e. Audit Data - Start at row 1, insert headings at row 0 later
    df.to_excel(writer, index=False, header=False, sheet_name='Audit Data', startrow=1)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet1 = writer.sheets["Audit Data"]

    # Convert compliance to %
    percent_format = workbook.add_format({"num_format": "0.00%"})
    worksheet1.set_column(6, 6, None, percent_format)

    # Add headings
    fill_format = workbook.add_format({"bg_color": "yellow"})

    worksheet1.write_string("A1", "HOST NAME", cell_format=fill_format)
    worksheet1.write_string("B1", "TOTAL PORTS", cell_format=fill_format)
    worksheet1.write_string("C1", "XGE PORTS", cell_format=fill_format)
    worksheet1.write_string("D1", "COMPLIANT PORTS", cell_format=fill_format)
    worksheet1.write_string("E1", "NON-COMPLIANT PORTS", cell_format=fill_format)
    worksheet1.write_string("F1", "NON-COMPLIANT OPEN PORTS", cell_format=fill_format)
    worksheet1.write_string("G1", "COMPLIANCE %", cell_format=fill_format)
    worksheet1.write_string("H1", "PHYSICAL LOCATION", cell_format=fill_format)
    worksheet1.write_string("I1", "COUNTRY", cell_format=fill_format)


if __name__ == '__main__':
    dataframe = load_csv("Data/huawei.csv")
    wr = pd.ExcelWriter("output.xlsx", engine="xlsxwriter")
    generate_sheet1(wr, dataframe)
    wr.close()
