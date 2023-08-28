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


if __name__ == '__main__':
    df = load_csv("Data/huawei.csv")

    writer = pd.ExcelWriter("output.xlsx", engine="xlsxwriter")
    df[6] = df[6].str.replace("%", "")  # remove the % symbol on compliance column
    df[6] = (df[6].astype(float)) / 100  # convert to float & divide by 100

    # write to the first sheet, i.e. Audit Data
    df.to_excel(writer, index=False, header=False, sheet_name='Audit Data')

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets["Audit Data"]

    # Convert compliance to %
    percent_format = workbook.add_format({"num_format": "0.00%"})
    worksheet.set_column(6, 6, None, percent_format)

    writer.close()
