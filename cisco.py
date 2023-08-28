import pandas as pd
from datetime import date


def format_headings_and_percentages(writer, sheet_name):
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # Add headings
    yellow_fill_format = workbook.add_format({"bg_color": "#f0fc03", "bold": True})

    worksheet.write_string("A1", "HOST NAME", cell_format=yellow_fill_format)
    worksheet.write_string("B1", "TOTAL PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("C1", "TRUNK PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("D1", "COMPLIANT PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("E1", "NON-COMPLIANT PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("F1", "COMPLIANCE %", cell_format=yellow_fill_format)

    percent_format = workbook.add_format({"num_format": "0.00%"})
    worksheet.set_column(5, 5, None, percent_format)


def generate_sheet1(writer, df):
    dataframe.to_excel(wr, index=False, header=False, startrow=1, sheet_name='Audit Data')
    format_headings_and_percentages(wr, 'Audit Data')


if __name__ == '__main__':
    # generate filename to be expected
    today = date.today()
    # filename = today.strftime("%Y%m%d_audit_row_hw_sw.csv")
    filename = "Data/cisco.xlsx"

    # load data
    dataframe = pd.read_excel('Data/cisco.xlsx')

    # Add 'Compliance % column' to dataframe
    wr = pd.ExcelWriter("output.xlsx", engine="xlsxwriter")
    compliance = [(dataframe['TRUNK PORTS'][x] + dataframe['COMPLIANT PORTS'][x]) / dataframe['TOTAL PORTS'][x] for x in
                  range(dataframe.shape[0])]
    dataframe.insert(5, "Compliance %", compliance)

    wr.close()
