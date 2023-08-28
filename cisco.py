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
    df.to_excel(writer, index=False, header=False, startrow=1, sheet_name='Audit Data')
    format_headings_and_percentages(wr, 'Audit Data')


def generate_sheet2(writer, df):
    # Filter out rows where compliance == 100%
    for num, x in enumerate(df['Compliance %'] == 1):
        if x:
            df = df.drop(num)

    # drop last row (it's a summary)
    df.drop(df.tail(1).index, inplace=True)

    # Transpose the data
    df = df.transpose()

    # write to the second sheet, i.e. Non-Compliance
    df.to_excel(writer, index=False, header=False, sheet_name='Non-Compliant Switches', startcol=1)

    # --- Transpose the data to the Excel
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets["Non-Compliant Switches"]

    # Convert compliance to %
    percent_format = workbook.add_format({"num_format": "0.00%"})

    # Add headings
    yellow_fill_format = workbook.add_format({"bg_color": "#f0fc03", "bold": True})
    non_compliant_fill_format = workbook.add_format({"bg_color": "#fca4a4", "font": "red"})
    interface_fill_format = workbook.add_format({"bg_color": "#7dfa9e"})

    worksheet.write_string("A1", "HOST NAME", cell_format=yellow_fill_format)
    worksheet.write_string("A2", "TOTAL PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("A3", "TRUNK PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("A4", "COMPLIANT PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("A5", "NON-COMPLIANT PORTS", cell_format=non_compliant_fill_format)
    worksheet.write_string("A6", "COMPLIANCE %", cell_format=yellow_fill_format)

    worksheet.conditional_format('A1:XFD1048576',
                                 {'type': 'text',
                                  'criteria': 'containing',
                                  'value': 'non-compliant',
                                  'format': non_compliant_fill_format})

    worksheet.conditional_format('A1:XFD1048576',
                                 {'type': 'text',
                                  'criteria': 'containing',
                                  'value': 'interface',
                                  'format': interface_fill_format})
    worksheet.set_row(5, None, percent_format)


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

    generate_sheet1(wr, dataframe)
    generate_sheet2(wr, dataframe)

    wr.close()
