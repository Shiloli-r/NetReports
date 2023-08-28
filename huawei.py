import pandas as pd


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
    yellow_fill_format = workbook.add_format({"bg_color": "yellow"})

    worksheet.write_string("A1", "HOST NAME", cell_format=yellow_fill_format)
    worksheet.write_string("B1", "TOTAL PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("C1", "XGE PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("D1", "COMPLIANT PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("E1", "NON-COMPLIANT PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("F1", "NON-COMPLIANT OPEN PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("G1", "COMPLIANCE %", cell_format=yellow_fill_format)
    worksheet.write_string("H1", "PHYSICAL LOCATION", cell_format=yellow_fill_format)
    worksheet.write_string("I1", "COUNTRY", cell_format=yellow_fill_format)


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
    # Filter out rows where compliance == 100%
    for num, x in enumerate(df[6] == 1):
        if x:
            df = df.drop(num)

    # drop last row (it's a summary)
    df.drop(df.tail(1).index, inplace=True)

    # Transpose the data
    df = df.transpose()

    # write to the second sheet, i.e. Non-Compliance
    df.to_excel(writer, index=False, header=False, sheet_name='Non-Compliance', startcol=1)

    # --- Transpose the data to the Excel
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets["Non-Compliance"]

    # Convert compliance to %
    percent_format = workbook.add_format({"num_format": "0.00%"})

    # Add headings
    yellow_fill_format = workbook.add_format({"bg_color": "#eafa73"})
    non_compliant_fill_format = workbook.add_format({"bg_color": "#fca4a4", "font": "red"})
    interface_fill_format = workbook.add_format({"bg_color": "#5ff569"})

    worksheet.write_string("A1", "HOST NAME", cell_format=yellow_fill_format)
    worksheet.write_string("A2", "TOTAL PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("A3", "XGE PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("A4", "COMPLIANT PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("A5", "NON-COMPLIANT PORTS", cell_format=non_compliant_fill_format)
    worksheet.write_string("A6", "NON-COMPLIANT OPEN PORTS", cell_format=non_compliant_fill_format)
    worksheet.write_string("A7", "COMPLIANCE %", cell_format=yellow_fill_format)
    worksheet.write_string("A8", "PHYSICAL LOCATION", cell_format=yellow_fill_format)
    worksheet.write_string("A9", "COUNTRY", cell_format=yellow_fill_format)

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
    worksheet.set_row(6, None, percent_format)


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
    """
    Generates sheet 4 (compliance %). A bunch of Excel formulas to summarize the data
    :param writer:
    :param df:
    :return:
    """
    no_of_hosts = df.shape[0]

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = workbook.add_worksheet('Compliance %')
    yellow_fill_format = workbook.add_format({"bg_color": "yellow"})
    bold_format = workbook.add_format({'bold': True})

    worksheet.write_string("B1", "KE", cell_format=yellow_fill_format)
    worksheet.write_string("B8", "KE", cell_format=yellow_fill_format)
    worksheet.write_string("C1", "Total", cell_format=yellow_fill_format)

    worksheet.write_string("A2", "Total Number of switches", cell_format=bold_format)

    # Total number of switches
    worksheet.write_formula("B2", "=COUNTIF('Audit Data'!I2: I{}, \"Kenya\")".format(no_of_hosts))
    worksheet.write_number("C2", no_of_hosts - 1)

    # fully compliant
    worksheet.write_formula("C3", "=COUNTIF('Audit Data'!G2:G{}, \"100%\")".format(no_of_hosts))
    worksheet.write_formula("B3", "=COUNTIFS('Audit Data'!G2:G{}, \"100%\", 'Audit Data'!I2:I{}, \"Kenya\")".format(
        no_of_hosts, no_of_hosts))

    # partially compliant
    worksheet.write_formula("B4", "=B2-B3")
    worksheet.write_formula("C4", "=C2-C3")

    # fully non-compliant
    worksheet.write_formula("B5",
                            "=COUNTIFS('Audit Data'!G2:G{}, \"0%\", 'Audit Data'!I2:I{}, \"Kenya\")".format(no_of_hosts,
                                                                                                            no_of_hosts))
    worksheet.write_formula("C5", "=COUNTIF('Audit Data'!G2:G{}, \"0%\")".format(no_of_hosts))

    worksheet.write_string("A3", "No. of Switches that are fully compliant", cell_format=bold_format)
    worksheet.write_string("A4", "No. of Switches that are partially compliant", cell_format=bold_format)
    worksheet.write_string("A5", "No. of Switches that are fully non-compliant", cell_format=bold_format)

    worksheet.write_string("A9", "mab is configured", cell_format=bold_format)
    worksheet.write_string("A10", "authentication open is configured", cell_format=bold_format)


if __name__ == '__main__':
    dataframe = load_csv("Data/huawei.csv")
    wr = pd.ExcelWriter("output.xlsx", engine="xlsxwriter")

    # Generate the worksheets
    raw_df = generate_sheet1(wr, dataframe)

    generate_sheet2(wr, raw_df)
    generate_sheet3(wr, dataframe)
    generate_sheet4(wr, dataframe)

    wr.close()
