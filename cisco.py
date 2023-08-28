import pandas as pd
from datetime import date

address = {
    "Ke680Avenue_BranchSwitch": ["680 Branch", "Kenya"],
    "ABC_BranchSwitch": ["ABC Branch", "Kenya"],
    "ARUSH_SW": ["Arusha Branch Tz", "Tanzania"],
    "BUGOLOBI_SWII": ["Bugolobi Branch UG", "Uganda"],
    "BUNGOMA-SW": ["Bungoma Branch", "Kenya"],
    "BuruBuru_BranchSwitch": ["Buruburu Branch", "Kenya"],
    "BURUBURU_SW": ["Buruburu Switch", "Kenya"],
    "Busia_Switch": ["Busia Switch", "Kenya"],
    "CarDuka_sw": ["Carduka Switch", "Kenya"],
    "CHANGAMWE_SW_NEW": ["Changamwe", "Kenya"],
    "CHWELE_SW": ["Chwele Mombasa", "Kenya"],
    "CIATAMALL_SW": ["Ciata Mall", "Kenya"],
    "DIANI_SWITCH_2960_12p": ["Diani Branch", "Kenya"],
    "Diani_ATMSw": ["Diani Branch", "Kenya"],
    "DianiSw2": ["Diani Branch", "Kenya"],
    "Eldoret_SW1": ["Eldoret Branch", "Kenya"],
    "Eldoret_SW2": ["Eldoret Branch", "Kenya"],
    "EVEN_PARK_SW": ["Embakasi Even Park", "Kenya"],
    "EMBU-SW": ["Embu Branch", "Kenya"],
    "NCBA-FOREST-MALL-SW": ["Forest Mall Branch", "Kenya"],
    "GALLERIA_SWITCH": ["Galleria Branch", "Kenya"],
    "GALLERIA_SW2": ["Galleria Branch", "Kenya"],
    "GARDEN_CITY_ATM": ["Galleria Branch", "Kenya"],
    "Garden-City-SW": ["Galleria Branch", "Kenya"],
    "Gikomba-Switch": ["Gikomba Branch", "Kenya"],
    "GreenSpan_SW": ["Greenspan Branch", "Kenya"],
    "ICRAF_SW": ["ICRAF UN Gigiri", "Kenya"],
    "ILRI_SW": ["ILRI Agency", "Kenya"],
    "JKIA_EXPORTS_SW": ["JKIA Exports Branch", "Kenya"],
    "JKIA_IMPORT_SW": ["JKIA Imports Branch", "Kenya"],
    "JUNCTION_BR_SW48": ["Junction Mall", "Kenya"],
    "KAHAWA_SW": ["Kahawa Sukari Branch", "Kenya"],
    "Kakamega_switch": ["Kakamega Branch", "Kenya"],
    "NCBA_KAMAKIS_SWITCH": ["Kamakis Branch", "Kenya"],
    "KARATINA-SW1": ["Karatina Branch", "Kenya"],
    "KAREN_SWITCH": ["Karen Switch", "Kenya"],
    "NCBA_KARIAKOO": ["Kariakoo Branch TZ", "Tanzania"],
    "RW-KAYONZA-SW": ["Kayonza Branch RW", "Rwanda"],
    "KENOL_SW": ["Kenol Branch", "Kenya"],
    "KERICHO_SW": ["Kericho Branch", "Kenya"],
    "RW-KH-CORE-SW1": ["Kigali Heights", "Rwanda"],
    "Kijitonyama_SW": ["Kijitonyama TZ", "Tanzania"],
    "KILIFIBR_Switch": ["", "Kenya"],
    "Kisumu_Switch": ["Kisumu Branch", "Kenya"],
}


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


def generate_sheet3(writer, df):
    # drop all columns after "compliance %" column
    df.drop(df.iloc[:, 6:], inplace=True, axis=1)
    df.drop(df.tail(1).index, inplace=True)  # drop last row (total)
    ipaddress = [df['HOST NAME'][x].split()[0] for x in range(df.shape[0])]
    key = [df['HOST NAME'][x].split()[2] for x in range(df.shape[0])]
    physical_address = [address[key[x]][0] if address.get(key[x]) is not None else "" for x in range(df.shape[0])]
    country = [address[key[x]][1] if address.get(key[x]) is not None else "" for x in range(df.shape[0])]

    dataframe.insert(0, "IP ADDRESS", ipaddress)
    dataframe.insert(2, "PHYSICAL ADDRESS", physical_address)
    dataframe.insert(3, "COUNTRY", country)
    df.to_excel(writer, index=False, header=False, sheet_name='Compliance', startrow=1)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets["Compliance"]

    # Add headings
    yellow_fill_format = workbook.add_format({"bg_color": "#f0fc03", "bold": True})

    worksheet.write_string("A1", "IP ADDRESS", cell_format=yellow_fill_format)
    worksheet.write_string("B1", "HOST NAME", cell_format=yellow_fill_format)
    worksheet.write_string("C1", "PHYSICAL ADDRESS", cell_format=yellow_fill_format)
    worksheet.write_string("D1", "COUNTRY", cell_format=yellow_fill_format)
    worksheet.write_string("E1", "TOTAL PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("F1", "TRUNK PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("G1", "COMPLIANT PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("H1", "NON-COMPLIANT PORTS", cell_format=yellow_fill_format)
    worksheet.write_string("I1", "COMPLIANCE %", cell_format=yellow_fill_format)
    worksheet.write_string("A{}".format(df.shape[0]+2), "TOTALS", cell_format=yellow_fill_format)

    last_row = df.shape[0]+1
    worksheet.write_formula("E{}".format(last_row+1), "=SUM(E2:E{})".format(last_row), cell_format=yellow_fill_format)
    worksheet.write_formula("F{}".format(last_row+1), "=SUM(F2:F{})".format(last_row), cell_format=yellow_fill_format)
    worksheet.write_formula("G{}".format(last_row+1), "=SUM(G2:G{})".format(last_row), cell_format=yellow_fill_format)
    worksheet.write_formula("H{}".format(last_row+1), "=SUM(H2:H{})".format(last_row), cell_format=yellow_fill_format)

    percent_format = workbook.add_format({"num_format": "0.00%"})
    worksheet.set_column(8, 8, None, percent_format)
    worksheet.write_formula("I{}".format(last_row + 1),
                            "=(F{}+G{})/E{}".format(last_row + 1, last_row + 1, last_row + 1), cell_format=percent_format)


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
    generate_sheet3(wr, dataframe)

    wr.close()
