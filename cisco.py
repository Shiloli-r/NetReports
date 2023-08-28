import pandas as pd
from datetime import date

address = {
    "Ke680Avenue_BranchSwitch": ["680 Branch", "Kenya"],
    "ABC_BranchSwitch": ["ABC Branch", "Kenya"],
    "ARUSHA-SW": ["Arusha Branch Tz", "Tanzania"],
    "BUGOLOBI_SWI": ["Bugolobi Branch UG", "Uganda"],
    "BUGOLOBI_SWII": ["Bugolobi Branch UG", "Uganda"],
    "BUNGOMA-SW": ["Bungoma Branch", "Kenya"],
    "BuruBuru_BranchSwitch": ["Buruburu Branch", "Kenya"],
    "BURUBURU_SW": ["Buruburu Switch", "Kenya"],
    "Busia-Switch": ["Busia Switch", "Kenya"],
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
    "Gikomba-Swtch": ["Gikomba Branch", "Kenya"],
    "GreenSpan_Sw": ["Greenspan Branch", "Kenya"],
    "ICRAF_SW": ["ICRAF UN Gigiri", "Kenya"],
    "ILRI_SW": ["ILRI Agency", "Kenya"],
    "JKIA_EXPORTS_SW": ["JKIA Exports Branch", "Kenya"],
    "JKIA_IMPORT-SW": ["JKIA Imports Branch", "Kenya"],
    "JUNCTION_BR_SW48": ["Junction Mall", "Kenya"],
    "KAHAWA_SW": ["Kahawa Sukari Branch", "Kenya"],
    "Kakamega_switch": ["Kakamega Branch", "Kenya"],
    "NCBA_KAMAKIS_SWITCH": ["Kamakis Branch", "Kenya"],
    "KARATINA-SW1": ["Karatina Branch", "Kenya"],
    "KAREN_SWITCH": ["Karen Switch", "Kenya"],
    "NCBA_KARIAKOO": ["Kariakoo Branch TZ", "Tanzania"],
    "RW-KAYONZA-SW": ["Kayonza Branch RW", "Rwanda"],
    "KENOL_SW": ["Kenol Branch", "Kenya"],
    "KERICHO_Sw": ["Kericho Branch", "Kenya"],
    "RW-KH-CORE-SW1": ["Kigali Heights", "Rwanda"],
    "Kijitonyama_SW": ["Kijitonyama TZ", "Tanzania"],
    "KILIFIBR_Switch": ["Kilifi Branch", "Kenya"],
    "KILIMANI_SW": ["Kilimani Branch", "Kenya"],
    "Kisii_SW": ["Kisii Branch", "Kenya"],
    "Kisumu_Switch": ["Kisumu Branch", "Kenya"],
    "KSM2-ASW": ["Kisumu Branch", "Kenya"],
    "CARGO_CENTRE_SW": ["Kisumu Cargo", "Kenya"],
    "KSM_Loop_SW": ["Kisumu Loopstore", "Kenya"],
    "KITALE_Branch_Sw": ["Kitale Branch", "Kenya"],
    "Kitengela_BranchSwitch": ["Kitengela Branch", "Kenya"],
    "LAVINGTON_SW": ["Lavington Branch", "Kenya"],
    "LUNGALUNGA_BranchSwitch": ["LungaLunga Branch", "Kenya"],
    "MAASAI_MALL_SW": ["Maasai Mall Branch", "Kenya"],
    "Machakos_SW": ["Machakos Branch", "Kenya"],
    "MALINDIBR_Switch": ["Malindi Branch", "Kenya"],
    "MAMA_NGINA_SW": ["Mama Ngina", "Kenya"],
    "Mamlaka_SW": ["Mamlaka Branch", "Kenya"],
    "MBV-Switch": ["MBV", "Uganda"],
    "MERU-ASW": ["Meru Branch", "Kenya"],
    "MIGORI_SW": ["Migori Branch", "Kenya"],
    "MITCHELLECOTT_SW": ["Mitchell Cotts Branch", "Kenya"],
    "MOI_AVE_SW3": ["Moi Avenue Mombasa", "Kenya"],
    "NCBA_SWT5": ["Moi Avenue Mombasa", "Kenya"],
    "NCBA_SW6_MOI": ["Moi Avenue Mombasa", "Kenya"],
    "MOI_AVE_SW1": ["Moi Avenue Mombasa", "Kenya"],
    "MOI_AVE_SW2": ["Moi Avenue Mombasa Branch", "Kenya"],
    "MOI_AVE_LOOP_SW": ["Moi Avenue Mombasa Loopstore", "Kenya"],
    "CBA-NAK-SW": ["Mombasa Nkurumah Branch", "Kenya"],
    "MURANGA_SW": ["Muranga", "Kenya"],
    "RW-MUSANZE-SW": ["Musanze", "Rwanda"],
    "MWANZA-SW": ["Mwanza Branch TZ", "Tanzania"],
    "Mwembe_Tayari_Switch": ["Mwembe Tayari Mombasa", "Kenya"],
    "NRB_Hosi_Sw": ["Nairobi Hospital Branch", "Kenya"],
    "NAIVASHA_SW": ["Naivasha Branch", "Kenya"],
    "Two-Rivers-SW": ["Two Rivers Mall", "Kenya"],
    "UG-NAKASERO-BR-SW01": ["Nakasero Branch UG", "Uganda"],
    "UG-NAKASERO-BR-SW02": ["Nakasero Branch UG", "Uganda"],
    "NAKURU2_SW": ["Nakuru Branch", "Kenya"],
    "Nanyuki_SW": ["Nanyuki Branch", "Kenya"],
    "NAROK_SW": ["Narok Branch", "Kenya"],
    "Ngong-Switch": ["Ngong Branch", "Kenya"],
    "NKURUMAH_NSSF_SW": ["Nkururmah Branch Mombasa", "Kenya"],
    "NyaliSW1": ["Nyali Branch", "Kenya"],
    "NYALI_SWITCH": ["Nyali Branch", "Kenya"],
    "NYERERE-SW-02": ["Nyerere Branch TZ", "Tanzania"],
    "NYERERE-SW": ["Nyerere Branch TZ", "Tanzania"],
    "Parkside_SW": ["Parksiide Westlands", "Kenya"],
    "Prestige_SW2": ["Prestige", "Kenya"],
    "PRUDENTIAL-ASW": ["Prudential Branch", "Kenya"],
    "PSSSF_TRAINING": ["PSSSF Tanzania", "Tanzania"],
    "REGAL_PLAZA_SW": ["Regal Plaza", "Kenya"],
    "RIVERROAD-SW": ["River road Branch", "Kenya"],
    "Riverside_BranchSwitch": ["Riverside Branch", "Kenya"],
    "RongaiBRV-SW": ["Rongai Branch", "Kenya"],
    "RW-KH-HQ-SW1": ["RW HQ Access SW", "Rwanda"],
    "RW-DT-SW": ["Rwanda DownTown Branch", "Rwanda"],
    "RW-NYABUGOGO-SWITCH": ["Rwanda Nyabugogo Branch", "Rwanda"],
    "UG-HQ-SW2": ["Ruwenzori Branch UG", "Uganda"],
    "UG-BR-SW1": ["Ruwenzori Branch UG", "Uganda"],
    "5th-Floor_Switch": ["Ruwenzori Branch UG", "Uganda"],
    "5TH_TRAININGROOM_SW": ["Ruwenzori Branch UG", "Uganda"],
    "SAMEER_MEZZ_SW1": ["Sameer Branch", "Kenya"],
    "SAMEERBR_SWITCH": ["Sameer Branch", "Kenya"],
    "SAMORA-SW": ["Samora Branch TZ", "Tanzania"],
    "SARIT_BRANCH_SWITCH": ["Sarit Branch", "Kenya"],
    "SARIT_LOOP_SW": ["Sarit Loopstore", "Kenya"],
    "NYERERE-SW-03": ["Tanzania Nyerere Branch", "Tanzania"],
    "TM-ASW": ["The Mall Branch", "Kenya"],
    "NCBA_THIKA_SWITCH": ["Thika Branch", "Kenya"],
    "TRM_SWITCH2": ["TRM Mall", "Kenya"],
    "TRM_Switch": ["TRM Mall", "Kenya"],
    "2_RIVERS_ATM_SW": ["Two Rivers Mall", "Kenya"],
    "AGG-SW-HQ-TZ": ["TZ aggregation", "Tanzania"],
    "04-ACCESS-SW-HQ-TZ": ["TZ HQ Switch", "Tanzania"],
    "06-ACCESS-SW-HQ-TZ": ["TZ HQ Switch", "Tanzania"],
    "03-ACCESS-SW-HQ-TZ": ["TZ HQ Switch", "Tanzania"],
    "01-ACCESS-SW-HQ-TZ": ["TZ HQ Switch", "Tanzania"],
    "02-ACCESS-SW-HQ-TZ": ["TZ HQ Switch", "Tanzania"],
    "08-ACCESS-SW-HQ-TZ": ["TZ HQ Switch", "Tanzania"],
    "07-ACCESS-SW-HQ-TZ": ["TZ HQ Switch", "Tanzania"],
    "UG-HQSWITCH-01": ["UG HQ Switch", "Uganda"],
    "UG-HQSWITCH-02": ["Uganda TWED", "Uganda"],
    "US_Embassy_SW": ["Us Embassy", "Kenya"],
    "NCBA_UTAWALA_SWITCH": ["Utawala Branch", "Kenya"],
    "NCBA_VILLAGE_MKT_SW": ["Utawala Branch", "Kenya"],
    "DR-CAMPUS-SWII": ["Village Market Mall", "Kenya"],
    "DR-CAMPUS-SWI": ["Wabera Building", "Kenya"],
    "WABERA_SW": ["Wabera 3rd Building", "Kenya"],
    "Wabera_3rdFlr": ["Wabera Building  3rd Floor", "Kenya"],
    "SW2-WABERA-4FLR": ["Wabera Building 4th Floor", "Kenya"],
    "SW-WABERA-4FLR": ["Wabera Building 4th Floor", "Kenya"],
    "NCBA_WABERA_5thFLOOR": ["Wabera Building 5th Floor", "Kenya"],
    "WARWICK_SW": ["Warwick Branch", "Kenya"],
    "WATAMUBR_Switch": ["Watamu Branch", "Kenya"],
    "Westlands_BranchSw": ["Westlands Branch", "Kenya"],
    "WorldVision_SW": ["World Vision", "Kenya"],
    "WOTE_SW": ["Wote Branch", "Kenya"],
    "Yaya_SWI": ["Yaya Branch", "Kenya"],
    "Yaya_SWII": ["Yaya Branch", "Kenya"],
    "YAYA_LOOP_SW": ["Yaya Loopstore", "Kenya"],
    "NCBA_ZANZIBAR_SWITCH": ["Zanzibar Branch TZ", "Tanzania"]
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


def adjust_col_sizes(df, writer, sheet_name):
    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), 20)
        col_idx = df.columns.get_loc(column)
        writer.sheets[sheet_name].set_column(col_idx, col_idx, column_length)


def generate_sheet1(writer, df):
    df.to_excel(writer, index=False, header=False, startrow=1, sheet_name='Audit Data')
    format_headings_and_percentages(wr, 'Audit Data')

    adjust_col_sizes(df, writer, 'Audit Data')


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

    adjust_col_sizes(df, writer, 'Non-Compliant Switches')


def generate_sheet3(writer, df):
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
    worksheet.write_string("A{}".format(df.shape[0] + 2), "TOTALS", cell_format=yellow_fill_format)

    last_row = df.shape[0] + 1
    worksheet.write_formula("E{}".format(last_row + 1), "=SUM(E2:E{})".format(last_row), cell_format=yellow_fill_format)
    worksheet.write_formula("F{}".format(last_row + 1), "=SUM(F2:F{})".format(last_row), cell_format=yellow_fill_format)
    worksheet.write_formula("G{}".format(last_row + 1), "=SUM(G2:G{})".format(last_row), cell_format=yellow_fill_format)
    worksheet.write_formula("H{}".format(last_row + 1), "=SUM(H2:H{})".format(last_row), cell_format=yellow_fill_format)

    percent_format = workbook.add_format({"num_format": "0.00%"})
    worksheet.set_column(8, 8, None, percent_format)
    worksheet.write_formula("I{}".format(last_row + 1),
                            "=(F{}+G{})/E{}".format(last_row + 1, last_row + 1, last_row + 1),
                            cell_format=percent_format)

    adjust_col_sizes(df, writer, 'Compliance')


def generate_sheet4(writer, df):
    no_of_hosts = df.shape[0] + 1

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = workbook.add_worksheet("Compliance per Country")

    yellow_fill_format = workbook.add_format({"bg_color": "#f0fc03", "bold": True})
    worksheet.write_string("B1", "KE", cell_format=yellow_fill_format)
    worksheet.write_string("C1", "UG", cell_format=yellow_fill_format)
    worksheet.write_string("D1", "RW", cell_format=yellow_fill_format)
    worksheet.write_string("E1", "TZ", cell_format=yellow_fill_format)
    worksheet.write_string("F1", "Total", cell_format=yellow_fill_format)

    worksheet.write_string("A2", "Total Number of Switches")
    worksheet.write_string("A3", "No. of Switches that are fully compliant")
    worksheet.write_string("A4", "No. of Switches that are partially compliant")
    worksheet.write_string("A5", "No. of Switches that are fully non-compliant")

    worksheet.write_string("B10", "KE", cell_format=yellow_fill_format)
    worksheet.write_string("C10", "UG", cell_format=yellow_fill_format)
    worksheet.write_string("D10", "RW", cell_format=yellow_fill_format)
    worksheet.write_string("E10", "TZ", cell_format=yellow_fill_format)
    worksheet.write_string("F10", "Total", cell_format=yellow_fill_format)
    worksheet.write_string("A11", "mab is configured")
    worksheet.write_string("A12", "authentication open is configured")

    # Total Number of Switches
    worksheet.write_formula("B2", "=COUNTIF(Compliance!D2:D{}, \"Kenya\")".format(no_of_hosts))
    worksheet.write_formula("C2", "=COUNTIF(Compliance!D2:D{}, \"Uganda\")".format(no_of_hosts))
    worksheet.write_formula("D2", "=COUNTIF(Compliance!D2:D{}, \"Rwanda\")".format(no_of_hosts))
    worksheet.write_formula("E2", "=COUNTIF(Compliance!D2:D{}, \"Tanzania\")".format(no_of_hosts))

    # Number of Switches that are Fully compliant
    worksheet.write_formula("B3",
                            "=COUNTIFS(Compliance!D2:D{}, \"Kenya\",Compliance!I2:I{}, \"100%\")".format(no_of_hosts,
                                                                                                          no_of_hosts))
    worksheet.write_formula("C3",
                            "=COUNTIFS(Compliance!D2:D{}, \"Uganda\",Compliance!I2:I{}, \"100%\")".format(no_of_hosts,
                                                                                                           no_of_hosts))
    worksheet.write_formula("D3",
                            "=COUNTIFS(Compliance!D2:D{}, \"Rwanda\",Compliance!I2:I{}, \"100%\")".format(no_of_hosts,
                                                                                                           no_of_hosts))
    worksheet.write_formula("E3",
                            "=COUNTIFS(Compliance!D2:D{}, \"Tanzania\",Compliance!I2:I{}, \"100%\")".format(
                                no_of_hosts,
                                no_of_hosts))

    # Number of switches that are partially compliant
    worksheet.write_formula("B4", "=(B2-B3)")
    worksheet.write_formula("C4", "=(C2-C3)")
    worksheet.write_formula("D4", "=(D2-D3)")
    worksheet.write_formula("E4", "=(E2-E3)")

    # Number of Switches that are Fully Non-compliant
    worksheet.write_formula("B5",
                            "=COUNTIFS(Compliance!D2:D{}, \"Kenya\",Compliance!I2:I{}, \"0%\")".format(no_of_hosts,
                                                                                                          no_of_hosts))
    worksheet.write_formula("C5",
                            "=COUNTIFS(Compliance!D2:D{}, \"Uganda\",Compliance!I2:I{}, \"0%\")".format(no_of_hosts,
                                                                                                           no_of_hosts))
    worksheet.write_formula("D5",
                            "=COUNTIFS(Compliance!D2:D{}, \"Rwanda\",Compliance!I2:I{}, \"0%\")".format(no_of_hosts,
                                                                                                           no_of_hosts))
    worksheet.write_formula("E5",
                            "=COUNTIFS(Compliance!D2:D{}, \"Tanzania\",Compliance!I2:I{}, \"0%\")".format(
                                no_of_hosts,
                                no_of_hosts))

    # Totals
    worksheet.write_formula("F2", "=SUM(B2:E2)")
    worksheet.write_formula("F3", "=SUM(B3:E3)")
    worksheet.write_formula("F4", "=SUM(B4:E4)")
    worksheet.write_formula("F5", "=SUM(B5:E5)")
    worksheet.write_formula("F11", "=SUM(B11:E11)")
    worksheet.write_formula("F12", "=SUM(B12:E12)")

    adjust_col_sizes(df, writer, "Compliance per Country")


def generate_sheet5(writer):
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = workbook.add_worksheet("Unscanned")

    yellow_fill_format = workbook.add_format({"bg_color": "#f0fc03", "bold": True})
    worksheet.write_string("A1", "IP Address", cell_format=yellow_fill_format)


if __name__ == '__main__':
    # generate filename to be expected
    today = date.today()
    # filename = today.strftime("%Y%m%d_audit_row_hw_sw.csv")
    filename = "Data/cisco.xlsx"

    # load data
    dataframe = pd.read_excel('Data/cisco.xlsx')

    # Add 'Compliance % column' to dataframe
    wr = pd.ExcelWriter("NAC analysis {}.xlsx".format(today.strftime('%d %m %Y')), engine="xlsxwriter")
    compliance = [(dataframe['TRUNK PORTS'][x] + dataframe['COMPLIANT PORTS'][x]) / dataframe['TOTAL PORTS'][x] for x in
                  range(dataframe.shape[0])]
    dataframe.insert(5, "Compliance %", compliance)

    generate_sheet1(wr, dataframe)
    generate_sheet2(wr, dataframe)
    generate_sheet3(wr, dataframe)
    generate_sheet4(wr, dataframe)
    generate_sheet5(wr)

    wr.close()
