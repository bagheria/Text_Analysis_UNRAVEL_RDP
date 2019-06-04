import xlrd
import xlsxwriter

path_dictionary = "synonyms.xlsx"
path_data = "Datadump.xlsx"
path_out = "Output.xlsx"

# Reading data
workbook = xlrd.open_workbook(path_data)
sheet = workbook.sheet_by_index(0)
textData = [sheet.cell_value(row, 0) for row in range(sheet.nrows)]


# Reading dictionaries
workbook = xlrd.open_workbook(path_dictionary)
sheet = workbook.sheet_by_index(0)

Clinical_Variables = ["NYHA Class", "CCS Class", "ankle oedema", "ascites", "pulmonary rales", "3rd heart sound",
                      "chronic renal failure", "neuromuscular dystrophy", "arterial hypertension", "diabetes",
                      "dyslipidaemia", "smoking", "alcohol consumption", "Previous ventricular fibrillation",
                      "Previous syncope", "Family history of SCD", "Competitive sports", "Previous sustained VT",
                      "Previous non-sustained VT", "Familial DCM"]

# Clinical Examination
NYHA_Class_synonyms = []
CCS_Class_synonyms = []
ankle_oedema_synonyms = []
ascites_synonyms = []
pulmonary_rales_synonyms = []
T3rd_heart_sound_synonyms = []

for col in range(4, sheet.ncols):
    if sheet.cell_value(1, col) != '':
        NYHA_Class_synonyms.append(sheet.cell_value(1, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(3, col) != '':
        CCS_Class_synonyms.append(sheet.cell_value(3, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(13, col) != '':
        ankle_oedema_synonyms.append(sheet.cell_value(13, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(14, col) != '':
        ascites_synonyms.append(sheet.cell_value(14, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(15, col) != '':
        pulmonary_rales_synonyms.append(sheet.cell_value(15, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(16, col) != '':
        T3rd_heart_sound_synonyms.append(sheet.cell_value(16, col))

###############


# Comorbidities
chronic_renal_failure_synonyms = []
neuromuscular_dystrophy_synonyms = []
arterial_hypertension_synonyms = []
diabetes_synonyms = []
dyslipidaemia_synonyms = []
smoking_synonyms = []
alcohol_consumption_synonyms = []

for col in range(4, sheet.ncols):
    if sheet.cell_value(18, col) != '':
        chronic_renal_failure_synonyms.append(sheet.cell_value(18, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(19, col) != '':
        neuromuscular_dystrophy_synonyms.append(sheet.cell_value(19, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(20, col) != '':
        arterial_hypertension_synonyms.append(sheet.cell_value(20, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(21, col) != '':
        diabetes_synonyms.append(sheet.cell_value(21, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(23, col) != '':
        dyslipidaemia_synonyms.append(sheet.cell_value(23, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(24, col) != '':
        smoking_synonyms.append(sheet.cell_value(24, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(25, col) != '':
        alcohol_consumption_synonyms.append(sheet.cell_value(25, col))

################


# Medical History
Previous_ventricular_fibrillation_synonyms = []
Previous_syncope_synonyms = []
Family_history_of_SCD_synonyms = []
Competitive_sports_synonyms = []
Previous_sustained_VT_synonyms = []
Previous_non_sustained_VT_synonyms = []
Familial_DCM_synonyms = []

for col in range(4, sheet.ncols):
    if sheet.cell_value(30, col) != '':
        Previous_ventricular_fibrillation_synonyms.append(sheet.cell_value(30, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(31, col) != '':
        Previous_syncope_synonyms.append(sheet.cell_value(31, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(32, col) != '':
        Family_history_of_SCD_synonyms.append(sheet.cell_value(32, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(33, col) != '':
        Competitive_sports_synonyms.append(sheet.cell_value(33, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(34, col) != '':
        Previous_sustained_VT_synonyms.append(sheet.cell_value(34, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(35, col) != '':
        Previous_non_sustained_VT_synonyms.append(sheet.cell_value(35, col))

for col in range(4, sheet.ncols):
    if sheet.cell_value(36, col) != '':
        Familial_DCM_synonyms.append(sheet.cell_value(36, col))

###############


workbook1 = xlsxwriter.Workbook(path_out)
worksheet = workbook1.add_worksheet()

format = workbook1.add_format()
format.set_text_wrap()
format.set_align('center')
worksheet.set_column('A:A', 40)
worksheet.set_column('B:B', 40)
format.set_font_size(14)

worksheet.write(0, 0, "NUMBER", format)
worksheet.write(0, 1, "DECURSUS", format)
for k in range(Clinical_Variables.__len__()):
    worksheet.write(0, k + 2, Clinical_Variables[k], format)


format1 = format
format1.set_font_size(12)
format1.set_align('vcenter')


# Searching ....
i = 1
for row in range(textData.__len__()):
    worksheet.write(i, 0, i, format)  # only number
    worksheet.write(i, 1, textData[row], format)

    col = 2

    if any([x in str(textData[row]) for x in NYHA_Class_synonyms]):
        if ["klasse  1/4" in str(textData[row])] or ["klasse  I" in str(textData[row])]:
            worksheet.write(i, col, "NYHA class I", format1)
        elif ["klasse  2/4" in str(textData[row])] or ["klasse  II" in str(textData[row])]:
            worksheet.write(i, col, "NYHA class II", format1)
        elif ["klasse  3/4" in str(textData[row])] or ["klasse  III" in str(textData[row])]:
            worksheet.write(i, col, "NYHA class III", format1)
        elif ["klasse  4/4" in str(textData[row])] or ["klasse  IV" in str(textData[row])]:
            worksheet.write(i, col, "NYHA class IV", format1)
        else:
            worksheet.write(i, col, "YES!", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col = col + 1

    if any([x in str(textData[row]) for x in CCS_Class_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in ankle_oedema_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in ascites_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in pulmonary_rales_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in T3rd_heart_sound_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    #####

    if any([x in str(textData[row]) for x in chronic_renal_failure_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in neuromuscular_dystrophy_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in arterial_hypertension_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in diabetes_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in dyslipidaemia_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in smoking_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in alcohol_consumption_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    #####

    if any([x in str(textData[row]) for x in Previous_ventricular_fibrillation_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in Previous_syncope_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in Family_history_of_SCD_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in Competitive_sports_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in Previous_sustained_VT_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in Previous_non_sustained_VT_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1

    if any([x in str(textData[row]) for x in Familial_DCM_synonyms]):
        worksheet.write(i, col, "Yes", format1)
    else:
        worksheet.write(i, col, "No", format1)
    col=col + 1


    i = i + 1



workbook1.close()
