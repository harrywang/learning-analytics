from openpyxl import load_workbook
import gender_guesser.detector as gender

wb2 = load_workbook('sample-names.xlsx')
sheet = wb2['names']
count = 0  # total name processed
d = gender.Detector()
for cells in tuple(sheet.rows):
    if cells[0].value == 'Student':  # skip the heading
        continue
    full_name = cells[0].value.split()
    first_name = full_name[0]
    gender = d.get_gender(first_name)
    count += 1
    print first_name, gender
    cells[1].value = gender
wb2.save('sample-names-gender.xlsx')
