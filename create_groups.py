import xlrd as xl

fname, lname, cohort, attendance = 1, 2, 3, 4
freshman, sophomore = set(), set()
group_size, min_freshman_per_group = 8, 4
groups = []

file_path = 'Responses.xlsx'
wb = xl.open_workbook(file_path).sheet_by_index(0)

for i in range(1, wb.nrows):
    if wb.cell_value(i, attendance) == "Yes":
        if wb.cell_value(i, cohort) == "Freshman":
            freshman.add("%s %s" % (wb.cell_value(i, fname), wb.cell_value(i, lname)))
        elif wb.cell_value(i, cohort) == "Sophomore":
            sophomore.add("%s %s" % (wb.cell_value(i, fname), wb.cell_value(i, lname)))

num_groups = round((len(freshman) + len(sophomore)) / group_size)
extra = len(freshman) % num_groups

for i in range(num_groups):
    new_group = []
    for j in range(min_freshman_per_group):
        new_group.append(freshman.pop())
    if extra != 0:
        new_group.append(freshman.pop())
        extra -= 1
    groups.append(new_group)

for group in groups:
    print(group)
