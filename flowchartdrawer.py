import xlrd

ExcelFileName = 'example.xlsx'
workbook = xlrd.open_workbook(ExcelFileName)
worksheet1 = workbook.sheet_by_name("Sheet2")

num_rows1 = worksheet1.nrows
num_cols1 = worksheet1.ncols

result_data1 = []
for curr_col in range(0, num_cols1, 1):
    row_data = []

    for curr_row in range(0, num_rows1, 1):
        data = worksheet1.cell_value(curr_row, curr_col)  # Read the data in the current cell
        # print(data)
        row_data.append(data)

    result_data1.append(row_data)

print(result_data1)
M = "|".join(result_data1[0])
N = "|".join(result_data1[1])
print(M)
print(N)