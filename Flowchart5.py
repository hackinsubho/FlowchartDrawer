import xlrd
import pydot

graph = pydot.Dot(graph_type='digraph')
graph.set_strict(1)
ExcelFileName = 'example.xlsx'
workbook = xlrd.open_workbook(ExcelFileName)
worksheet = workbook.sheet_by_name("Sheet1")

num_rows = worksheet.nrows
num_cols = worksheet.ncols

result_data = []
for curr_row in range(0, num_rows, 1):
    row_data = []

    for curr_col in range(0, num_cols, 1):
        data = worksheet.cell_value(curr_row, curr_col)  # Read the data in the current cell
        # print(data)
        row_data.append(data)

    result_data.append(row_data)


node = result_data[1]

for i in range(len(result_data[1])):
    node[i] = pydot.Node(result_data[1][i], style="rounded, filled", shape="box")
    graph.add_node(node[i])

j = 0
k = 1

for i in range(len(set(node))-1):
    graph.add_edge(pydot.Edge(node[j], node[k], dir="forward", arrowhead="normal", style=""))
    j += 1
    k += 1


graph.write_png('example3_graph.png')
