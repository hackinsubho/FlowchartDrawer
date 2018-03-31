import pydot
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
print(len(result_data1[1]))
M = "|".join(result_data1[0])
N = "|".join(result_data1[1])
callgraph = pydot.Dot(graph_type='digraph', fontname="Verdana")
callgraph.set_strict(1)
callgraph.set_label("Assignment Pass Due Scenarios")

cluster_foo=pydot.Cluster('Ids',label='')
cluster_foo.add_node(pydot.Node('foo', label="{"+M+"}|{"+N+"}", shape="record", orientation="180"))
callgraph.add_subgraph(cluster_foo)

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

cluster_graph=pydot.Cluster('graph', label='')

for i in range(len(result_data[1])):
    node[i] = pydot.Node(result_data[1][i], style="rounded, filled", shape="box", rotate="")
    cluster_graph.add_node(node[i])

j = 0
k = 1

for i in range(len(set(node))-1):
    cluster_graph.add_edge(pydot.Edge(node[j], node[k], dir="forward", arrowhead="normal", style=""))
    j += 1
    k += 1

callgraph.add_subgraph(cluster_graph)



callgraph.write_png("test2.png")