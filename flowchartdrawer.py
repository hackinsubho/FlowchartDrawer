#!/usr/bin/env python
import pydot
import xlrd
import sys
import os
cwd = os.getcwd()

#ExcelFileName = 'example.xlsx'
ExcelFileName = sys.argv[1]
#ExcelFileName = ExcelFileName[16:-2]
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

node = []
node = result_data

for i in range(num_rows):
    node[i] = [x for x in node[i] if x]


cluster_graph = pydot.Cluster('graph', label='')

for j in range(num_rows):
    for i in range(len(set(node[j]))):
        if i != 0:
            cluster_graph.add_node(pydot.Node(result_data[j][i], style="rounded, filled", shape="box", label=result_data[j][i]+"\n Test_ID: "+result_data[j][0]))


for m in range(num_rows):
    for i in range(len(set(node[m]))-1):
        if i != 0:
            cluster_graph.add_edge(pydot.Edge(node[m][i], node[m][i+1], dir="forward", arrowhead="normal", style=""))


callgraph.add_subgraph(cluster_graph)

callgraph.write_png(sys.argv[2]+".png")