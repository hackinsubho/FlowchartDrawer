#!/usr/bin/env python
import time
import pydot
import xlrd
import sys
import os
cwd = os.getcwd()
os.environ["PATH"] += os.pathsep + 'C:\\Users\\subhodip.ghosh\\Documents\\Graphviz\\bin'

callgraph = pydot.Dot(graph_type='digraph', fontname="Verdana")
callgraph.set_strict(1)
callgraph.set_label(sys.argv[2])

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

test_Ids = "|".join(result_data1[0])
jira_Ids = "|".join(result_data1[1])
test_case_name = "|".join(result_data1[2])

cluster_table=pydot.Cluster('Ids',label='')
cluster_table.add_node(pydot.Node('foo', label="{"+test_Ids+"}|{"+jira_Ids+"}|{"+test_case_name+"}", shape="record", orientation="180"))
callgraph.add_subgraph(cluster_table)

ExcelFileName = sys.argv[1]
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
del result_data[0]
node = result_data

for i in range(num_rows-1):
    node[i] = [x for x in node[i] if x]


cluster_graph = pydot.Cluster('graph', label='')
l = []
for j in range(num_rows-1):
    l.append(result_data[j][0])
    for i in range(len(set(node[j]))):
        if (i != 0):
            set(l)
            cluster_graph.add_node(pydot.Node(result_data[j][i], style="rounded, filled", shape="box", label=result_data[j][i]+"\n Test_ID: "+', '.join(l)))

#print(set(l))
for m in range(num_rows-1):
    for i in range(len(set(node[m]))-1):
        if i != 0:
            cluster_graph.add_edge(pydot.Edge(node[m][i], node[m][i+1], dir="forward", arrowhead="normal", style=""))


callgraph.add_subgraph(cluster_graph)
print("Preparing Flowchart", end="")
animation = " ..................\n"
for l in animation:
    sys.stdout.write(l)
    sys.stdout.flush()
    time.sleep(0.2)
print("\nFlowchart created in file "+sys.argv[3]+".png")
callgraph.write_png(sys.argv[3]+".png")
#callgraph.write_png("grh.png")