from openpyxl import load_workbook

def examine(cellid, cellmap, graph, w):
    formula = cellmap[cellid]
    update_cellmap([cellid], cellmap)
    sheet, cell = w.split("!")
    value = w[sheet][cell].value
    error = False
    if isinstance(value, str):
        if cellmap[cellid]==value:
            error = True
    else:
        if abs(cellmap[cellid]-value)>=0.001:
            error = True
    if error:
            print(cellid, formula, cellmap[cellid], value)


def main():
    w = load_workbook(filename="/home/vikrant/Documents/prayas/RE_Tariff_and_Financial_Analysis_Tool_v2.2-unprotected.xlsm", data_only=True)
    cellid = "Inputs&Summary!K4"
    cells =  list(dfs_postorder_nodes(graph, cellid))
    cycle = find_cycle_(graph, cellid)

    count = 5
    while count>0:
        index = min([cells.index(c) for c in cycle] or [0])
        update_cellmap(cells[:index], cellmap)
        for c in cycle:
            update_cellmap(reversed(cycle), cellmap)
        update_cellmap(cells[index:], cellmap)
        #graph = build_graph(cellmap)
        cycle = find_cycle_(graph, cellid)
        if not cycle:
            break
        if sum([1 for c in cycle if isinstance(cellmap[c], tuple)])==0:
            break
                
        count -= 1

