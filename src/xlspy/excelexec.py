import excelparser as xlp
from openpyxl import load_workbook
import tree_evaluator
import sys
import yaml
import networkx as nx
from networkx.algorithms.traversal.depth_first_search import dfs_postorder_nodes
import collections
from debug import debug



def create_cellmap(filename, inputs):
    w = load_workbook(filename=filename)
    cellmap = {}
    #graph = nx.DiGraph()
    
    et = xlp.ExpressionTreeBuilder(w)
    
    
    #put named ranges in cellmap
    for r in w.get_named_ranges():
        try:
            e = et.parse("="+r.attr_text)
            cellmap[r.name] = e
            #update_graph(graph, r.name, e) 
        except Exception as e:
            print("Skipping ", r.name , e)

    #put every cell in cellmap
    for i,name in enumerate(w.get_sheet_names()):
        w.active = i
        sheet = w[name]
        for col in sheet:
            for c in col:
                cellid = "!".join([name, c.coordinate])
                if c.value==None:
                    pass
                else:
                    if c.data_type == c.TYPE_FORMULA:
                        e = et.parse(c.value)
                        cellmap[cellid] = e
                        #update_graph(graph,cellid, e)
                    else:
                        cellmap[cellid] = c.value
                        
    cellmap.update(inputs)
    return cellmap


def cells(lispexpression):
    if isinstance(lispexpression, tuple):
        if lispexpression[0]=="CELL":
            yield lispexpression[1]
        else:
            for e in lispexpression[1:]:
                yield from cells(e)
    elif isinstance(lispexpression, list):
        for item in lispexpression:
            yield from cells(item)

def test_cells():
    e = ("*" , ("CELL", "A1"), ("+" , ("CELL", "A2"), ("CELL", "A3")))
    assert list(cells(e)) == ["A1", "A2", "A3"]
    e = ("*", ("SUM", [[("CELL","A1")],[("CELL","A2")]]), ("CELL", "A3"))
    assert list(cells(e)) == ["A1", "A2", "A3"]
    
def update_graph(G, parent, lispexpression):
    G.add_node(parent)
    for c in cells(lispexpression):
        G.add_node(c)
        G.add_edge(parent, c)
    
def update_cellmap(cells, cellmap):
    """
    evaluate every cell in cells and update cellmap
    """
    for cellid in cells:
        c = cellmap.get(cellid,None)
        if isinstance(c, (tuple,list)):
            try:
                v = tree_evaluator.evaluate(c, cellmap)
                cellmap[cellid] = v
            except tree_evaluator.TreeError as t:
                print(cellid, t)
                raise t
            except ZeroDivisionError as z:
                #print(cellid, z, cellmap[cellid])
                cellmap[cellid] = 0
            except Exception as ex:
                print(cellid, c)
                raise ex
 
def print_dict(d):
    for k in sorted(d.keys()):
        print(k, d[k])

def build_graph(data):
    g = nx.DiGraph()
    for k,v in data.items():
        if isinstance(v, tuple):
            update_graph(g, k, v)
    return g


def graphdata():
    cellmap = {"A1":1,"A2":2,
               "B1":("+", ("+" ,1, ("CELL", "A1")),("CELL", "A2")),
               "C1":("+" ,("CELL","B1"), 1),
               "D1":("+", ("CELL","C1"), 1),
               "E1":("SUM", [[("*", ("+" , ("CELL","C1"),("CELL", "D1")),("+" , ("CELL","C1"),("CELL", "D1")))], ["A1"],["B1"]])
    }
    return cellmap
    
def test_build_graph():
    cellmap = graphdata()
    g = build_graph(cellmap)
    assert sorted(g.nodes()) == sorted(["A1","A2","B1","C1","D1","E1"])
    assert set(g.edges()) == {('C1', 'B1'), ('B1', 'A1'), ('B1', 'A2'), ('D1', 'C1'),("E1","D1"),("E1","C1")}
    assert set(g.successors("B1")) == {"A1","A2"}
    assert set(g.successors("C1")) == {"B1"}
    assert set(g.successors("E1")) == {"C1","D1"}    
    
def evaluate_cell(cellid, cellmap, graph):
    cells =  dfs_postorder_nodes(graph, cellid)
    update_cellmap(cells, cellmap)
    return cellmap[cellid]


def test_exec_excel():
    filename = "sample.xlsx"
    accuracy = 0.0001
    cellmap = create_cellmap(filename, {})
    g = build_graph(cellmap)
    assert abs(evaluate_cell("Sheet1!C5", cellmap, g) - 1.417)<=accuracy
    assert abs(evaluate_cell("Sheet1!C2", cellmap, g) - 0.96)<=accuracy
    assert abs(evaluate_cell( "Sheet1!C3", cellmap, g) - 1.25)<=accuracy
    assert abs(evaluate_cell("Sheet1!C4", cellmap, g) - 3.34)<=accuracy
    assert evaluate_cell("Sheet1!C6", cellmap, g) == 1


if __name__ == "__main__":
    #sys.setrecursionlimit(15000)
    filename = sys.argv[1]
    cellid = sys.argv[2]
    inputs = {'Inputs&Summary!D15':"Andhra Pradesh"}
    outputs = [c+str(i) for i in range(8, 3, -1) for c in ["J","K","L","M"]]
    outputs = reversed(["Inputs&Summary!"+c for c in outputs])
    cellmap = create_cellmap(filename, inputs)


    #cellmap = yaml.load(open("cellmap.yaml"))
    g = build_graph(cellmap)

    for o in outputs:
        print(evaluate_cell(o, cellmap, g))
        
    """
    with open("cellmap2.yaml", "w") as f:
        f.write(yaml.dump(cellmap))
    """
    
