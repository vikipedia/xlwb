from tree_evaluator import TreeError, evaluate
import sys
import networkx as nx
from networkx.algorithms.traversal.depth_first_search import dfs_postorder_nodes
from networkx.algorithms.cycles import find_cycle
from networkx.exception import NetworkXNoCycle
import collections
from debug import debug
import argparse
import yaml, pickle
from excelfunctions import excelrange
import importlib
from math import isnan

def cells(lispexpression):
    """
    This function finds cells on which this lisp expression depends.
    e = ("*" , ("CELL", "A1"), ("+" , ("CELL", "A2"), ("CELL", "A3")))
    expression e depends on A1, A2 and A3.
    """
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
    add_node = G.add_node
    add_edge = G.add_edge
    add_node(parent)
    for c in cells(lispexpression):
        add_node(c)
        add_edge(parent, c)
    
def update_cellmap(cells, cellmap, w=None):
    """
    evaluate every cell from list of cells and update cellmap
    w is dummy argument for testing purpose
    """
    count = 0

    for cellid in cells:
        c = cellmap.get(cellid,None)
        if isinstance(c, (tuple,list)):
            try:
                v = evaluate(c, cellmap)
                cellmap[cellid] = v
            except TreeError as t:
                count += 1
            except ZeroDivisionError as z:
                #print(cellid, z, cellmap[cellid])
                cellmap[cellid] = None
            except Exception as ex:
                print(cellid, c)
                raise ex
        elif isinstance(c,(int,float,str)):
            pass
        else:
            pass
            #print(cellid, c)
            
    return count
 
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


def find_cycle_(graph, cellid):
    try:
        cycle = [c[0] for c in find_cycle(graph, cellid)]
    except NetworkXNoCycle as nc:
        cycle = []
    return cycle

def evaluate_cell(cellid, cellmap, graph, w=None):
    if not isinstance(cellmap[cellid], tuple):
        return cellmap[cellid]

    cells =  list(dfs_postorder_nodes(graph, cellid))
    cycle = find_cycle_(graph, cellid)

    count = 5
    while count>0:
        index = min([cells.index(c) for c in cycle] or [0])
        update_cellmap(cells[:index], cellmap, w)
        for c in cycle:
            update_cellmap(reversed(cycle), cellmap, w)
        update_cellmap(cells[index:], cellmap, w)
        cycle = find_cycle_(graph, cellid)
        if not cycle:
            break
        if sum([1 for c in cycle if isinstance(cellmap[c], tuple)])==0:
            break
                
        count -= 1
    
    return cellmap[cellid]

def test_exec_excel():
    from excelparser import create_cellmap
    filename = "sample.xlsx"
    accuracy = 0.0001
    cellmap = create_cellmap(filename, {})
    graph = build_graph(cellmap)
    assert abs(evaluate_cell("Sheet1!C5", cellmap, graph) - 1.417)<=accuracy
    assert abs(evaluate_cell("Sheet1!C2", cellmap, graph) - 0.96)<=accuracy
    assert abs(evaluate_cell( "Sheet1!C3", cellmap, graph) - 1.25)<=accuracy
    assert abs(evaluate_cell("Sheet1!C4", cellmap, graph) - 3.34)<=accuracy
    assert evaluate_cell("Sheet1!C6", cellmap, graph) == 1    

def handle_macro(cm, inputs, graph, w=None):
    """
    w is precalcuted sheet by excel for testing purpose only
    """
    if "macro" not in inputs:
        return
    input_cells = inputs['input_cells']
    for k, v in input_cells.items():
        print(k, v)
    print("="*20)
    macro = inputs['macro']
    module = importlib.import_module(macro['module'])
    func = getattr(module, macro["function"])
    args = {}
    for k,v in macro['args'].items():
        args[k] = input_cells[v]
    func(cm, graph, w, **args)

def parse_args():
    parser = argparse.ArgumentParser("Excelsheet execution utility")
    parser.add_argument("exceldata",
                        type=str,
                        help="excel data generated from excelparser.py")

    parser.add_argument("inputs",
                        type=str,
                        help="inputs filename, inputs file should be in yaml format")
    return parser.parse_args()
    


def main(exceldata, inputs, w=None):
    with open(exceldata, "rb") as e:
        cellmap = pickle.load(e)
    with open(inputs) as inp:
        inputs = yaml.load(inp)

    graph = build_graph(cellmap)
    handle_macro(cellmap, inputs, graph, w)
    outputs = excelrange(inputs['output'])

    for row in outputs:
        for o in row:
            print(o,evaluate_cell(o, cellmap, graph, w))
        

if __name__ == "__main__":
    args = parse_args()
    main(args.exceldata, args.inputs)
