
from excelfunctions import functionsmap


def SUM(*args):
    flattned = []
    for item in args:
        if isinstance(item, list):
            flattned.extend(item)
        else:
            flattned.append(item)
    return sum(flattned)


operators={
    "*":lambda x,y:x*y,
    "/":lambda x,y:x/y,
    "+":lambda x,y:x+y,
    "-":lambda x,y:x-y,
    "SUM":SUM
    }


def evaluate(node, inputvalues):
    if isinstance(node, tuple):
        function = functionsmap[node[0]]
        arguments = [evaluate(item, inputvalues) for item in node[1:]]
        return function(*arguments)
    elif isinstance(node, list):
        return [evaluate(item, inputvalues) for item in node]
    elif isinstance(node, str):
        return inputvalues[node]
    else:
        return node


def test_evaluate():
    func1 = lambda x: x**2 + 2*x
    func2 = lambda x: x**3 + x**2 + x + 1
        
    accuracy = 0.00001
    vals = {'Sheet1!A4':0.2, 'Sheet1!A6':0.4, "Sheet1!A5":0.3}
    
    expr = ('+', ('*', 'Sheet1!A6', 'Sheet1!A6'), ('*', 'Sheet1!A6', 2.0))
    assert abs(evaluate(expr, vals) - func1(0.4)) <= accuracy
    expr = ('+', ('*', 'Sheet1!A4', 'Sheet1!A4'), ('*', 'Sheet1!A4', 2.0))
    assert abs(evaluate(expr, vals) - func1(0.2)) <= accuracy
    expr = ('+', ('+', ('*', ('*', 'Sheet1!A5', 'Sheet1!A5'), 'Sheet1!A5'), ('+', ('*', 'Sheet1!A5', 'Sheet1!A5'), 'Sheet1!A5')), 1.0)
    assert abs(evaluate(expr, vals)- func2(0.3))<= accuracy

def test_sum():
    vals = {"Sheet1!A2":1,"Sheet1!A3":1,"Sheet1!A4":1,"Sheet1!A5":1,"Sheet1!A6":1}
    expr = ('SUM', ['Sheet1!A2', 'Sheet1!A3', 'Sheet1!A4', 'Sheet1!A5', 'Sheet1!A6'])
    assert evaluate(expr, vals) == 5

