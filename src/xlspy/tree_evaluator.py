
from excelfunctions import functionsmap
from memoize import memoize
from debug import trace

class TreeError(TypeError):
    pass



def search(t, item):
    if not t: return False
    if isinstance(t, tuple):
        if item in t:
            return True
        else:
            for t_ in t:
                if search(t_, item):
                    return True
                else:
                    continue
            return False
    

def test_search():
    t = (1,(2,3,4,(5)))
    assert search(t, 1)
    assert search(t, 3)
    assert search(t, 5)
    assert search(t, 6) == False
    
def IFERROR(v, alternate, inputvalues):
    try:
        return evaluate(v, inputvalues)
    except Exception as e:
        return evaluate(alternate, inputvalues)


def IF(inp, *args):
    if len(args)==1:
        return evaluate(args[0], inp)
    elif len(args)==2:
        return evaluate(args[1], inp) if evaluate(args[0], inp) else False
    else:
        if evaluate(args[0], inp):
            return evaluate(args[1], inp) 
        else:   
            return evaluate(args[2], inp)
        
def debugfunc(name, node, inputvalues):
    if node[0]==name:
        for a in node[1:]:
            print(a, "=",evaluate(a, inputvalues))

    
def evaluate(node, inputvalues):

    if isinstance(node, tuple):
        #debugfunc("CELL", node, inputvalues)
        if node[0]=="IFERROR":
            return IFERROR(*node[1:], inputvalues)
        elif node[0]=="IF":
            return IF(inputvalues, *node[1:])

        function = functionsmap[node[0]]
        arguments = [evaluate(item, inputvalues) for item in node[1:]]
        for a in arguments:
            if isinstance(a, tuple):
                raise TreeError("Invalid arguments to {0} : {1}".format(node[0], a))
        return function(*arguments)
    elif isinstance(node, list):
        return [evaluate(item, inputvalues) for item in node]
    elif isinstance(node, str): # these are cellids or literal strings
        if "!" in node:
            default = 0
        else:
            default = node
        return inputvalues.get(node, default)#to handle invalid cells
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

