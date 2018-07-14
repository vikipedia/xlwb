import numpy
from debug import debug, trace

def npv(*args):
    discount_rate = args[0]
    cashflow = args[1]
    return sum([float(x)*(1+discount_rate)**-(i+1) for (i,x) in enumerate(cashflow)])


def SUM(*args):
    def flatten(lists):
        l = []
        for item in lists:
            if isinstance(item, list):
                l.extend(item)
            else:
                l.append(item)
        return l
    return sum(flatten(args))


def ROUND(v, ndigits=None):
    if ndigits:
        return round(v, int(ndigits))
    else:
        return round(v)


def IRR(values):
    return numpy.irr(values)


def SEARCH(tok, text):
    return text.find(tok)

    
def IFERROR(v, alternate):
    return v if v!=-1 else alternate
        

functionsmap ={
    "*":lambda x,y:x*y,
    "/":lambda x,y:x/y,
    "+":lambda x,y:x+y,
    "-":lambda x,y:x-y,
    "<":lambda x,y:x<y,
    "<=":lambda x,y:x<=y,
    ">":lambda x,y:x>y,
    ">=":lambda x,y:x>=y,
    "SUM": SUM,
    "ROUND":ROUND,
    "IF":lambda t,v,o: v if t else o,
    "NPV":npv,
    "PMT":numpy.pmt,
    "IRR":numpy.irr,
    "SEARCH":SEARCH,
    "IFERROR":IFERROR
    }

for name, func in functionsmap.items():
    functionsmap[name] = trace(func)
