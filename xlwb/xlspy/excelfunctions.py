import numpy
import math
import re
import string
import functools
import operator

def extract_column_row(cell):
    pattern = re.compile(r'\$?(?P<COL>[A-Z]+)\$?(?P<ROW>\d+)')
    m = pattern.match(cell)
    column, row = m.groups()
    return column, row

def columnchange(c, n):
    sheet, cell = c.split("!")
    column, row = extract_column_row(cell)
    return "!".join([sheet,indtocol(columnind(column)+n) + row])

def nextcolumn(c):
    return columnchange(c, 1)

def prevcolumn(c):
    return columnchange(c, -1)

def test_columnchange():
    assert columnchange("S!B3", 1) == "S!C3"
    assert columnchange("S!B3", -1) == "S!A3"

def extractdata(data, range_):
    return [[data[c] for c in row] for row in excelrange(range_)]

def extract_ranges(ranges):
    """
    >>> extract_ranges("S!A1:A3,S!B3")
    [S!A1,S!A2,S!A3,S!B3]
    """
    ranges_ = [excelrange(o) for o in ranges.strip().split(",")]
    o = [s.strip() for s in flatten(flatten(flatten(ranges_)))]
    return o

def test_extract_ranges():
    assert extract_ranges("S!A1:A3,S!B3") == ["S!A1","S!A2","S!A3","S!B3"]

def excelrange(range_):
    def cell(c,r):
        return "!".join([sheet, indtocol(c)+ str(r)])
    if ":" not in range_:
        return [[range_]]
    sheet, r = range_.split("!")
    start,end = r.split(":")
    sc,sr = extract_column_row(start)
    ec,er = extract_column_row(end)
    sc, sr = columnind(sc), int(sr)
    ec, er = columnind(ec), int(er)
    return [[ cell(j, i) for j in range(sc, ec+1)] for i in range(sr, er+1)]

def test_excelrange():
    assert excelrange("S!A1:A3") == [["S!A1"], ["S!A2"], ["S!A3"]]
    assert excelrange("S!A1:C1") == [["S!A1",  "S!B1", "S!C1"]]
    assert excelrange("S!A1:B2") == [["S!A1",  "S!B1"],["S!A2","S!B2"]]
    assert excelrange("S!A1") == [["S!A1"]]


def updateformula(formula, data):
    """
    >>> updateformula(("+", 1, ("CELL", "A1")), {"A1", "B1"})
    >>> ("+", 1, ("CELL", "B1"))
    """
    if isinstance(formula, tuple):
        if formula[0] == "CELL":
            return ("CELL", data.get(formula[1],formula[1]))
        else:
            return tuple(updateformula(item, data) for item in formula)
    elif isinstance(formula, list):
        return [updateformula(item, data) for item in formula]
    else:
        return formula

def test_update_formula():
    data = {"A1":"B1", "A2":"B2"}
    f1 = ("CELL" ,"A1")
    g1 = ("CELL" ,"B1")
    f2 = [("CELL", "A1"), ("CELL", "A2"),1, 2]
    g2 = [("CELL", "B1"), ("CELL", "B2"),1, 2]
    f3 = ("FUNC" , 1 , ("CELL", "A1"), ("+", 2, ("CELL", "A2")))
    g3 = ("FUNC" , 1 , ("CELL", "B1"), ("+", 2, ("CELL", "B2")))
    assert updateformula(f1, data) == g1
    assert updateformula(f2, data) == g2
    assert updateformula(f3, data) == g3

def copypaste(cellmap,source, target):
    """
    copy contents of source range and paste it to target range.
    There is assumption that source and target have same size and shape.
    """
    s = excelrange(source)
    t = excelrange(target)
    mapping = {}

    for sr, tr in zip(s,t):
        for sc, tc in zip(sr, tr):
            mapping[sc] = tc

    for k,v in mapping.items():
        f = updateformula(cellmap.get(k,None), mapping)
        cellmap[v] = f

def test_copypaste():
    cm = {"S!A1":10,"S!A2":("+", ("CELL", "S!A1"),2),"S!A3":("CELL", "S!C1"),
          "S!B1":20,"S!B2":0, "S!B3":0}
    copypaste(cm, "S!A1:A3", "S!B1:B3")
    assert cm["S!B1"] == 10
    assert cm["S!B2"] == ("+", ("CELL", "S!B1"),2)
    assert cm["S!B3"] == ("CELL","S!C1")
    cm = {"S!A1":10,"S!A2":("+", ("CELL", "S!A1"),2),"S!A3":("CELL", "S!C1"),
          "S!B1":20,"S!B2":0, "S!B3":0}
    copypaste(cm, "S!A1", "S!B3")
    assert cm['S!B3'] == 10

def print_table(table):
    vwrap()


def indtocol(num):
    """
    >>> indtocol(1)
    A
    """
    if num < 1:
        raise Exception("Number must be larger than 0: %s" % num)

    s = ''
    q = num
    while q > 0:
        (q,r) = divmod(q,26)
        if r == 0:
            q = q - 1
            r = 26
        s = string.ascii_uppercase[r-1] + s
    return s

def test_indtocol():
    assert indtocol(1) == "A"
    assert indtocol(27) == "AA"

def columnind(ch):
    """
    accepts column numbers in excels' columnname format
    and returns actual column number.
    >>> column("A")
    1
    >>> column("AA")
    27
    """
    n = len(ch)
    return sum((ord(c)-64)*26**(n-i-1) for i,c in enumerate(ch))


def test_columnind():
    assert columnind("A")==1
    assert columnind("AA")==27
    assert sorted(["A","AA","B","C","AAA"], key=columnind) == ["A","B","C","AA","AAA"]

def npv(*args):
    discount_rate = float_(args[0])
    cashflow = numeric(flatten(flatten(args[1])))
    try:
        return sum([float_(x)*(1+discount_rate)**-(i+1) for (i,x) in enumerate(cashflow)])
    except TypeError as t:
        print("NPV", [type(i) for i in args], args)

#def test_npv():
#    assert npv(10, range(10)) == numpy.npv(10, range(10))


def flatten(lists):
    """
    flattens the hirarchy in tuples and returns it as flattened tuple
    """
    l = []
    for item in lists:
        if isinstance(item, list):
            l.extend(item)
        else:
            l.append(item)
    return l


def float_(x):
    if not x : return 0
    if x == "-" : return 0
    return x

def numeric(l):
    return [i or 0 for i in l]


def SUM(*args):
    return math.fsum(numeric(flatten(flatten(args))))

def test_SUM():
    assert SUM([1,1,1]) == 3
    assert SUM(1, 1, 1, 2) == 5
    assert SUM(1) == 1
    assert SUM([1]) == 1
    assert SUM([[1,1],2]) == 4
    assert SUM([[1,1],[1,1,1]]) == 5
    assert SUM([[1]]) == 1


def regexspecial(s):
    return re.compile(s.replace(".","\.").replace("^","\^").replace("$","\$").replace("*",".*").replace("+","\+").replace("?","."))


def VLOOKUP(value, data, columnnum, range_look=True):
    col = [row[0] for row in data]
    if not range_look:
        if isinstance(value, str) and ("*" in value or "?" in value):
            p = regexspecial(value)
            index = [i for i,x in enumerate(col) if p.match(x)][0]
        else:
            index = col.index(value)
    else:
        diff = [value-(item or 0) for item in col]
        s = [d for d in sorted(diff) if d>=0]
        index = diff.index(s[0])
    return data[index][int(columnnum)-1]

def transpose(d):
    return [column(d, i) for i in range(len(d[0]))]

def column(d, n):
    return [row[n] for row in d]

def test_VLOOKUP():
    data = [[1,2,3,3],
            [2,3,4,3],
            [5,1,3,3]]
    value = 5
    assert VLOOKUP(5, data, 2, False) == 1
    assert VLOOKUP(5, data, 3, False) == 3
    assert VLOOKUP(2.1, data, 2) == 3

    data = [["hello","yellow","apple","ap.le"],
            [1,2,3,4],
            [5,6,7,8]]
    data = transpose(data)

    assert VLOOKUP("ye*", data, 2, False) == 2
    assert VLOOKUP("?ello", data, 2, False) == 1
    assert VLOOKUP("?ppl*", data, 2, False) == 3
    assert VLOOKUP("*p.l?", data, 2, False) == 4

def AVERAGE(*args):
    a = flatten(flatten(args))
    na = [item for item in a if isinstance(item,(float,int))]
    l = len(na)

    if not l: return 0
    return SUM(na)/l

def test_AVERAGE():
    assert AVERAGE([1,1,1,1,1])==1
    assert AVERAGE([1,1,0,0])==0.5
    assert AVERAGE([0,0,0,0])==0
    assert AVERAGE([1, 1, "-"])==1
    assert AVERAGE([1, 1, 1, None])==1

def COUNTIF(array, condition):
    array = flatten(flatten(array))
    if isinstance(condition, (int, float)):
        return array.count(condition)
    elif condition.startswith(">="):
        return len([x for x in array if x >= float(condition[2:])])
    elif condition.startswith("<="):
        return len([x for x in array if x <= float(condition[2:])])
    elif condition.startswith(">"):
        return len([x for x in array if x > float(condition[1:])])
    elif condition.startswith("<>"):
        try:
            cond = float(condition[2:])
        except:
            return len([x for x in array if str(x)!=condition[2:]])
        else:
            return len([x for x in array if abs(float(x)-float(condition[2:]))>0.0001])

    elif condition.startswith("<"):
        return len([x for x in array if x < float(condition[1:])])
    elif "*" in condition or "?" in condition:
        p = regexspecial(condition)
        return len([x for x in array if p.match(x)])
    else:
        return array.count(condition)

def test_COUNTIF():
    names = ["apple","orrange","apple", "pinapple"]
    value = [10,20,30,40,50,40,40,50]
    assert COUNTIF(names, "*e") == len(names)
    assert COUNTIF(names, "apple") == 2
    assert COUNTIF(value, "<40") == 3
    assert COUNTIF(value, "<=40") == 6
    assert COUNTIF(value, ">40") == 2
    assert COUNTIF(value, ">=40") == 5
    assert COUNTIF(value, 40) == 3
    assert COUNTIF(value, "<>40") == 5
    value = [10,20,30,40,50,40,40,50, 0.0, 0.0, 0.0]
    assert COUNTIF(value, "<>0") == 8
    value = [0.2]*30 + [0.0, 0.0, 0.0]
    assert COUNTIF(value, "<>0") == 30




def AND(*args):
    return functools.reduce(operator.and_, args, True)

def OR(*args):
    return functools.reduce(operator.or_, args, False)


def ROUND(v, ndigits=None):
    if ndigits:
        return round(v, int(ndigits))
    else:
        return round(v)


def IRR(values):
    values = numeric(flatten(flatten(values)))
    return numpy.irr(values)


def SEARCH(tok, text):
    return text.find(tok) + 1


def INDEX(data, row, col=None):
    if not col:
        data = flatten(data)
        return data[int(row)-1]
    else:
        return data[int(row)-1][int(col)-1]

def MATCH(value, array, match_type=1):
    array = flatten(array)
    if match_type == 0:
        return array.index(value)+1
    elif match_type == 1:
        diff = [value-item for item in array]
        s = [d for d in sorted(diff) if d>=0]
        return diff.index(s[0])+1
    elif match_type == -1:
        diff = [item-value for item in array]
        s = [d for d in sorted(diff) if d>=0]
        return diff.index(s[0])+1

def test_MATCH():
    data = [1,2,3,4,5,6]
    assert MATCH(3, data, 0) == 3
    assert MATCH(3.5, data, 1) == 3
    assert MATCH(3.5, data, -1) == 4


def OFFSET(ref, *args):
    pattern = re.compile(r"('?(?P<SHEET>[\w &\-\|]+)'?[\!\.])?(?P<RANGE>(?P<CELL>[A-Z]+\d+)(:[A-Z]+\d+)?)")
    ref = ref.replace("$","")
    m = pattern.match(ref)
    SHEET = m.group("SHEET")
    RANGE = m.group("RANGE")
    CELL = m.group("CELL")

    col, row = extract_column_row(CELL)

    rows, cols = args[:2]
    col = columnind(col) + int(cols)
    row = int(row) + int(rows)

    r = indtocol(col)+str(row)

    if len(args)==3:
        height = args[2]
    elif len(args)==4:
        height, width = args[2:]
        col1 = col + int(width)-1
        row1 = row + int(height)-1

        r =  ":".join([r, indtocol(col1)+str(row1)])
    if SHEET:
        return "!".join([SHEET, r])
    return r


def test_OFFSET():
    assert OFFSET("A4", 0, 0) == "A4"
    assert OFFSET("A4", -1, 0) == "A3"
    assert OFFSET("sheet!A4", 0, 0) == "sheet!A4"

    assert OFFSET("A4", 0, 1) == "B4"
    assert OFFSET("A2:B8", 1, 2) == "C3"
    assert OFFSET("A2:B8", 1, 2, 3, 3) == "C3:E5"
    assert OFFSET("sheet!A2:B8", 1, 2, 3, 3) == "sheet!C3:E5"


def SUMIFS(data, array, condition):
    """
    supports only one condition
    """
    d = numeric(flatten(flatten(data)))
    a = flatten(flatten(array))
    if isinstance(condition, (int, float)):
        return math.fsum(d[i] for i,x in enumerate(a) if x==condition)
    elif condition.startswith(">="):
        return math.fsum([d[i] for i,x in enumerate(a) if (x or 0) >= float(condition[2:])])
    elif condition.startswith("<="):
        return math.fsum([d[i] for i,x in enumerate(a) if (x or 0) <= float(condition[2:])])
    elif condition.startswith(">"):
        return math.fsum([d[i] for i,x in enumerate(a) if (x or 0) > float(condition[1:])])
    elif condition.startswith("<>"):
        return math.fsum([d[i] for i,x in enumerate(a) if x!=float(condition[2:])])
    elif condition.startswith("<"):
        return math.fsum([d[i] for i,x in enumerate(a) if (x or 0) < float(condition[1:])])
    elif "*" in condition or "?" in condition:
        p = condition.replace("*", ".*")
        p = re.compile(p.replace("?", "."))
        return math.fsum([d[i] for i,x in enumerate(a) if p.match(x)])
    else:
        return math.fsum(d[i] for i,x in enumerate(a) if x==condition)


def test_SUMIFS():
    d = [1,2,3,4,5,4,4,5]
    a = [10,20,30,40,50,40,40,50]

    assert SUMIFS(d, a, "<40") == 6
    assert SUMIFS(d, a, "<=40") == 18
    assert SUMIFS(d, a, ">40") == 10
    assert SUMIFS(d, a, ">=40") == 22
    assert SUMIFS(d, a, 40) == 12
    assert SUMIFS(d, a, "<>40") == 16



def SUMPRODUCT(array, *arrays):
    if not arrays:
        return math.fsum(array)
    b = (flatten(flatten(a)) for a in arrays)
    a = flatten(flatten(array))
    return math.fsum(functools.reduce(operator.mul,numeric(item)) for item in zip(a,*b))


def test_SUMPRODUCT():
    assert SUMPRODUCT([1,2,3,4], [1]*4) == 10
    assert SUMPRODUCT([1,2,3,4], [1]*4, [1]*4) == 10
    assert SUMPRODUCT([1,2,3,4]) == 10


def PMT(rate, nper, pv, fv=0, type_=0):
    t = "end" if type_ == 0 else "begin"
    return numpy.pmt(rate, nper, pv, fv, t)


def ROUNDUP(v, n=0):
    n = int(n)
    v = v * 10**n
    return math.ceil(v)/10**n

def IF(*args):
    if len(args)==3:
        return args[1] if args[0] else args[2]
    elif len(args)==2:
        return args[1] if args[0] else False
    else:
        return bool(args[0])

def MAX(*args):
    a = flatten(flatten(args))
    return max([x for x in a if isinstance(x, (int, float))])

def MIN(*args):
    a = flatten(flatten(args))
    return min([x for x in a if isinstance(x, (int, float))])


functionsmap ={
    "*":lambda x,y:float_(x) * float_(y),
    "/":lambda x,y:x/y,
    "+":lambda x,y:float_(x) + float_(y),
    "-":lambda x,y:float_(x) - float_(y),
    "&":lambda x,y:(x or "") + (y or ""),#string concatination
    "<":lambda x,y:float_(x)<float_(y),
    "<=":lambda x,y:float_(x)<=float_(y),
    ">":lambda x,y:float_(x)>float_(y),
    ">=":lambda x,y:float_(x)>=float_(y),
    "=":lambda x,y: x==y,
    "<>":lambda x,y: x!=y,
    "^":lambda x,y: float_(x) ** int(float_(y)),
    "CELL":lambda x: x,
    "SUM": SUM,
    "ROUND":ROUND,
    "IF":IF,
    "NPV":npv,
    "PMT":PMT,
    "IRR":IRR,
    "SEARCH":SEARCH,
    "AND":AND,
    "OR": OR,
    "AVERAGE":AVERAGE,
    "COUNTIF":COUNTIF,
    "INDEX": INDEX,
    "IPMT":numpy.ipmt,
    "ISNUMBER":lambda x: isinstance(x, (int, float)),
    "MATCH":MATCH,
    "MAX":MAX,
    "MIN":MIN,
    "MOD":lambda x,y: x%y,
    "OFFSET":OFFSET,
    "ROUNDUP":ROUNDUP,
    "SQRT":math.sqrt,
    "SUMIFS":SUMIFS,
    "SUMPRODUCT":SUMPRODUCT,
    "VLOOKUP":VLOOKUP
    }

for name, func in functionsmap.items():
    functionsmap[name] = func
