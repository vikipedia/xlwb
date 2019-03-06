from xlwb.xlspy.excelfunctions import extractdata, flatten
from xlwb.xlspy import excelexec
import json, yaml, re

def column(data, c):
    return tuple([row[c] for row in data])

def transpose(data):

    return [column(data, i) for i in range(len(data[0]))]

def to_py(s):
    return s.replace("!","__").replace(" ","___").replace("&","____")

def evaluate_cond(cond, celldata):
    p = re.compile('[\w &-]+[\!\.][A-Z]+\d+')
    cells = p.findall(cond)
    if cells:
        excelexec.compute_range(celldata, ",".join(cells))
    mapping = {c:to_py(c) for c in cells}
    for k, v in mapping.items():
        cond = cond.replace(k, v)
    return eval(cond, {mapping[c]:celldata[c] for c in cells})

def process_chartdata(celldata, inputs_conf):
    c = []
    for charts_conf in inputs_conf['charts']:
        if 'cond' in charts_conf:
            if not evaluate_cond(charts_conf['cond'], celldata):
                continue # skip this chart if cond is false
            else:
                pass
        columns, data = linechart_(celldata, charts_conf)
        d = {k:charts_conf[k] for k in ['types','id']}
        d['columns'] = columns
        d['data'] = data
        c.append(d)
    return c


def linechart_(celldata, conf):
    excelexec.compute_range(celldata, conf['X'])
    excelexec.compute_range(celldata, conf['names'])
    excelexec.compute_range(celldata, conf['series'])

    X = flatten(extractdata(celldata, conf['X']))
    name = conf['xname']
    names = flatten(extractdata(celldata, conf['names']))
    series = extractdata(celldata, conf['series'])

    category = conf['xcategory']
    rowwise = conf['rowwise']
    return linechart(celldata, name, X, names, series, category, rowwise)

def linechart(celldata, name, X, names, series, category=True, rowwise=True):
    """
    name => name of categories (for example Year)
    X => range for x axis (range that contains years)
    names => name for ever series
    series => series data to be plotted
    rowwise => take series as rows else column
    """
    if category:
        columns = [{'type':'string', 'label':name}]
        X = [str(int(item)) if isinstance(item, float) else str(item).strip() for item in X]
    else:
        columns = [{'type':'number', 'label':name}]
    columns = columns + [{'type':'number', 'label':s} for s in names if s]
    X = [item for item in X if item]

    data = [ v for v in series]
    if rowwise:
        data = [X] + data
        data_t = list(zip(*data))
        data_t = [[s for s in c if str(s).strip()] for c in data_t]
        return columns, data_t
    else:
        return columns, [ [x] + data[i] for i, x in enumerate(X)]
