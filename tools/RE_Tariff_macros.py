from xlwb.xlspy.excelfunctions import indtocol, copypaste, excelrange, columnind, flatten
from xlwb.xlspy.excelexec import evaluate_cell, build_graph, update_graph
from xlwb.xlspy.excelexec import dfs_postorder_nodes


def copy_dependency(cellid, cm, graph):
    cells = dfs_postorder_nodes(graph, cellid)
    return {c:cm[c] for c in cells if c in cm}


def copy_formulas(src, dest):
    x = {c:src[c] for c in src if isinstance(src[c], tuple)}
    dest.update(x)


def UpdateWCapitalRequirement(cm, graph, w):

    for i in range(1, 36):
        TariffDiff = 100
        wcaps = "W Capital!{}15".format(indtocol(i+2))
        tariff = "Tariff!{}43".format(indtocol(i+2))
        cm_ = copy_dependency(tariff, cm , graph)
        g = build_graph(cm_)
        _cm = dict(cm_)

        AssumedTariff = evaluate_cell(wcaps, _cm, g, w)#this has only values

        while TariffDiff > 0.02:
            copy_formulas(cm_, _cm)
             # make use of formulas to compute, do not change
             # Do not replace formula with value

            Tariff = evaluate_cell(tariff, _cm, g, w)
            while isinstance(Tariff, tuple):
                Tariff = evaluate_cell(tariff, _cm, g, w)

            if not isinstance(Tariff, (float,int)):
                Tariff = AssumedTariff
            TariffDiff = abs(Tariff - AssumedTariff)
            AssumedTariff = round(Tariff, 2)

            _cm[wcaps] = AssumedTariff

        cm.update(_cm)


def gettechrow(cm, range_):
    """
    returns excel row as a list. handles missing cells
    """
    techs = []
    for cell in excelrange(range_)[0]:
        techs.append(cm[cell] if cell in cm else None)
    return techs

def RecallStoredInputs(cm, technology, state):
    techheaders = gettechrow(cm, "Inputs!E1:Z1")
    col = techheaders.index(technology) + 5

    row = 73
    s = "Inputs!{col}{startrow}:{endcol}{row}".format(col=indtocol(col),
                                                      startrow=2,
                                                      endcol=indtocol(col+1),
                                                      row=row)

    target = "Inputs&Summary!D32:E103"
    copypaste(cm, s, target)

    #APCC
    target = "Inputs&Summary!D107:K107"
    states = [cm[cell] for cell in flatten(excelrange("Inputs-REC!C4:C22"))]
    row = states.index(state) + 4
    source = "Inputs-REC!D{row}:K{row}".format(row=row)
    copypaste(cm,source, target)

    if "Off-grid" in technology:
        return

    #FiTIP
    target = "Inputs&Summary!D100"
    techs = [cm[cell] for cell in excelrange("Inputs-REC!D36:K36")[0]]
    states = [cm[cell] for cell in flatten(excelrange("Inputs-REC!C37:C55"))]
    col = techs.index(technology) + 4
    row = states.index(state) + 37

    source ="Inputs-REC!{col}{row}".format(col=indtocol(col), row=row)
    if source not in cm or cm[source]==None:
        source = "Inputs-REC!{col}37".format(col=indtocol(col))
    f = lambda x: print(x, cm[x])
    copypaste(cm, source, target)


    if "Solar" in technology:
        source = "Inputs-REC!E30:L32"
    else:
        source = "Inputs-REC!E27:L29"
    target = "Inputs&Summary!D109:K111"
    copypaste(cm, source, target)


def HandleTechOrStateChange(data,w=None, **kwargs):
    technstate = ["Inputs&Summary!D14", "Inputs&Summary!D15", "Inputs&Summary!D16", "Inputs&Summary!D17","Inputs&Summary!D18"]
    for item in technstate:
        data[item] = kwargs[item]
    technology = data['Inputs&Summary!D16']
    state = data['Inputs&Summary!D15']
    RecallStoredInputs(data, technology, state)
    advanceditems = [item for item in kwargs if item not in technstate]
    for item in advanceditems:
        data[item] = kwargs[item]
    if technology not in ["Biogass", "Bagasse", "Biomass Gasifier", "Biomass Rankine Cycle"]:
        data['Inputs&Summary!D94'] = 0
    graph = build_graph(data)
    UpdateWCapitalRequirement(data, graph, w)
