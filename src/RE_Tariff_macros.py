from excelfunctions import indtocol, copypaste, excelrange, columnind, flatten
from excelexec import evaluate_cell

def UpdateWCapitalRequirement(cm, graph, w):

    for i in range(1, 36):
        TarrifDiff = 100

        while TarrifDiff > 0.02:
            wcaps = "W Capital!{}15".format(indtocol(i+2))
            AssumedTarrif = evaluate_cell(wcaps, cm, graph, w)
            tarrif = "Tariff!{}43".format(indtocol(i+2))
            Tarrif = evaluate_cell(tarrif, cm, graph, w)
            if not isinstance(Tarrif, (float,int)):
                Tarrif = AssumedTarrif
            TarrifDiff = abs(Tarrif - AssumedTarrif)
            AssumedTarrif = round(Tarrif, 2)
            cm[wcaps] = AssumedTarrif

    
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

    
def HandleTechOrStateChange(data, graph,w=None, technology="Solar PV", state="CERC"):
    s = "Inputs&Summary!D15"
    t = "Inputs&Summary!D16"
    if True or data[s] != state or data[t]!=technology:
        data[s] = state
        data[t] = technology
        RecallStoredInputs(data, technology, state)
        if technology not in ["Biogass", "Bagasse", "Biomass Gasifier", "Biomass Rankine Cycle"]:
            data['Inputs&Summary!D94'] = 0
        
    UpdateWCapitalRequirement(data, graph, w)



