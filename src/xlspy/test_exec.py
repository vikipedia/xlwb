from openpyxl import load_workbook
import excelexec
import excelfunctions
import excelparser
import pytest
from excelexec import update_cellmap
from excelfunctions import excelrange, flatten
from math import isnan

class ComparisonFailure(Exception):
    pass

def examine(cellid, cellmap, w):
    """
    A function to do testing. it allows examing computed
    value of given cell with precalculted excel sheet.
    this is most unclean function as it goes thorugh all
    confilting behaviours of excel about how it treats empty cell or
    how it treats nonexistent references!
    """
    
    ignorelist = []
    examinlist = ['Fuel!C9', "Inputs&Summary!D94", "Inputs&Summary!D35","Inputs&Summary!D34"]
    formula = cellmap.get(cellid, None)
    update_cellmap([cellid], cellmap)
    
    if cellid in ignorelist:
        return

    if cellid in examinlist:
        print(cellid, formula, cellmap[cellid])

    if not formula:
        #print("ignoring empty cell",cellid)
        return
    
    sheet, cell = cellid.split("!")
    wvalue = w[sheet][cell].value
    value = cellmap.get(cellid, None)
    #if wvalue == "#REF!":
        #print(cellid, value, wvalue)

    if isinstance(value, tuple):
        #print("ignoring uncomputed cell", cellid)
        return
    error = False
    
    if isinstance(wvalue, str):
        if value!=wvalue:
            if wvalue=="#REF!":
                if value==None or int(value)==0:
                    error = False
            elif wvalue == "#VALUE!":
                if value==None or isnan(value) or  int(value)==0:
                    error = False
            else:
                error = True
    elif wvalue==None:
        if value:
            error = True
    else:
        if value==None:
            if int(wvalue)==0:
                error = False
        elif abs(wvalue-value)>=0.001:
            error = True
    if error:
        print(cellid, formula, value,"*", wvalue)
        #print("***", cellid,  value, wvalue)
        raise ComparisonFailure(cellid + " incorrect results")


def update_cellmap_(cells, cellmap, w):
    for cel in cells:
        examine(cel, cellmap, w) 


        
def test_end_to_end(monkeypatch):
    """
    make sure that excel file is saved with desired inputs. this
    test makes use of cached data saved by excel to check cell by cell 
    comparison on computed cells
    """
    filename = "RE_Tariff_and_Financial_Analysis_Tool_v2.2-unprotected.xlsm"
    excelparser.main(filename)

    exceldata = excelparser.output_extn(filename)
    inputs = "inputs.yaml"
    
    w = load_workbook(filename=filename, data_only=True)
    monkeypatch.setattr(excelexec, "update_cellmap", update_cellmap_)
    excelexec.main(exceldata, inputs, w)
    
