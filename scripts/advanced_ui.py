"""
a script to automaticaly generate input conf based on cells given
"""
from xlwb.xlspy import excelfunctions
from xlwb.xlspy import excelexec
from xlwb.xlspy.excelfunctions import prevcolumn, nextcolumn
import yaml, sys
from openpyxl import load_workbook
from collections import OrderedDict
import click

def excelcell(w, c):
    s, cell = c.split("!")
    return w[s][cell]

@click.command()
@click.option('--conf', help='Existing conf file')
@click.option('--excel', help='Excel File')
def advanced_input_cells(conf, excel):
    '''
    This utility helps in creating advanced inputs.
    it assumes that you have handcrafted the basic yaml
    file. It modifies existing file and advanced_inputs
    section to it.
    '''
    w = load_workbook(excel)
    sections = ["Off-grid Parameters","Grid Extension Parameters", "Power Genaration",
                "Loan Details", "Depreciation", "Tax",
                "ROE/Doscount Rate","Fuel","Clean Eneregy Benifits"]
                #"REC Inputs-2013", "REC Inputs-2018","REC Inputs-2023","REC Inputs-2028"]
    sheet = "Inputs&Summary"
    ranges = ["D24:D27","D29:D31","D34:D44","D47:D61",
                "D64:D73","D76:D82","D85:D90","D93:D97",
                "D100:D103"]
    exceptions =["D36", "D38", "D48", "D49", "D69"]
    range_list = ["!".join([sheet,item]) for item in ranges]
    e = ["!".join([sheet, item]) for item in exceptions]

    uiconf = OrderedDict()
    for section, range_ in zip(sections, range_list):
        cells = excelfunctions.extract_ranges(range_)
        uiconf[section] = [generate_ui_data(c, excelcell(w, c)) for c in cells if c not in e]

    with open(conf) as f:
        d = yaml.load(f.read())
    d['advanced_inputs'] = uiconf
    print(yaml.dump(d))

def get_type(format_):
    if format_.count("0") == 1:
        return "int"
    else:
        return "float"

def generate_ui_data(c, cell):
    t = get_type(cell.number_format)
    return { "id": c,
        "description": prevcolumn(c),
        "name": c,
        "type": t,
        "ui": t,
        "value": c,
        "default": nextcolumn(c),
        "unit": nextcolumn(nextcolumn(c)),
        "format":cell.number_format,
        }

if __name__ == "__main__":
    if len(sys.argv)==1:
        print("For help \n$python advanced_ui.py --help")
    else:
        advanced_input_cells()
