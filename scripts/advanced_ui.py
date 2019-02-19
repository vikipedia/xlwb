"""
a script to automaticaly generate input conf based on cells given
"""
from xlwb.xlspy import excelfunctions
from xlwb.xlspy import excelexec
from xlwb.xlspy.excelfunctions import prevcolumn, nextcolumn
import yaml, sys
from openpyxl import load_workbook
from collections import OrderedDict

def excelcell(w, c):
    s, cell = c.split("!")
    return w[s][cell]

def advanced_input_cells(existing_conf, excelfile):
    w = load_workbook(excelfile)
    sections = ["Off-grid Parameters","Grid Extension Parameters", "Power Genaration",
                "Loan Details", "Depreciation", "Tax",
                "ROE/Doscount Rate","Fuel","Clean Eneregy Benifits"]
                #"REC Inputs-2013", "REC Inputs-2018","REC Inputs-2023","REC Inputs-2028"]
    sheet = "Inputs&Summary"
    ranges = ["D24:D27","D29:D31","D34:D44","D47:D61",
                "D64:D73","D76:D82","D85:D90","D93:D97",
                "D100:D103"]
    range_list = ["!".join([sheet,item]) for item in ranges]

    uiconf = OrderedDict()
    for section, range_ in zip(sections, range_list):
        cells = excelfunctions.extract_ranges(range_)
        uiconf[section] = [generate_ui_data(c, excelcell(w, c)) for c in cells]

    with open(existing_conf) as f:
        d = yaml.load(f.read())
    d['advanced_inputs'] = uiconf
    print(yaml.dump(d))

def generate_ui_data(c, cell):
    return { "id": c,
        "description": prevcolumn(c),
        "name": c,
        "type": "float",
        "ui": "float",
        "value": c,
        "default": nextcolumn(c),
        "unit": nextcolumn(nextcolumn(c)),
        "format":cell.number_format,
        }

if __name__ == "__main__":
    advanced_input_cells(sys.argv[1], sys.argv[2])
