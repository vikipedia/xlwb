from flask import Flask, render_template, request, redirect
import yaml
from xlwb.xlspy.excelfunctions import excelrange
from xlwb.xlspy import excelexec
import forms
import pickle
import json

app = Flask(__name__)
app.secret_key = b'\x08\x19\xf3\x0e\xfb\x80\x11\x13\x13\xb8\x82c\x99}\x9e{'

def get_range(data, r):
    return [[data[c] for c in row] for row in r]

def prepare_data():
    confdata = yaml.load(open("inputs_conf.yaml"))
    with open(confdata['exceldata'], "rb") as f:
        exceldata = pickle.load(f)
    input_cells = confdata['input_cells']
    
    for item, value in input_cells.items():
        d = value['description']
        value['description'] = exceldata.get(d, d)
        if value['ui'] == "menu":
            m = value['menudata']
            if isinstance(m, str) and "!" in m:
                r = excelrange(m)
                value['menudata'] = get_range(exceldata, r)
            else:
                value['menudata'] = m

    return  confdata

app.conf  = prepare_data()

def _exceldata():
    with open(app.conf['exceldata'], "rb") as f:
        return pickle.load(f)
    

@app.route("/", methods = ["GET", "POST"])
def index():
    Form = forms.get_form(app.conf['input_cells'], _exceldata())
    form = Form()
    return render_template("inputform.html", form=form)

@app.route("/compute", methods = ["GET","POST"])
def compute():
    exceldata = _exceldata()
    Form = forms.get_form(app.conf['input_cells'], _exceldata())
    form = Form(request.form)
    
    if request.method == 'POST':
        inputs = {k:app.conf[k] for k in ["output", "macro"]}
        inputs['input_cells'] = {item:getattr(form,item).data for item in app.conf['input_cells']}
    
        excelexec.compute(exceldata,inputs)
        o = get_range(exceldata,  excelrange(app.conf['output']))
        columns = [('string', 'Year'),
                   ('number', 'Fuel Cost'),
                   ('number', 'Expenses')]
        chartdata = [('2004',  1000,      400),
                     ('2005',  1170,      460),
                     ('2006',  660,       1120),
                     ('2007',  1030,      540)]

        return render_template("table.html", output=o, data=chartdata, columns=columns)

