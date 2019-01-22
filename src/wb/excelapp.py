from flask import Flask, render_template, request, redirect
import yaml
from xlwb.xlspy.excelfunctions import excelrange, flatten
from xlwb.xlspy import excelexec
import forms, charts
import pickle
import json

app = Flask(__name__)
app.secret_key = b'\x08\x19\xf3\x0e\xfb\x80\x11\x13\x13\xb8\x82c\x99}\x9e{'

def get_range(data, r):
    return [[data[c] for c in row] for row in r]


def prepare_data(conffile):
    print("Reading configuration files....")
    confdata = yaml.load(open(conffile))
    exceldata = _exceldata(confdata)

    print("Reading Data Done")

    input_cells = confdata['input_cells']

    for value in input_cells:
        d = value['description']
        value['description'] = exceldata.get(d, d)
        if value['ui'] == "menu":
            m = value['menudata']
            if isinstance(m, str) and "!" in m:
                r = excelrange(m)
                value['menudata'] = flatten(get_range(exceldata, r))
            else:
                value['menudata'] = m

    return  confdata

def _exceldata(conf):
    with open(conf['exceldata'], "rb") as f:
        return pickle.load(f)


@app.route("/<toolname>", methods = ["GET","POST"])
def compute(toolname):
    conf  = prepare_data(toolname + ".yaml" )
    exceldata = _exceldata(conf)
    Form = forms.get_form(conf['input_cells'], exceldata)

    if request.method == 'POST':
        form = Form(request.form)
        inputs = {k:conf[k] for k in ["output", "macro"]}
        inputs['input_cells'] = {item['id']:getattr(form,item['id']).data for item in conf['input_cells']}

        excelexec.compute(exceldata,inputs)
        o = get_range(exceldata,  excelrange(conf['output']))
        chartdata = charts.process_chartdata(exceldata, conf)

        return render_template("table.html", output=o, chartdata=chartdata)

    form = Form()
    return render_template("inputform.html", toolname=toolname, form=form)
