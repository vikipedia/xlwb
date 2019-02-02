from flask import Flask, render_template, request, redirect, url_for
import yaml
from xlwb.xlspy.excelfunctions import excelrange, flatten
from xlwb.xlspy import excelexec
import forms, charts
import pickle
import json, os

app = Flask(__name__)
app.secret_key = b'\x08\x19\xf3\x0e\xfb\x80\x11\x13\x13\xb8\x82c\x99}\x9e{'

def get_range(data, r):
    return [[data[c] for c in row] for row in r]


def evaluate_conf(input_cells, exceldata):
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

def prepare_data(conffile):
    print("Reading configuration files....")
    confdata = yaml.load(open(conffile))
    exceldata = _exceldata(confdata)
    input_cells = confdata['input_cells']
    evaluate_conf(input_cells, exceldata)
    return  confdata

def _exceldata(conf):
    with open(conf['exceldata'], "rb") as f:
        return pickle.load(f)

def from_form(input_cells, form):
    def value(item):
        v = getattr(form,item['id']).data
        if "percent" in item:
            return v/100.0 if v else 0
        else:
            return v
    return {item['id']:value(item) for item in input_cells}

def prepare_inputs(conf, input_cells, form):
    """
    prepare inputs map for exelexec
    """
    items = ["output"]
    if "macro" in conf:
        items.append("macro")
    inputs = {k:conf[k] for k in items}
    inputs['input_cells'] = from_form(input_cells, form)
    return inputs

def get_tool_info():
    conffiles = [f for f in os.listdir() if f.endswith("yaml") and os.path.isfile(f)]
    d = {}
    for file in conffiles:
        with open(file) as f:
            conf = yaml.load(f)
            d[file.split(".")[0]] = conf['title']
    return d

@app.route("/", methods = ["GET"])
def index():
    d = get_tool_info()
    return render_template("index.html", title="Excel Web Tool", toolinfo=d)

@app.route("/<toolname>", methods = ["GET","POST"])
def compute(toolname):
    d = get_tool_info()
    conf  = prepare_data(toolname + ".yaml" )
    exceldata = _exceldata(conf)
    Form = forms.get_form(conf['input_cells'], exceldata)
    form = Form()
    advanced =   {"advanced":True} if "advanced_inputs" in conf else {}

    if request.method == 'POST' and form.validate_on_submit():
        if "advanced" in request.form:
            return redirect(url_for('advanced_compute', toolname=toolname), code=307)#post
        else:
            inputs = prepare_inputs(conf, conf['input_cells'], form)
            params = [v for v in inputs['input_cells'].values()]
            excelexec.compute(exceldata,inputs)
            o = get_range(exceldata,  excelrange(conf['output']))
            chartdata = charts.process_chartdata(exceldata, conf)
            return render_template("table.html", toolname=toolname, output=o, title=conf['title'],
                                    toolinfo=d, chartdata=chartdata, params=params)

    return render_template("inputform.html", toolname=toolname,
                            form=form, title=conf['title'],
                            toolinfo = d,
                            **advanced)

def get_other_data(exceldata, advanced_inputs, itemname="default"):
    def value(item):
        v = exceldata.get(item[itemname])
        if "percent" in item and itemname=="default":
            return v*100.0 if v else 0
        else:
            return v
    return {c['id']:value(c) for c in advanced_inputs}

def pre_execute_cells(exceldata, advanced_inputs):
    ids = ",".join([item['id'] for item in advanced_inputs])
    excelexec.compute_range(exceldata, ids)
    defaults =  ",".join([item['default'] for item in advanced_inputs if "default" in item])
    excelexec.compute_range(exceldata, defaults)


@app.route("/advanced/<toolname>", methods = ["POST"])
def advanced_compute(toolname):
    d = get_tool_info()
    conf  = prepare_data(toolname + ".yaml" )
    exceldata = _exceldata(conf)
    Form1 = forms.get_form(conf['input_cells'] , exceldata)
    form1 = Form1()

    inputs = prepare_inputs(conf, conf['input_cells'], form1)
    params = [v for v in inputs['input_cells'].values()]
    excelexec.handle_macro(exceldata, inputs)
    pre_execute_cells(exceldata, conf['advanced_inputs'])

    Form2 = forms.get_form(conf['advanced_inputs'], exceldata)
    form2 = Form2()
    advanced_inputs = from_form(conf['advanced_inputs'], form2)
    inputs['input_cells'].update(advanced_inputs)

    if "finish" in request.form and form1.validate_on_submit() and form2.validate_on_submit():
        exceldata = _exceldata(conf)
        excelexec.compute(exceldata,inputs)
        o = get_range(exceldata,  excelrange(conf['output']))
        chartdata = charts.process_chartdata(exceldata, conf)
        return render_template("table.html", toolname=toolname, output=o,
                                title=conf['title'], toolinfo=d,chartdata=chartdata)
    else:
        defaults = get_other_data(exceldata, conf['advanced_inputs'], "default")
        units = get_other_data(exceldata, conf['advanced_inputs'], "unit")
        return render_template("advancedform.html", toolname=toolname, title=conf['title'],
                                toolinfo=d,params=params,
                                form1=form1, form2= form2, defaults=defaults, units=units)

if __name__=="__main__":
    app.run()
