from flask import Flask, render_template, request, redirect, url_for
import yaml
from xlwb.xlspy.excelfunctions import excelrange, flatten
from xlwb.xlspy import excelexec
from . import forms, charts
import pickle
import json, os, sys

app = Flask(__name__)
app.secret_key = b'\x08\x19\xf3\x0e\xfb\x80\x11\x13\x13\xb8\x82c\x99}\x9e{'
app.config.from_envvar('XLWB_SETTINGS')
sys.path.insert(0, app.config['EXCELTOOLSDIR'])

def get_range(data, r):
    return [[data[c] for c in row] for row in r]


def evaluate_conf(inputs, exceldata):
    def evaluate_conf_(input_cells):
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
    for section in inputs:
        evaluate_conf_(inputs[section])

def prepare_data(toolname):
    print("Reading configuration files....")
    path = filepath_( toolname + ".yaml")
    confdata = yaml.load(open(path))
    exceldata = _exceldata(confdata)
    input_cells = confdata['input_cells']
    evaluate_conf(input_cells, exceldata)
    return  confdata

def filepath_(f):
    return os.path.join(app.config['EXCELTOOLSDIR'], f)

def _exceldata(conf):
    with open(filepath_(conf['exceldata']), "rb") as f:
        return pickle.load(f)


def from_form(inputs,form):
    def from_form_(input_cells):
        def value(item):
            v = getattr(form,item['id']).data
            if "format" in item and "%" in item['format']:
                return float(v)/100.0 if v else 0
            else:
                return v
        return {item['id']:value(item) for item in input_cells}
    d = {}
    for section in inputs:
        d.update(from_form_(inputs[section]))
    return d

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
    conffiles = [t.strip() for t in app.config['EXCELTOOLS'].strip().split(",")]
    d = {}
    for file in conffiles:
        with open(filepath_( file)) as f:
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
    conf  = prepare_data(toolname)
    exceldata = _exceldata(conf)
    Form = forms.get_form(conf['input_cells'], exceldata)
    form = Form()
    advanced =   {"advanced":True} if "advanced_inputs" in conf else {}

    if request.method == 'POST' and form.validate_on_submit():
        if "advanced" in request.form:
            return redirect(url_for('advanced_compute', toolname=toolname), code=307)#post
        else:
            inputs = prepare_inputs(conf, conf['input_cells'], form)
            excelexec.compute(exceldata,inputs)
            o = get_range(exceldata,  excelrange(conf['output']))
            chartdata = charts.process_chartdata(exceldata, conf)
            return render_template("table.html", toolname=toolname, output=o, title=conf['title'],
                                    toolinfo=d, chartdata=chartdata, params=get_params(form))

    return render_template("inputform.html", toolname=toolname,
                            form=form, title=conf['title'],
                            toolinfo = d,
                            **advanced)

def get_other_data(exceldata, advanced_inputs, itemname="default"):
    def value_(v, item):
        if item['ui'] in ['int', 'float']:
            return v if v else 0
        return v if v else ""

    def value(item):
        f, z = forms.get_format(item)
        v = exceldata.get(item[itemname])
        if z and itemname=="default":
            return f.format(v*100.0 if v else 0)
        elif itemname=="default":
            return f.format(value_(v, item))
        else:
            return v
    d = {}
    for s in advanced_inputs:
        d.update({c['id']:value(c) for c in advanced_inputs[s]})
    return d

def pre_execute_cells(exceldata, advanced_inputs):
    for s in advanced_inputs:
        ids = ",".join([item['id'] for item in advanced_inputs[s]])
        excelexec.compute_range(exceldata, ids)
        defaults =  ",".join([item['default'] for item in advanced_inputs[s] if "default" in item])
        excelexec.compute_range(exceldata, defaults)


def get_params(*forms):
    fields = []
    for form in forms:
        fl  = [(f.label, f.data) for f in form.get_fields_() if f.id!="csrf_token"]
        fields.extend(fl)
    return fields


@app.route("/advanced/<toolname>", methods = ["POST"])
def advanced_compute(toolname):
    d = get_tool_info()
    conf  = prepare_data(toolname)
    exceldata = _exceldata(conf)
    Form1 = forms.get_form(conf['input_cells'] , exceldata)
    form1 = Form1()

    basicinputs = prepare_inputs(conf, conf['input_cells'], form1)
    params = [v for v in basicinputs['input_cells'].values()]
    excelexec.handle_macro(exceldata, basicinputs)
    pre_execute_cells(exceldata, conf['advanced_inputs'])

    Form2 = forms.get_form(conf['advanced_inputs'], exceldata)
    form2 = Form2()
    advanced_inputs = from_form(conf['advanced_inputs'], form2)
    inputs = dict(basicinputs)
    inputs['input_cells'].update(advanced_inputs)

    if "finish" in request.form and form1.validate_on_submit() and form2.validate_on_submit():
        exceldata = _exceldata(conf)
        excelexec.compute(exceldata,inputs)
        o = get_range(exceldata,  excelrange(conf['output']))
        chartdata = charts.process_chartdata(exceldata, conf)
        return render_template("table.html", toolname=toolname, output=o,params=get_params(form1, form2),
                                title=conf['title'], toolinfo=d,chartdata=chartdata)
    elif "basic" in request.form:
        return redirect(url_for("compute", toolname=toolname))
    else:
        defaults = get_other_data(exceldata, conf['advanced_inputs'], "default")
        units = get_other_data(exceldata, conf['advanced_inputs'], "unit")
        return render_template("advancedform.html", toolname=toolname, title=conf['title'],
                                toolinfo=d,
                                form1=form1, form2= form2, defaults=defaults, units=units)

if __name__=="__main__":
    app.run()
