from wtforms import StringField, IntegerField, FloatField
from wtforms import SelectField, BooleanField, validators
from flask_wtf import FlaskForm
import re

import collections



def validators_(fielddata):
    fd= fielddata
    v = []
    if "max" in fd and "min" in fd:
        v.append(validators.NumberRange(min=fd['min'], max=fd['max']))
    elif fielddata['ui'] == "text":
        v.append(validators.Length(min=fd.get('minlen',0),max=fd.get('maxlen',30)))
    return v

def get_format(fielddata):

    if "format" in fielddata:
        nums = re.compile(r'(0*)\.?(0*)(%)?')
        f = fielddata['format']
        g = nums.match(f)
    else:
        return "{}", None

    if f.strip().lower()=="general" or not g:
        return "{}", None
    else:
        x,y,z = g.groups()
        return "{"+ ":{}.{}f".format(len(x), len(y)) + "}", z


def get_value(v, data):
    if "!" in v:
        return data.get(v, None)
    return v

def create_field(fielddata, data):
    typemap = {
        "int":IntegerField,
        "text":StringField,
        "menu":SelectField,
        "float":FloatField,
        "bool":BooleanField,
        }
    name = fielddata['id']
    fieldclass = typemap[fielddata['ui']]

    def value(v):
        if fielddata['ui'] in ['int', 'float']:
            return v if v else 0
        return v

    def format_(v):
        f,percent = get_format(fielddata)
        if percent :
            return f.format(v*100 if v else 0)
        else:
            return f.format(value(v))

    kwargs= {'default':format_(get_value(fielddata.get('value', None), data)),
             'label':get_value(fielddata.get('description', None), data),
             'validators': validators_(fielddata),
             'id':name,
    }
    if fielddata['ui']== "menu":
        kwargs['choices'] = [(k,k) for k in fielddata['menudata']]
    return fieldclass(**kwargs)

def get_form(data, cellmap):
    class InputsForm(FlaskForm):
        def get_fields_(self):
            return [getattr(self, name) for name in self._fields]

        def get_fields__(self, section):
            return (getattr(self, name) for name in sections[section])

        def get_sections(self):
            return sections.keys()

    sections = {section:[item['id'] for item in data[section]] for section in data}

    fields = []
    for section in data:
        for item in data[section]:
            f = create_field(item, cellmap)
            setattr(InputsForm, item['id'], f)
            fields.append(item)

    return InputsForm
