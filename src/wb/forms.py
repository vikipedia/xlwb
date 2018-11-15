from wtforms import StringField, IntegerField, SelectField, validators
from flask_wtf import FlaskForm

import collections

def validators_(fielddata):
    fd= fielddata
    v = []
    if "max" in fd and "min" in fd:
        v.append(validators.NumberRange(min=fd['min'], max=fd['max']))
    elif fielddata['ui'] == "text":
        v.append(validators.Length(min=fd.get('minlen',0),max=fd.get('maxlen',20)))
    return v


def get_value(v, data):
    if "!" in v:
        return data.get(v, None)
    return v

def create_field(name, fielddata, data):
    typemap = {
        "int":IntegerField,
        "text":StringField,
        "menu":SelectField
        }
    
    fieldclass = typemap[fielddata['ui']]
    kwargs= {'default':get_value(fielddata.get('value', None), data),
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
            

    Field = collections.namedtuple("Field", ["name","field"])

    fields = []
    for item, value in data.items():
        f = create_field(item, value, cellmap)
        setattr(InputsForm, item, f)
        fields.append(item)

    return InputsForm
