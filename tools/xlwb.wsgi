activate_this = '/var/www/xlwb/xlwb3/bin/activate_this.py'
with open(activate_this) as file_:
    exec(file_.read(), dict(__file__=activate_this))

import os
os.environ['XLWB_SETTINGS'] = "/var/www/xlwb/xlwb.cfg"

from xlwb.wbtool import app as application
