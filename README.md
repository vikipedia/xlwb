# xlwb
This project aims at creating automation tool for updating predefined website
from a excel sheet

How to use it?
1. If you have excelsheet with some formulas and some data in it, parse it using
 excelparse.py utility
```
python excelparse.py -h
usage: Excel data generator [-h] [-o OUTPUT] filename

positional arguments:
  filename              input excel filename

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        output filename, if not given input file's name will
                        be used with bin extension.
```
2. Once data for excelsheet is generated using excelparse utility. follow
wbtool/guild_inputs.txt to write a yaml file for web application. This yaml
file will be used to created WEB based UI for your application.

3. create a folder to store your configurations and files. lets call it
<YOURTOOLDATA>. Copy exceldata generated using excelrun utility and yaml file
in the folder <YOURTOOLDATA>

4. At any of your preferred location create a configuration file xlwb.cfg. this
should be configured with following variables.
```
EXCELTOOLSDIR="<YOURTOOLDATA>""
EXCELTOOLS="<Comma seperated list of yaml files>"
```

5. create virtualenv with
```
python -m venv <VIRTUALENVPATH>
```
here <VIRTUALENVPATH> is some location where virtual environment will store
python packages. for more details see documentation of `virtualenv`

6. activate virtualenv by excexuting following command on command prompt.
```
source <VIRTUALENVPATH>/bin/activate
```
on windows use cmd
```
<VIRTUALENVPATH>\Lib\activate.bat
```

7. run following commands in activated environment. <XLWB_CLONED_REPO> is path
where you have cloned git repository on your computer.
```
pip install --editable <XLWB_CLONED_REPO>
```

8. set environment variables and run flask
```
XLWB_SETTINGS="<path of xlwb.cfg file>" FLASK_APP=xlwb.wbtool flask run
```

If you are on Windows, the environment variable syntax depends on command line
interpreter.

On Command Prompt:
```
C:\path\to\app>set XLWB_SETTINGS="<path of xlwb.cfg file>"
C:\path\to\app>set FLASK_APP=xlwb.wbtool
C:\path\to\app>flask run
```
And on PowerShell:
```
PS C:\path\to\app> $env:XLWB_SETTINGS = "<path of xlwb.cfg file>"
PS C:\path\to\app> $env:FLASK_APP = xlwb.wbtool
PS C:\path\to\app> flask run
```

Alternatively you can use python -m flask:
python -m flask run

this should start a server with following prompt messages
```
 * Serving Flask app "excelapp" (lazy loading)
 * Environment: production
   WARNING: Do not use the development server in a production environment.
   Use a production WSGI server instead.
 * Debug mode: off
 * Running on http://127.0.0.1:5000/ (Press CTRL+C to quit)
```
9. Open browser and open location  http://127.0.0.1:5000/
