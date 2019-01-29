# xlwb
This project aims at creating automation tool for updating predefined website from a excel sheet

How to use it?
1. If you have excelsheet with some formulas and some data in it, parse it using excelparse.py utility
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
2. Once data for excelsheet is generated using excelparse utility. follow wbtool/guild_inputs.txt to write a yaml file for web application.

3. copy exceldata generated using excelrun utility and yaml file in folder wbtool.

4. create virtualenv with
```
python -m venv FOLDERPATH
```
5. activate virtualenv using
```
source FODERPATH/bin/activate
```
on windows use
```
FOLDERPATH\bin\activate
```
6. run following command in activated environment
```
python xlwb/src/setup.py install
cd xlwb/src/wbtool
pip install -r requirements.txt
python excelapp.py
```
this should start a server with following prompt messages
```
 * Serving Flask app "excelapp" (lazy loading)
 * Environment: production
   WARNING: Do not use the development server in a production environment.
   Use a production WSGI server instead.
 * Debug mode: off
 * Running on http://127.0.0.1:5000/ (Press CTRL+C to quit)
```
7. Open browser and open location  http://127.0.0.1:5000/
