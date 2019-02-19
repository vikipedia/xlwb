from distutils.core import setup, Extension


setup(
    name="xlwb",
    version='1.0',
    description='xlwb, excel to web Library to run excel functionality from python as webservice!',
    author='Vikrant Patil',
    author_email='vikrant.patil@gmial.com',
    url='https://github.com/vikipedia/xlwb.git',
    packages=['xlwb','xlwb.xlspy', 'xlwb.wbtool'],
    include_package_data=True,
    install_requires=[
    "atomicwrites==1.2.1",
    "attrs==18.2.0",
    "Click==7.0",
    "decorator==4.3.0",
    "et-xmlfile==1.0.1",
    "Flask==1.0",
    "Flask-WTF==0.14.2",
    "itsdangerous==1.1.0",
    "jdcal==1.4",
    "Jinja2==2.10",
    "MarkupSafe==1.1.0",
    "more-itertools==4.3.0",
    "networkx==2.2",
    "numpy==1.15.4",
    "openpyxl==2.4.8",
    "pathlib2==2.3.2",
    "pluggy==0.8.0",
    "py==1.7.0",
    "pytest==3.10.1",
    "PyYAML>=4.2b1",
    "six==1.11.0",
    "Werkzeug==0.14.1",
    "WTForms==2.2.1",
    ]
)
