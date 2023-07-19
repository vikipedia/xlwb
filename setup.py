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
    "Flask==2.2.5",
    "Flask-WTF==0.14.2",
    "networkx==2.2",
    "PyYAML>=4.2b1",
    "numpy",
    "openpyxl==2.4.8",
    "Click",
    "pytest",
    ]
)
