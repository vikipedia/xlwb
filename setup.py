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
)
