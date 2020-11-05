from setuptools import setup, find_packages, Extension
from setuptools.command.build_ext import build_ext
from setuptools.command.build_py import build_py as _build_py
from pathlib import Path
import os
from setuptools.dist import Distribution
import fnmatch
import sysconfig
import numpy
from setuptools_cythonize import get_cmdclass

from Cython.Build import cythonize

this_directory = os.path.abspath(os.path.dirname(__file__))

with open(os.path.join(this_directory, "README.md"), "r") as fh:
    long_description = fh.read()

setup(
    name='ClointFusion',
    packages=find_packages(), 
    include_package_data=True,
    zip_safe=False,
    version='0.0.82',
    description="Python based functions for RPA (Automation)",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author='Cloint India Pvt. Ltd',
    author_email='automation@cloint.com',
    url='https://github.com/ClointFusion/ClointFusion',
    setup_requires=['numpy'],
    include_dirs=numpy.get_include(),
    
    keywords=['ClointFusion','RPA','Python','Automation','BOT','Software BOT','ROBOT'],

    install_requires=[            
          "wheel == 0.34.2",
          "cmake",
          "Pillow == 7.2.0",
          "PyAutoGUI == 0.9.52",
          "PyQt5_sip == 12.8.1",
          "PySimpleGUI == 4.29.0",
          "bs4 == 0.0.1",
          "clipboard == 0.0.4",
          "emoji == 0.6.0",
          "folium == 0.11.0",
          "helium == 3.0.5",
          "imutils == 0.5.3",
          "kaleido == 0.0.3.post1",
          "keyboard == 0.13.5",
          "matplotlib == 3.3.2",
          "numpy == 1.19.2",
          'PyObjC;platform_system=="Darwin"',
          'PyGObject;platform_system=="Linux"',
          'pyreadline; platform_system == "Windows"',
          "opencv_python == 4.4.0.44",
          "openpyxl == 3.0.5",
          "pandas == 1.1.3",
          "pdfplumber == 0.5.23",
          "plotly == 4.11.0",
          "requests == 2.24.0",
          "selenium == 3.141.0",
          "setuptools == 50.3.2",
          "texthero == 1.0.9",
          "watchdog == 0.10.3",
          "wordcloud == 1.8.0",
          "xlrd == 1.2.0",
          "zipcodes == 1.1.2",
          "pathlib3x == 1.3.9",
          "pathlib == 1.0.1",
          "PyQt5 == 5.15.1",
          "pynput == 1.7.1",
          "pif == 0.8.2",
          "email-validator == 1.1.1",
          "slack-webhook == 1.0.3",
          "scikit-image == 0.17.2",
          "jupyterlab",
          "notebook",
      ],
  classifiers=[
    'Development Status :: 3 - Alpha',
    'Intended Audience :: Developers',      
    'Topic :: Software Development :: Build Tools',
    'License :: OSI Approved :: BSD License',
    'Natural Language :: English',
    'Operating System :: OS Independent',
    'Framework :: Robot Framework',
    'Programming Language :: Python',
  ],
  python_requires='>=3.8',
)

# python setup.py build sdist bdist_wheel

# twine upload dist/*