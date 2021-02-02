from setuptools import setup, find_packages
from pathlib import Path
import os
import setuptools

this_directory = os.path.abspath(os.path.dirname(__file__))

with open(os.path.join(this_directory, "README.md"), "r") as fh:
    long_description = fh.read()

setup(
    name='ClointFusion',
    packages=find_packages(), 
    include_package_data=True,
    zip_safe=False,
    version='0.0.91',
    description="Pythonic RPA (Automation) Platform",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author='Cloint India Pvt. Ltd',
    author_email='automation@cloint.com',
    url='https://github.com/ClointFusion/ClointFusion',
    setup_requires=["wheel",'numpy',"setuptools"],
        
    keywords=['ClointFusion','RPA','Python','Automation','BOT','Software BOT','ROBOT'],

    install_requires=[            
          "watchdog == 1.0.2",
          "Pillow == 8.1.0",
          "pynput == 1.7.2",
          "pif == 0.8.2",
          "PyAutoGUI == 0.9.52",
          "PySimpleGUI == 4.33.0",
          "bs4 == 0.0.1",
          "clipboard == 0.0.4",
          "emoji == 0.6.0",
          "folium == 0.12.0",
          "helium == 3.0.5",
          "imutils == 0.5.3",
          "kaleido == 0.0.3.post1",
          "keyboard == 0.13.5",
          "matplotlib == 3.3.3",
          "numpy == 1.19.5",
          "opencv-python == 4.5.1.48",
          "openpyxl == 3.0.5",
          "pandas == 1.1.3",
          "plotly == 4.14.3",
          "requests == 2.25.1",
          "selenium == 3.141.0",
          "texthero == 1.0.9",
          "wordcloud == 1.8.1",
          "xlrd == 1.2.0",
          "zipcodes == 1.1.2",
          "pathlib3x == 1.3.9",
          "pathlib == 1.0.1",
          "PyQt5 == 5.15.2",
          "email-validator == 1.1.2",
          "testresources == 2.0.1",
          "scikit-image == 0.18.1",
          "pivottablejs == 0.9.0",
          "ipython == 7.19.0",
          "cryptocode == 0.1",
          "ImageHash == 4.2.0",
          "jupyterlab",
          "notebook"
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
  python_requires='>=3.9',
)

# python setup.py build sdist bdist_wheel

# twine upload dist/*