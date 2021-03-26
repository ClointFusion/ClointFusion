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
    version='0.0.99',
    description="Pythonic RPA (Automation) Platform",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author='Cloint India Pvt. Ltd',
    author_email='automation@cloint.com',
    url='https://github.com/ClointFusion/ClointFusion',
    setup_requires=["wheel",'numpy',"setuptools"],
        
    keywords='ClointFusion,RPA,Python,Automation,BOT,Software BOT,ROBOT',

    install_requires=[            
          "setuptools >= 51.1.2",
          "wheel >= 0.34.2",
          "watchdog >= 1.0.2",
          "Pillow >= 7.2.0",
          "pynput >= 1.7.1",
          "pif >= 0.8.2",
          "PyAutoGUI >= 0.9.52",
          "PySimpleGUI >= 4.29.0",
          "bs4 >= 0.0.1",
          "clipboard >= 0.0.4",
          "emoji >= 0.6.0",
          "folium >= 0.11.0",
          "helium >= 3.0.5",
          "imutils >= 0.5.3",
          "kaleido >= 0.0.3.post1",
          "keyboard >= 0.13.5",
          "matplotlib >= 3.3.2",
          "numpy >= 1.19.2",
          "opencv-python >= 4.4.0.44",
          "openpyxl >= 3.0.5",
          "pandas >= 1.1.3",
          "plotly >= 4.11.0",
          "requests >= 2.24.0",
          "selenium >= 3.141.0",
          "texthero >= 1.0.9",
          "wordcloud >= 1.8.0",
          "zipcodes >= 1.1.2",
          "pathlib3x >= 1.3.9",
          "pathlib >= 1.0.1",
          "PyQt5 >= 5.15.2",
          "email-validator >= 1.1.1",
          "testresources >= 2.0.1",
          "scikit-image >= 0.17.2",
          "pivottablejs >= 0.9.0",
          "ipython >= 7.19.0",
          "comtypes >= 1.1.7",
          "cryptocode >= 0.1",
          "ImageHash >= 4.2.0",
          "get-mac >= 0.8.2",
          "xlsx2html >= 0.2.2 ",
          "simplegmail >= 3.1.5",
          "xlwings >= 0.22.3",
          "jupyterlab >= 3.0.0",    
          "notebook",
          "pygments >= 2.7.4",
          ],
  classifiers=[
    'Development Status :: 4 - Beta',
    'Intended Audience :: Developers',      
    'Topic :: Software Development :: Build Tools',
    'License :: OSI Approved :: BSD License',
    'Natural Language :: English',
    'Operating System :: OS Independent',
    'Framework :: Robot Framework',
    'Programming Language :: Python',
  ],
  python_requires='>=3.7, <4',

  project_urls={  # Optional
      'Bug Reports': 'https://github.com/ClointFusion/ClointFusion/issues',
      'Discussion Forum': 'https://github.com/ClointFusion/ClointFusion/discussions',
      'Source': 'https://github.com/ClointFusion/ClointFusion/',
  },
)

# python setup.py build sdist bdist_wheel

# twine upload dist/*