from setuptools import setup, find_packages
from pathlib import Path
import os

this_directory = os.path.abspath(os.path.dirname(__file__))

with open(os.path.join(this_directory, "README.md"), "r") as fh:
    long_description = fh.read()

setup(
    options={'bdist_wheel':{'universal':True}},
    name='ClointFusion',
    author='ClointFusion',
    author_email = 'ClointFusion@cloint.com',
    packages=find_packages(), 
    include_package_data=True,
    zip_safe=False,
    version='1.1.4',
    description="Python based Automation (RPA) Platform",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url='https://github.com/ClointFusion/ClointFusion',
    setup_requires=["wheel",'numpy',"setuptools"],
    keywords='ClointFusion,RPA,Python,Automation,BOT,Software BOT,ROBOT,Dost',
    license="BSD",
    install_requires=open('requirements.txt').read().split('\n'),

    # py_modules=['ClointFusion'],
  classifiers=[
    'Development Status :: 5 - Production/Stable',
    'Environment :: Console',
    'Intended Audience :: Developers',      
    'Topic :: Software Development :: Build Tools',
    'License :: OSI Approved :: BSD License',
    'Natural Language :: English',
    'Operating System :: OS Independent',
    'Framework :: Robot Framework',
    'Programming Language :: Python',
  ],
    entry_points={
        'console_scripts': [
            'colab = ClointFusion.ClointFusion:cli_colab_launcher',
            'dost = ClointFusion.ClointFusion:cli_dost',
            'bol = ClointFusion.ClointFusion:cli_bol',
            'cf_tray = ClointFusion.ClointFusion:cli_whm',
            'cf = ClointFusion.ClointFusion:cli_cf',
            'cf_vlookup = ClointFusion.ClointFusion:cli_vlookup',
            'cf_st = ClointFusion.ClointFusion:cli_speed_test',
            'cf_work = ClointFusion.ClointFusion:cli_bre_whm',
            'cf_wm = ClointFusion.ClointFusion:cli_send_whatsapp_msg',
            'cf_sm = ClointFusion.ClointFusion:cli_call_sm',
            'cf_like = ClointFusion.ClointFusion:cli_auto_liker',
            'cf_py = ClointFusion.ClointFusion:cli_cf_py',
            'cf_tour = ClointFusion.ClointFusion:cli_cf_tour',
        ],
    },
  python_requires='>=3.8.5, <=3.9.8',

  project_urls={  # Optional
      'Date ❤️ with ClointFusion': 'https://sites.google.com/view/clointfusion-hackathon/date-with-clointfusion',
      'WhatsApp Community': 'https://chat.whatsapp.com/JKr7m0avmkIFwYgMarShZG',
      'Hackathon Website': 'https://tinyurl.com/ClointFusion',
      'Documentation': 'https://clointfusion.readthedocs.io',
      'Bug Reports': 'https://github.com/ClointFusion/ClointFusion/issues',
      'Windows EXE': 'https://github.com/ClointFusion/ClointFusion/releases/download/v1.0.0/ClointFusion_Community_Edition.exe',
      'Medium' : 'https://medium.com/@clointfusion/bd152f4a1e0d?source=friends_link&sk=25b6051d75a8a4bb3e9a3e1a46516766'
  },
    # package_data={"ClointFusion": ["*.pyd"]},
)

# python -m pip install --upgrade pip setuptools wheel build
# python setup.py build bdist_wheel --universal rotate --match=*.exe*,*.egg*,*.tar.gz*,*.whl* --keep=1

# twine upload dist/* --verbose
# import time; start = time.process_time() ; import ClointFusion  ; print(time.process_time() - start)
