from setuptools import setup, find_packages
from pathlib import Path
import os
import setuptools

this_directory = os.path.abspath(os.path.dirname(__file__))

with open(os.path.join(this_directory, "README.md"), "r") as fh:
    long_description = fh.read()

setup(
    # options={'bdist_wheel':{'universal':True}},
    name='ClointFusion',
    author='Mayur Patil',
    author_email = 'mayur@cloint.com',
    packages=find_packages(), 
    include_package_data=True,
    zip_safe=False,
    version='0.1.34',
    description="Python based Automation (RPA) Platform",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url='https://github.com/ClointFusion/ClointFusion',
    setup_requires=["wheel",'numpy',"setuptools"],
    keywords='ClointFusion,RPA,Python,Automation,BOT,Software BOT,ROBOT',
    license="BSD",
    install_requires=open('requirements.txt').read().split('\n'),
    # py_modules=['ClointFusion'],
  classifiers=[
    'Development Status :: 4 - Beta',
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
            'cf = ClointFusion.ClointFusion:cli_cf',
            'cf_vlookup = ClointFusion.ClointFusion:cli_vlookup',
            'cf_st = ClointFusion.ClointFusion:cli_speed_test',
            'whm = ClointFusion.ClointFusion:cli_bre_whm',
        ],
    },
  python_requires='>=3.7, <4',

  project_urls={  # Optional
      'Date ❤️ with ClointFusion': 'https://lnkd.in/gh_r9YB',
      'WhatsApp Community': 'https://chat.whatsapp.com/DkY9QKmQkTZIv1CsOVrgWW',
      'Hackathon Website': 'https://tinyurl.com/ClointFusion',
      'Discord': 'https://discord.com/invite/tsMBN4PXKH',
      'Bug Reports': 'https://github.com/ClointFusion/ClointFusion/issues',
      'Discussion Forum': 'https://github.com/ClointFusion/ClointFusion/discussions',
      'Source Code': 'https://github.com/ClointFusion/ClointFusion/'
  },
    # package_data={"ClointFusion": ["*.pyd"]},
    has_ext_modules=lambda: True
)

# python setup.py build bdist_wheel rotate --match=*.exe*,*.egg*,*.tar.gz*,*.whl* --keep=1

# twine upload dist/* --verbose
# import time; start = time.process_time() ; import ClointFusion  ; print(time.process_time() - start)

# setup(
#     ...,
#     install_requires=[
#         "enum34;python_version<'3.4'",
#         "pywin32 >= 1.0;platform_system=='Windows'",
#     ],
# )