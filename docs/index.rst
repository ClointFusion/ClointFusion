.. ClointFusion documentation master file, created by
   sphinx-quickstart on Mon Nov 22 13:09:24 2021.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

==============================
ClointFusion Community Edition
==============================

.. image:: https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Cloint-LOGO-New.png


*Welcome to ClointFusion , Made in India with* ❤️

**About ClointFusion**

ClointFusion is an Indian firm based in Vadodara, Gujarat. ClointFusion is a Python-based RPA platform for developing Software BOTs. Using AI,
we're working on Common Man's RPA.

**Project Status**

|PyPI| |PyPI - License| |PyPI - Status| |ClointFusion| |PyPI -
Downloads| |Libraries.io SourceRank| |PyPI - Format| |GitHub
contributors| |GitHub last commit|

|GitHub Repo stars| |Twitter URL| |YouTube Channel Subscribers| |Twitter
Follow|

=======================
Installing ClointFusion
=======================


**Windows OS:**


*Using EXE:*

Windows users can download EXE pre-loaded with Python 3.9 and ClointFusion package: <a href='https://github.com/ClointFusion/ClointFusion/releases/download/v1.0.0/ClointFusion.Community.Edition.exe' target="_blank">Windows EXE</a>

*Manually (without EXE):*

-  ClointFusion is compatible with both Windows 10 and Windows 11.
-  Installing on a Windows PC is a breeze.
-  Make certain that Python 3.8 or Python 3.9 is installed.
-  Then from the command prompt, execute the following command.

   ::

       pip install -U ClointFusion

**Ubuntu OS :**


-  Clointfusion requires sudo rights to install on Ubuntu.
-  Additional Linux packages must be installed before Clointfusion can
   be installed.
-  Make certain that Python 3.8 or Python 3.9 is installed.
-  Then, from the command prompt, execute the following command.

   ::

       sudo apt-get install python3-tk python3-dev
       sudo pip3 install ClointFusion

==================
Using ClointFusion
==================


**Windows** 

*ClointFusion can be accessed using one of two methods:*

-  **Terminal : Opens a Python interpreter with "import ClointFusion as cf " pre-loaded**


   ::

       cf_py

-  **Code Editor or IDE : Import ClointFusion first, and then run the file in Python.**


   ::

       # cf_bot.py

       import ClointFusion as cf

       cf.browser_activate()

   ::

       python cf_bot.py

**Ubuntu :**

*ClointFusion can be accessed using one of two methods:*

-  **Terminal : Opens a Python interpreter with the "import ClointFusion as cf" pre-loaded and the required sudo privileges.**
   

   ::

       sudo cf_py

-  **Code Editor or IDE : Run the file with sudo permissions.**


   ::

       # cf_bot.py

       import ClointFusion as cf

       cf.browser_activate()

   ::

       sudo python3 cf_bot.py

=======
Add-ons
=======


*1) ClointFusion's Voice-Guided, Fully Automated Self-Test:*


When you import ClointFusion for the first time, or upgrade to a new
version, you'll be prompted with the "ClointFusion's Automated Self-Test"
which highlights all of ClointFusion's 100+ features in action on your
computer while also confirming ClointFusion's compatibility with your
PC's settings and configurations. Once you have successfully completed
the self-test, you will receive an email with a self-test report.

Below is the speed up version of self-test.


`Click here to watch the Self-Test in
Action. <https://user-images.githubusercontent.com/67296473/139620682-d63f6ee6-a3f5-4ca9-9ea9-23216e571e3e.mp4>`__

*2) ClointFusion's function works in different modes:*


ClointFusion's Semi Automatic Mode


1. If you pass all the required parameters, function works silently. So,
   this is expert (Non-GUI) mode. This mode gives you more control over
   the function's parameters.
2. If you do not pass any parameter, GUI would pop-up asking you the
   required parameters. Next time, when you run the BOT, based upon your
   configuration, which you get to choose at the beginning of BOT run:

   -  If ``Semi-Automatic mode`` is OFF, GUI would pop-up again, showing
      you the previous entries, allowing you to modify the parameters.
   -  If ``Semi-Automatic mode`` in ON, BOT works silently taking your
      previous GUI entries.
   -  Toggle ``Semi-Automatic mode`` by using the following command

   ::

       cf.ON_semi_automatic_mode   # To turn ON semi automatic mode
       cf.OFF_semi_automatic_mode  # To turn OFF semi automatic mode

3. GUI Mode is for beginners. Anytime, if you are not getting how to use
   the function, just call an empty function (without parameters) and
   GUI would pop-up asking you for required parameters.



.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Auto_Semi_Auto.gif?raw=true


*3) BOL : Your automation voice based assistant*


``BOL`` is voice based automation assistant designed to execute BOTs
build out of ClointFusion without any human computer interaction.

Usage of BOL:

Open your favorite terminal and type ``bol`` or ``sudo bol`` for ubuntu
users. Within a moment, a personalized Virtual Assistant will be at your
service.

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/VERSION%201.1.0/CLI_BOL_GIF.gif?raw=true

*Note: bol is currently in development stage. More functionalities
are yet to be added.*

*4) WORK - The Work Hour Monitor*


``WORK`` is an intelligent application that detects each and every work
you do in your PC and displays a detailed work report.


Usage of WORK:


Open your favorite terminal and type ``cf_work``. A detailed work report
will be displayed.

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/VERSION%201.1.0/CLI_CF_WORK.gif?raw=true

***Note: All the information that is being collected by ``WORK`` is
stored in a securely maintained database in your system.***

*5) WhatsApp Bot - Send bulk WhatsApp messages*

ClointFusion's "WhatsApp Bot" is an automated utility tool that allows
you to send many customized messages to your contacts at once.

Usage of WhatsApp Bot:

Open your favorite terminal and type ``cf_wm``, and give path of the
excel, or ``cf_wm -e excel_path.xlsx``

`Click here to watch the WhatsApp Bot in
Action. <https://user-images.githubusercontent.com/67296473/139722199-37036526-2b1c-4120-a12d-bde3df2eb0d7.mp4>`__

=======================================
DOST: Build Build RPA Bots WITHOUT Code
=======================================

**DOST : Your friend in Automation**

``DOST`` is an interactive Blockly based ``no-code`` BOT Builder
platform built and optimized for ClointFusion-based BOT building. We
feel that automation is important for people other than programmers.
Using DOST, even a common man can create a BOT in minutes.

**Advantages of DOST**

-  Easy to Use.
-  Build BOT in minutes.
-  No prior Programming knowledge needed.

**Launch DOST client:**

**Windows**


Open your favorite browser and go to `https://dost.clointfusion.com` and start building bots.

Note : Make sure ClointFusion Tray is present or open terminal and type `cf_tray` to activate ClointFusion Tray menu.


**Ubuntu**


    Open your favorite terminal and type ``sudo dost`` and then type
    ``python3 dost.py``.

-  Want to change the chrome profile ?

   -  Use\ ``python3 dost.py "Profile 1"``

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/VERSION%201.1.0/CLI_DOST_GIF.gif?raw=true

**Build BOT with DOST :** `DOST
Website <https://dost.clointfusion.com/>`__


======================
ClointFusion in Action
======================

Now access more than 100 functions (hit ctrl+space in your IDE):

***TIP: You can find and inspect all of ClointFusion's functions using
only one function i.e., ``find()``. Just pass the partial name of the
function.***

::

    cf.find("sort")

    cf.find("gui")

********************************************
6 gui functions, to take any input from user
********************************************


+------------------------------------------------------+----------------------------------------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| Function                                             | Accepted Parameters                                                                    | Description                                                                                                                                                                  |
+======================================================+==========================================================+=============================+==============================================================================================================================================================================+
| cf.gui_get_any_input_from_user()                     | msgForUser="Please enter : ", password=False, multi_line=False, mandatory_field=True   |   Generic function to accept any input (text / numeric) from user using GUI. Returns the value in string format.                                                             |
+------------------------------------------------------+----------------------------------------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.gui_get_any_file_from_user()                      | msgForUser="the file : ", Extension_Without_Dot="*"                                    |   Generic function to accept file path from user using GUI. Returns the filepath value in string format.Default allows all files.                                            |
+------------------------------------------------------+----------------------------------------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.gui_get_consent_from_user()                       | msgForUser="Continue ?"                                                                | Generic function to get consent from user using GUI. Returns the string 'yes' or 'no'                                                                                        |
+------------------------------------------------------+----------------------------------------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.gui_get_dropdownlist_values_from_user()           | msgForUser=" ", dropdown_list=[], multi_select=True                                    | Generic function to accept one of the drop-down value from user using GUI. Returns all chosen values in list format.                                                         |
+------------------------------------------------------+----------------------------------------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.gui_get_excel_sheet_header_from_user()            | msgForUser=" "                                                                         | Generic function to accept excel path, sheet name and header from user using GUI. Returns all these values in dictionary format.                                             |
+------------------------------------------------------+----------------------------------------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.gui_get_folder_path_from_user()                   | msgForUser="the folder : "                                                             |  Generic function to accept folder path from user using GUI. Returns the folder path value in string format.                                                                 |
+------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Functions%20Light%20GIFs/gui_function.gif?raw=true


*******************************
4 functions on Mouse Operations
*******************************



+-------------------------------------------------------+----------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| Function                                              | Accepted Parameters                                      | Description                                                                                                                                                                  |
+=======================================================+==========================================================+==============================================================================================================================================================================+
| cf.mouse\_click()                                     | x=" ", y=" ", left\_or\_right="left", no\_of\_clicks=1   | Clicks at the given X Y Co-ordinates on the screen using ingle / double / triple click(s). Optionally copies selected data to clipboard (works for double / triple clicks)   |
+-------------------------------------------------------+----------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.mouse\_move()                                      | x=" ", y=" "                                             | Moves the cursor to the given X Y Co-ordinates                                                                                                                               |
+-------------------------------------------------------+----------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.mouse\_drag\_from\_to()                            | x1=" ", y1=" ", x2=" ",y2=" ", delay=0.5                 | Clicks and drags from X1 Y1 co-ordinates to X2 Y2 Co-ordinates on the screen                                                                                                 |
+-------------------------------------------------------+----------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.mouse\_search\_snip\_return\_coordinates\_x\_y()   | img=" ", wait=180                                        | Searches the given image on the screen and returns its center of X Y co-ordinates.                                                                                           |
+-------------------------------------------------------+----------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Mouse_Operations.gif?raw=true

***********************************************************
6 functions on Window Operations (works only in Windows OS)
***********************************************************


+--------------------------------------------------+-----------------------+-------------------------------------------------------+
| Function                                         | Accepted Parameters   | Description                                           |
+==================================================+=======================+=======================================================+
| cf.window\_show\_desktop()                       | None                  | Minimizes all the applications and shows Desktop.     |
+--------------------------------------------------+-----------------------+-------------------------------------------------------+
| cf.window\_get\_all\_opened\_titles\_windows()   | window\_title=" "     | Gives the title of all the existing (open) windows.   |
+--------------------------------------------------+-----------------------+-------------------------------------------------------+
| cf.window\_activate\_and\_maximize\_windows()    | windowName=" "        | Activates and maximizes the desired window.           |
+--------------------------------------------------+-----------------------+-------------------------------------------------------+
| cf.window\_minimize\_windows()                   | windowName=" "        | Activates and minimizes the desired window.           |
+--------------------------------------------------+-----------------------+-------------------------------------------------------+
| cf.window\_close\_windows()                      | windowName=" "        | Close the desired window.                             |
+--------------------------------------------------+-----------------------+-------------------------------------------------------+
| cf.launch\_any\_exe\_bat\_application()          | pathOfExeFile=" "     | Launches any exe or batch file or excel file etc.     |
+--------------------------------------------------+-----------------------+-------------------------------------------------------+

.. only:: html

   .. figure:: 'https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Functions%20Light%20GIFs/Window Operations.gif?raw=true'

********************************************************
5 functions on Window Objects (works only in Windows OS)
********************************************************


+------------------------------------+-------------------------------------------------------------------------------+--------------------------------------------------------------------+
| Function                           | Accepted Parameters                                                           | Description                                                        |
+====================================+===============================================================================+====================================================================+
| cf.win\_obj\_open\_app()           | title, program\_path\_with\_name, file\_path\_with\_name=" ", backend="uia"   | Open any windows application.                                      |
+------------------------------------+-------------------------------------------------------------------------------+--------------------------------------------------------------------+
| cf.win\_obj\_get\_all\_objects()   | main\_dlg, save=False, file\_name\_with\_path=" "                             | Print or Save all the windows object elements of an application.   |
+------------------------------------+-------------------------------------------------------------------------------+--------------------------------------------------------------------+
| cf.win\_obj\_mouse\_click()        | main\_dlg,title=" ", auto\_id=" ", control\_type=" "                          | Simulate high level mouse clicks on windows object elements.       |
+------------------------------------+-------------------------------------------------------------------------------+--------------------------------------------------------------------+
| cf.win\_obj\_key\_press()          | main\_dlg,write, title=" ", auto\_id=" ", control\_type=" "                   | Simulate high level key-press on windows object elements.          |
+------------------------------------+-------------------------------------------------------------------------------+--------------------------------------------------------------------+
| cf.win\_obj\_get\_text()           | main\_dlg, title=" ", auto\_id=" ", control\_type=" ", value = False          | Read text from windows object element.                             |
+------------------------------------+-------------------------------------------------------------------------------+--------------------------------------------------------------------+

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Windows_Object_Operation.gif?raw=true

********************************
8 functions on Folder Operations
********************************


+----------------------------------------------+----------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| Function                                     | Accepted Parameters                                            | Description                                                                                                                 |
+==============================================+================================================================+=============================================================================================================================+
| cf.folder\_read\_text\_file()                | txt\_file\_path=" "                                            | Reads from a given text file and returns entire contents as a single list                                                   |
+----------------------------------------------+----------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| cf.folder\_write\_text\_file()               | txt\_file\_path=" ", contents=" "                              | Writes given contents to a text file                                                                                        |
+----------------------------------------------+----------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| cf.folder\_create()                          | strFolderPath=" "                                              | When you are making leaf directory, if any intermediate-level directory is missing, folder\_create() method creates them.   |
+----------------------------------------------+----------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| cf.folder\_create\_text\_file()              | textFolderPath=" ", txtFileName=" "                            | Creates text file in the given path.                                                                                        |
+----------------------------------------------+----------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| cf.folder\_get\_all\_filenames\_as\_list()   | strFolderPath=" ", extension='all'                             | Get all the files of the given folder in a list.                                                                            |
+----------------------------------------------+----------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| cf.folder\_delete\_all\_files()              | fullPathOfTheFolder=" ", file\_extension\_without\_dot="all"   | Deletes all the files of the given folder                                                                                   |
+----------------------------------------------+----------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| cf.file\_rename()                            | old\_file\_path='', new\_file\_name='', ext=False              | Renames the given file name to new file name with same extension.                                                           |
+----------------------------------------------+----------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| cf.file\_get\_json\_details()                | path\_of\_json\_file='', section=''                            | Returns all the details of the given section in a dictionary                                                                |
+----------------------------------------------+----------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Folder_Operations.gif?raw=true

********************************
28 functions on Excel Operations
********************************
  

+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| Function                                             | Accepted Parameters                                                                                                                                                                                                                      | Description                                                                                                                                |
+======================================================+==========================================================================================================================================================================================================================================+============================================================================================================================================+
| cf.excel\_get\_all\_sheet\_names()                   | excelFilePath=" "                                                                                                                                                                                                                        | Gives you all names of the sheets in the given excel sheet.                                                                                |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_create\_excel\_file\_in\_given\_folder()   | fullPathToTheFolder=" ", excelFileName=" ", sheet\_name="Sheet1"                                                                                                                                                                         | Creates an excel file in the desired folder with desired filename                                                                          |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_if\_value\_exists()                        | excel\_path=" ", sheet\_name="Sheet1", header=0, usecols=" ", value=" "                                                                                                                                                                  | Check if a given value exists in given excel. Returns True / False                                                                         |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_create\_file()                             | fullPathToTheFile=" ", fileName=" ", sheet\_name="Sheet1"                                                                                                                                                                                | Create a Excel file in fullPathToTheFile with filename.                                                                                    |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_copy\_paste\_range\_from\_to\_sheet()      | excel\_path=" ", sheet\_name="Sheet1", startCol=0, startRow=0, endCol=0, endRow=0, copiedData=" "                                                                                                                                        | Pastes the copied data in specific range of the given excel sheet.                                                                         |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_get\_row\_column\_count()                  | excel\_path=" ", sheet\_name="Sheet1", header=0                                                                                                                                                                                          | Gets the row and column count of the provided excel sheet.                                                                                 |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_copy\_range\_from\_sheet()                 | excel\_path=" ", sheet\_name="Sheet1", startCol=0, startRow=0, endCol=0, endRow=0                                                                                                                                                        | Copies the specific range from the provided excel sheet and returns copied data as a list                                                  |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_split\_by\_column()                        | excel\_path=" ", sheet\_name="Sheet1", header=0, columnName=" "                                                                                                                                                                          | Splits the excel file by Column Name                                                                                                       |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_split\_the\_file\_on\_row\_count()         | excel\_path=" ", sheet\_name = "Sheet1", rowSplitLimit=" ", outputFolderPath=" ", outputTemplateFileName ="Split"                                                                                                                        | Splits the excel file as per given row limit                                                                                               |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_merge\_all\_files()                        | input\_folder\_path=" ", output\_folder\_path=" "                                                                                                                                                                                        | Merges all the excel files in the given folder                                                                                             |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_drop\_columns()                            | excel\_path=" ", sheet\_name="Sheet1", header=0, columnsToBeDropped = " "                                                                                                                                                                | Drops the desired column from the given excel file                                                                                         |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_sort\_columns()                            | excel\_path=" ", sheet\_name="Sheet1", header=0, firstColumnToBeSorted=None, secondColumnToBeSorted=None, thirdColumnToBeSorted=None, firstColumnSortType=True, secondColumnSortType=True, thirdColumnSortType=True, view\_excel=False   | A function which takes excel full path to excel and column names on which sort is to be performed                                          |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_clear\_sheet()                             | excel\_path=" ",sheet\_name="Sheet1", header=0                                                                                                                                                                                           | Clears the contents of given excel files keeping header row intact                                                                         |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_set\_single\_cell()                        | excel\_path=" ", sheet\_name="Sheet1", header=0, columnName=" ", cellNumber=0, setText=" "                                                                                                                                               | Writes the given text to the desired column/cell number for the given excel file                                                           |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_get\_single\_cell()                        | excel\_path=" ",sheet\_name="Sheet1", header=0, columnName=" ",cellNumber=0                                                                                                                                                              | Gets the text from the desired column/cell number of the given excel file                                                                  |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_remove\_duplicates()                       | excel\_path=" ",sheet\_name="Sheet1", header=0, columnName=" ", saveResultsInSameExcel=True, which\_one\_to\_keep="first"                                                                                                                | Drops the duplicates from the desired Column of the given excel file                                                                       |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_vlook\_up()                                | filepath\_1=" ", sheet\_name\_1 = "Sheet1", header\_1 = 0, filepath\_2=" ", sheet\_name\_2 = "Sheet1", header\_2 = 0, Output\_path=" ", OutputExcelFileName=" ", match\_column\_name=" ", how='left', view\_excel=False                  | Performs excel\_vlook\_up on the given excel files for the desired columns. Possible values for how are "inner","left", "right", "outer"   |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_describe\_data()                           | excel\_path=" ",sheet\_name="Sheet1", header=0, view\_excel=False                                                                                                                                                                        | Describe statistical data for the given excel                                                                                              |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_change\_corrupt\_xls\_to\_xlsx()           | xls\_file ='',xlsx\_file = '', xls\_sheet\_name=''                                                                                                                                                                                       | Repair corrupt excel file                                                                                                                  |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_get\_all\_header\_columns()                | excel\_path=" ",sheet\_name="Sheet1",header=0                                                                                                                                                                                            | Gives you all column header names of the given excel sheet                                                                                 |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_convert\_to\_image()                       | excel\_file\_path=" "                                                                                                                                                                                                                    | Returns an Image (PNG) of given Excel                                                                                                      |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_split\_on\_user\_defined\_conditions()     | excel\_file\_path, sheet\_name="Sheet1", column\_name='', condition\_strings=None,output\_dir='', view\_excel=False                                                                                                                      | Splits the excel based on user defined row/column conditions                                                                               |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_apply\_format\_as\_table()                 | excel\_file\_path, table\_style="TableStyleMedium21", sheet\_name="Sheet1"                                                                                                                                                               | Applies table format to the used range of the given excel                                                                                  |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_convert\_xls\_to\_xlsx()                   | xls\_file\_path='',xlsx\_file\_path=''                                                                                                                                                                                                   | Converts given XLS file to XLSX                                                                                                            |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.isNaN()                                           | value                                                                                                                                                                                                                                    | Returns TRUE if a given value is NaN False otherwise                                                                                       |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.convert\_csv\_to\_excel()                         | csv\_path=" ", sep=" "                                                                                                                                                                                                                   | Function to convert CSV to Excel                                                                                                           |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_sub\_routines()                            | None                                                                                                                                                                                                                                     | Excel VBA Macros called from ClointFusion                                                                                                  |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+
| cf.excel\_to\_colored\_html()                        | formatted\_excel\_path=" "                                                                                                                                                                                                               | Converts given Excel to HTML preserving the Excel format and saves in same folder as .html                                                 |
+------------------------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------+

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Excel_Operations.gif?raw=true

**********************************
3 functions on Keyboard Operations
**********************************


+--------------------------+-------------------------------------------------------------------------------+------------------------------------------------------------------------+
| Function                 | Accepted Parameters                                                           | Description                                                            |
+==========================+===============================================================================+========================================================================+
| cf.key\_hit\_enter()     | write\_to\_window=" "                                                         | Enter key will be pressed once.                                        |
+--------------------------+-------------------------------------------------------------------------------+------------------------------------------------------------------------+
| cf.key\_press()          | key\_1='', key\_2='', key\_3='', write\_to\_window=" "                        | Emulates the given keystrokes.                                         |
+--------------------------+-------------------------------------------------------------------------------+------------------------------------------------------------------------+
| cf.key\_write\_enter()   | text\_to\_write=" ", write\_to\_window=" ", delay\_after\_typing=1, key="e"   | Writes/Types the given text and press enter (by default) or tab key.   |
+--------------------------+-------------------------------------------------------------------------------+------------------------------------------------------------------------+

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/KB_Operations.gif?raw=true

*****************************************
5 functions on Screen-scraping Operations
*****************************************

+-------------------------------------------------------+---------------------------------------------------------------------+-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| Function                                              | Accepted Parameters                                                 | Description                                                                                                                                                                   |
+=======================================================+=====================================================================+===============================================================================================================================================================================+
| cf.scrape\_save\_contents\_to\_notepad()              | folderPathToSaveTheNotepad=" ", switch\_to\_window=" ",X=0, Y=0     | Copy pastes all the available text on the screen to notepad and saves it.                                                                                                     |
+-------------------------------------------------------+---------------------------------------------------------------------+-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.scrape\_get\_contents\_by\_search\_copy\_paste()   | highlightText=" "                                                   | Gets the focus on the screen by searching given text using crtl+f and performs copy/paste of all data. Useful in Citrix applications. This is useful in Citrix applications   |
+-------------------------------------------------------+---------------------------------------------------------------------+-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.screen\_clear\_search()                            | delay=0.2                                                           | Clears previously found text (crtl+f highlight)                                                                                                                               |
+-------------------------------------------------------+---------------------------------------------------------------------+-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.search\_highlight\_tab\_enter\_open()              | searchText=" ", hitEnterKey="Yes", shift\_tab='No'                  | Searches for a text on screen using crtl+f and hits enter. This function is useful in Citrix environment.                                                                     |
+-------------------------------------------------------+---------------------------------------------------------------------+-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.find\_text\_on\_screen()                           | searchText=" ", delay=0.1, occurance=1, isSearchToBeCleared=False   | Clears previous search and finds the provided text on screen.                                                                                                                 |
+-------------------------------------------------------+---------------------------------------------------------------------+-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Screen_Scraping.gif?raw=true

**********************************
11 functions on Browser Operations
**********************************


+------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------+-------------------------------------------------------------------+
| Function                           | Accepted Parameters                                                                                                                                        | Description                                                       |
+====================================+============================================================================================================================================================+===================================================================+
| cf.browser\_activate()             | url=" ", files\_download\_path='', dummy\_browser=True, open\_in\_background=False, incognito=False, clear\_previous\_instances=False, profile="Default"   | Function to launch browser and start the session.                 |
+------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------+-------------------------------------------------------------------+
| cf.browser\_navigate\_h()          | url=" "                                                                                                                                                    | Navigates to Specified URL.                                       |
+------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------+-------------------------------------------------------------------+
| cf.browser\_write\_h()             | Value=" ", User\_Visible\_Text\_Element=" "                                                                                                                | Write a string on the given element.                              |
+------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------+-------------------------------------------------------------------+
| cf.browser\_mouse\_click\_h()      | User\_Visible\_Text\_Element=" ", element=" ", double\_click=False, right\_click=False                                                                     | Click on the given element.                                       |
+------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------+-------------------------------------------------------------------+
| cf.browser\_locate\_element\_h()   | selector=" ", get\_text=False, multiple\_elements=False                                                                                                    | Find the element by Xpath, id or css selection.                   |
+------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------+-------------------------------------------------------------------+
| cf.browser\_wait\_until\_h()       | text=" ", element="t"                                                                                                                                      | Wait until a specific element is found.                           |
+------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------+-------------------------------------------------------------------+
| cf.browser\_refresh\_page\_h()     | None                                                                                                                                                       | Refresh the page.                                                 |
+------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------+-------------------------------------------------------------------+
| cf.browser\_quit\_h()              | None                                                                                                                                                       | Close the Helium browser.                                         |
+------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------+-------------------------------------------------------------------+
| cf.browser\_hit\_enter\_h()        | None                                                                                                                                                       | Hits enter KEY using Browser Helium Functions                     |
+------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------+-------------------------------------------------------------------+
| cf.browser\_key\_press\_h()        | key\_1=" ", key\_2=" "                                                                                                                                     | Type text using Browser Helium Functions and press hot keys       |
+------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------+-------------------------------------------------------------------+
| cf.browser\_mouse\_hover\_h()      | User\_Visible\_Text\_Element=" "                                                                                                                           | Performs a Mouse Hover over the Given User Visible Text Element   |
+------------------------------------+------------------------------------------------------------------------------------------------------------------------------------------------------------+-------------------------------------------------------------------+

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Functions%20Light%20GIFs/browser_functions.gif?raw=true

*****************************
4 functions on Alert Messages
*****************************
   

+--------------------------------------+-----------------------------------------------------------------------+----------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| Function                             | Accepted Parameters                                                   | Description                                                                                                                                                          |
+======================================+=======================================================================+======================================================================================================================================================================+
| cf.message\_counter\_down\_timer()   | strMsg="Calling ClointFusion Function in (seconds)", start\_value=5   | Function to show count-down timer. Default is 5 seconds.                                                                                                             |
+--------------------------------------+-----------------------------------------------------------------------+----------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.message\_pop\_up()                | strMsg=" ", delay=3                                                   | Specified message will popup on the screen for a specified duration of time.                                                                                         |
+--------------------------------------+-----------------------------------------------------------------------+----------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.message\_flash()                  | msg=" ", delay=3                                                      | Specified msg will popup for a specified duration of time with OK button.                                                                                            |
+--------------------------------------+-----------------------------------------------------------------------+----------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.message\_toast()                  | message,website\_url=" ", file\_folder\_path=" "                      | Function for displaying Windows 10 Toast Notifications. Pass website URL OR file / folder path that needs to be opened when user clicks on the toast notification.   |
+--------------------------------------+-----------------------------------------------------------------------+----------------------------------------------------------------------------------------------------------------------------------------------------------------------+


********************************
3 functions on String Operations
********************************


+--------------------------------------------+-----------------------+--------------------------------------------------+
| Function                                   | Accepted Parameters   | Description                                      |
+============================================+=======================+==================================================+
| cf.string\_remove\_special\_characters()   | inputStr=" "          | Removes all the special character.               |
+--------------------------------------------+-----------------------+--------------------------------------------------+
| cf.string\_extract\_only\_alphabets()      | inputString=" "       | Returns only alphabets from given input string   |
+--------------------------------------------+-----------------------+--------------------------------------------------+
| cf.string\_extract\_only\_numbers()        | inputString=" "       | Returns only numbers from given input string     |
+--------------------------------------------+-----------------------+--------------------------------------------------+

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/String_Operations.gif?raw=true


***********************
Miscellaneous functions 
***********************

+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| Function                                  | Accepted Parameters                                                                                | Description                                                                                                                                                                                            |
+===========================================+====================================================================================================+========================================================================================================================================================================================================+
| cf.clear\_screen()                        | None                                                                                               | Clears Python Interpreter Terminal Window Screen                                                                                                                                                       |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.print\_with\_magic\_color()            | strMsg:str=" ", magic:bool=False                                                                   | Function to color and format terminal output                                                                                                                                                           |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.schedule\_create\_task\_windows()      | Weekly\_Daily="D", week\_day="Sun", start\_time\_hh\_mm\_24\_hr\_frmt="11:00"                      | Schedules (weekly & daily options as of now) the current BOT (.bat) using Windows Task Scheduler. Please call create\_batch\_file() function before using this function to convert .pyw file to .bat   |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.schedule\_delete\_task\_windows()      | None                                                                                               | Deletes already scheduled task. Asks user to supply task\_name used during scheduling the task. You can also perform this action from Windows Task Scheduler.                                          |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.show\_emoji()                          | strInput=" "                                                                                       | Function which prints Emojis                                                                                                                                                                           |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.dismantle\_code()                      | strFunctionName=" "                                                                                | This functions dis-assembles given function and shows you column-by-column summary to explain the output of disassembled bytecode.                                                                     |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.ON\_semi\_automatic\_mode()            | None                                                                                               | This function sets semi\_automatic\_mode as True => ON                                                                                                                                                 |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.OFF\_semi\_automatic\_mode()           | None                                                                                               | This function sets semi\_automatic\_mode as False => OFF                                                                                                                                               |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.email\_send\_via\_desktop\_outlook()   | toAddress=" ", ccAddress=" ", subject=" ",htmlBody=" ", embedImgPath=" ", attachmentFilePath=" "   | Send email using Outlook from Desktop email application                                                                                                                                                |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.download\_this\_file()                 | url=" "                                                                                            | Downloads a given url file to BOT output folder or Browser's Download folder                                                                                                                           |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.pause\_program()                       | seconds="5"                                                                                        | Stops the program for given seconds                                                                                                                                                                    |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.string\_regex()                        | inputStr=" ", strExpAfter=" ", strExpBefore=" ", intIndex=0                                        | Regex API service call, to search within a given string data                                                                                                                                           |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.ocr\_now()                             | img\_path=" "                                                                                      | Recognize and read the text embedded in images using Google's Tesseract-OCR                                                                                                                            |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.update\_log\_excel\_file()             | message=" "                                                                                        | Given message will be updated in the excel log file of output folder                                                                                                                                   |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
| cf.create\_batch\_file()                  | application\_exe\_pyw\_file\_path=" "                                                              | Creates .bat file for the given application / exe or even .pyw BOT developed by you. This is required in Task Scheduler.                                                                               |
+-------------------------------------------+----------------------------------------------------------------------------------------------------+--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/miscallaneous.gif?raw=true
    

=============================
BOTS made out of ClointFusion
=============================


1) Outlook Email BOT implemented using ClointFusion

.. only:: html

   .. figure:: https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Functions%20Light%20GIFs/Gmail_and_Outlook_BOT.gif?raw=true


2) Others
   

**Date ❤️ with ClointFusion**

This an initiative for fast track entry into our growing workforce. For
more details, please visit: `Date with
ClointFusion <https://lnkd.in/gh_r9YB>`__

*We love your contribution*

Contribute to us by giving a star, writing articles on ``ClointFusion``,
giving comments, reporting bugs, bug fixes, feature enhancements, adding
documentation, and many other ways.

*Invitation to our Monthly Branded Hackathon*

We also invite everyone to take part in our monthly branded event, the
``ClointFusion Hackathon``, and stand a chance to work with us.

Checkout our Hackathon Website for more details here: `ClointFusion
Hackathon <https://sites.google.com/view/clointfusion-hackathon>`__


===========================
Need help in Building BOTS?
===========================

Write us with your requirments at ClointFusion@cloint.com

================
Acknowledgements
================

We sincerely thanks to all it's dependent packages for the great
contribution, which made ``ClointFusion`` possible!

Please find all the dependencies
`here <https://openbase.com/python/ClointFusion/dependencies>`__

Credits:

*ReadMe File Maintainer*

fharookshaik, Intern @ ClointFusion. Incase of any queries reach him on: https://www.linkedin.com/in/fharook-shaik-7a757b181/


.. |PyPI| image:: https://img.shields.io/pypi/v/ClointFusion?label=PyPI%20Version
.. |PyPI - License| image:: https://img.shields.io/pypi/l/ClointFusion?label=License
.. |PyPI - Status| image:: https://img.shields.io/pypi/status/ClointFusion?label=Release%20Status
.. |ClointFusion| image:: https://snyk.io/advisor/python/ClointFusion/badge.svg
.. |PyPI - Downloads| image:: https://img.shields.io/pypi/dm/ClointFusion?label=PyPI%20Downloads
.. |Libraries.io SourceRank| image:: https://img.shields.io/librariesio/sourcerank/pypi/ClointFusion
.. |PyPI - Format| image:: https://img.shields.io/pypi/format/ClointFusion?label=PyPI%20Format
.. |GitHub contributors| image:: https://img.shields.io/github/contributors/ClointFusion/ClointFusion?label=Contributors
.. |GitHub last commit| image:: https://img.shields.io/github/last-commit/ClointFusion/ClointFusion?label=Last%20Commit
.. |GitHub Repo stars| image:: https://img.shields.io/github/stars/ClointFusion/ClointFusion?label=Stars&style=social
.. |Twitter URL| image:: https://img.shields.io/twitter/url?style=social&url=https%3A%2F%2Ftwitter.com%2FClointFusion
.. |YouTube Channel Subscribers| image:: https://img.shields.io/youtube/channel/subscribers/UCIygBtp1y_XEnC71znWEW2w?style=social
.. |Twitter Follow| image:: https://img.shields.io/twitter/follow/ClointFusion?style=social

