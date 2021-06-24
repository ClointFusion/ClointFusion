Welcome to <img src="https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Cloint-LOGO.png" height="30"> , Made in India with &#10084;&#65039;
==============================================================================================================================================================

<img src="https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/CCE.PNG">
========================================================================================

Description:
============

Cloint India Pvt. Ltd - Python functions for Automation (RPA)

What is ClointFusion ?
======================

ClointFusion is a company registered at Vadodara, Gujarat, India. ClointFusion is our home-grown Python based RPA platform for Software BOT development. We are working towards Common Man’s RPA using AI. 

Welcome to ClointFusion, Made in India with &#10084;&#65039; 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

![PyPI](https://img.shields.io/pypi/v/ClointFusion?label=PyPI%20Version) ![PyPI - License](https://img.shields.io/pypi/l/ClointFusion?label=License) ![PyPI - Status](https://img.shields.io/pypi/status/ClointFusion?label=Release%20Status) ![ClointFusion](https://snyk.io/advisor/python/ClointFusion/badge.svg) ![PyPI - Downloads](https://img.shields.io/pypi/dm/ClointFusion?label=PyPI%20Downloads) ![Libraries.io SourceRank](https://img.shields.io/librariesio/sourcerank/pypi/ClointFusion) ![PyPI - Format](https://img.shields.io/pypi/format/ClointFusion?label=PyPI%20Format) ![GitHub contributors](https://img.shields.io/github/contributors/ClointFusion/ClointFusion?label=Contributors) ![GitHub last commit](https://img.shields.io/github/last-commit/ClointFusion/ClointFusion?label=Last%20Commit) 

![GitHub Repo stars](https://img.shields.io/github/stars/ClointFusion/ClointFusion?label=Stars&style=social) ![Twitter URL](https://img.shields.io/twitter/url?style=social&url=https%3A%2F%2Ftwitter.com%2FClointFusion) ![YouTube Channel Subscribers](https://img.shields.io/youtube/channel/subscribers/UCIygBtp1y_XEnC71znWEW2w?style=social) ![Twitter Follow](https://img.shields.io/twitter/follow/ClointFusion?style=social)

Test Drive ClointFusion on Google Colabs
========================================

<a href='https://colab.research.google.com/github/ClointFusion/ClointFusion/blob/master/ClointFusion_Labs.ipynb' target="_blank"><img src='https://colab.research.google.com/assets/colab-badge.svg' alt="Open In Colab\"/></a>
================================================================================================================================================================================================================================

Installation on your local computer
===================================

ClointFusion is now supported in Windows / Ubuntu / macOS !
===========================================================

1. Please install Python 3.8.5 with 64 bit: Python 3.8.5 64 Bit

2. It is recommended to run ClointFusion in a Virtual Environment.
   Please refer these steps to create one, as per your OS: Creating a
   virtual environment in Windows / Mac / Ubuntu

3. Install ClointFusion by executing this package in command promt (with
   Admin rights):

pip install --upgrade ClointFusion
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

4. Open a new file in your favorite Python IDE and type:

import ClointFusion as cf
~~~~~~~~~~~~~~~~~~~~~~~~~

PS: Ubuntu users: May need to install some additional packages: 
1) sudo apt-get install python3-tk python3-dev
2) sudo apt-get install -y fonts-symbola
3) sudo apt-get install scrot 
4) sudo apt-get install libcairo2-dev libjpeg-dev libgif-dev
5) sudo apt-get install libgirepository1.0-dev
6) sudo apt-get install python3-apt
7) sudo apt-get install  python3-xlib

Now access more than 130 functions (hit ctrl+space in your IDE)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

-  6 gui functions, to take any input from user:

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   cf.gui\_get\_any\_file\_from\_user() : Generic function to accept
   file path from user using GUI. Returns the filepath value in string
   format.Default allows all files.

   cf.gui\_get\_consent\_from\_user() : Generic function to get consent
   from user using GUI. Returns the string 'yes' or 'no'

   cf.gui\_get\_dropdownlist\_values\_from\_user() : Generic function to
   accept one of the drop-down value from user using GUI. Returns all
   chosen values in list format.

   cf.gui\_get\_excel\_sheet\_header\_from\_user() : Generic function to
   accept excel path, sheet name and header from user using GUI. Returns
   all these values in disctionary format.

   cf.gui\_get\_folder\_path\_from\_user() : Generic function to accept
   folder path from user using GUI. Returns the folderpath value in
   string format.

   cf.gui\_get\_any\_input\_from\_user() : Generic function to accept
   any input (text / numeric) from user using GUI. Returns the value in
   string format.

-  8 functions on Mouse operations:

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   cf.mouse\_click() : Clicks at the given X Y Co-ordinates on the
   screen using ingle / double / tripple click(s). Optionally copies
   selected data to clipboard (works for double / triple clicks)

   cf.mouse\_move() : Moves the cursor to the given X Y Co-ordinates

   cf.mouse\_get\_color\_by\_position() : Gets the color by X Y
   co-ordinates of the screen

   cf.mouse\_drag\_from\_to() : Clicks and drags from X1 Y1 co-ordinates
   to X2 Y2 Co-ordinates on the screen

   cf.mouse\_search\_snip\_return\_coordinates\_x\_y() : Searches the
   given image on the screen and returns its center of X Y co-ordinates.

   cf.mouse\_search\_snips\_return\_coordinates\_x\_y() : Searches the
   given set of images on the screen and returns its center of X Y
   co-ordinates of FIRST OCCURANCE

   cf.mouse\_search\_snip\_return\_coordinates\_box() : Searches the
   given image on the screen and returns the 4 bounds co-ordinates
   (x,y,w,h)

   cf.mouse\_find\_highlight\_click() : Searches the given text on the
   screen, highlights and clicks it

-  5 functions on Window operations (works only in Windows OS):

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   cf.window\_show\_desktop() : Minimizes all the applications and shows
   Desktop.

   cf.window\_get\_all\_opened\_titles\_windows() : Gives the title of
   all the existing (open) windows.

   cf.window\_activate\_and\_maximize\_windows() : Activates and
   maximizes the desired window.

   cf.window\_minimize\_windows() : Activates and minimizes the desired
   window.

   cf.window\_close\_windows() : Close the desired window.

-  6 functions on Folder operations:

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   cf.folder\_read\_text\_file() : Reads from a given text file and
   returns entire contents as a single list

   cf.folder\_write\_text\_file() : Writes given contents to a text file

   cf.folder\_create() : while making leaf directory if any
   intermediate-level directory is missing, folder\_create() method will
   create them all.

   cf.folder\_create\_text\_file() : Creates Text file in the given
   path.

   cf.folder\_get\_all\_filenames\_as\_list() : Get all the files of the
   given folder in a list.

   cf.folder\_delete\_all\_files() : Deletes all the files of the given
   folder

-  20 functions on Excel operations:

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   cf.excel\_get\_all\_sheet\_names() : Gives you all names of the
   sheets in the given excel sheet.

   cf.excel\_create\_cf.excel\_file\_in\_given\_folder()

   cf.excel\_if\_value\_exists() : Check if a given value exists in
   given excel. Returns True / False

   cf.excel\_create\_file()

   cf.excel\_copy\_paste\_range\_from\_to\_sheet() : Pastes the copied
   data in specific range of the given excel sheet.

   cf.excel\_get\_row\_column\_count() : Gets the row and coloumn count
   of the provided excel sheet.

   cf.excel\_copy\_range\_from\_sheet() : Copies the specific range from
   the provided excel sheet and returns copied data as a list

   cf.excel\_split\_by\_column() : Splits the excel file by Column Name

   cf.excel\_split\_the\_file\_on\_row\_count() : Splits the excel file
   as per given row limit

   cf.excel\_merge\_all\_files() : Merges all the excel files in the
   given folder

   cf.excel\_drop\_columns() : Drops the desired column from the given
   excel file

   cf.excel\_sort\_columns() : A function which takes excel full path to
   excel and column names on which sort is to be performed

   cf.excel\_clear\_sheet() : Clears the contents of given excel files
   keeping header row intact

   cf.excel\_set\_single\_cell() : Writes the given text to the desired
   column/cell number for the given excel file

   cf.excel\_get\_single\_cell() : Gets the text from the desired
   column/cell number of the given excel file

   cf.excel\_remove\_duplicates() : Drops the duplicates from the
   desired Column of the given excel file

   cf.excel\_vlook\_up() : Performs excel\_vlook\_up on the given excel
   files for the desired columns. Possible values for how are
   "inner","left", "right", "outer"

   cf.excel\_draw\_charts() : Interactive data visualization function,
   which accepts excel file, X & Y column. Chart types accepted are bar
   , scatter , pie , sun , histogram , box , strip. You can pass color
   column as well, having a boolean value.

   cf.excel\_describe\_data() : Describe statistical data for the given
   excel

-  3 functions on Keyboard operations:

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   cf.key\_hit\_enter() : Enter key will be pressed once.

   cf.key\_press() : Emulates the given keystrokes.

   cf.key\_write\_enter() : Writes/Types the given text and press enter
   (by default) or tab key.

-  2 functions on Screenscraping operations:

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   cf.scrape\_save\_contents\_to\_notepad : Copy pastes all the
   available text on the screen to notepad and saves it.

   cf.scrape\_get\_contents\_by\_search\_copy\_paste : Gets the focus on
   the screen by searching given text using crtl+f and performs
   copy/paste of all data. Useful in Citrix applications. This is useful
   in Citrix applications

-  12 functions on Browser operations:

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   cf.browser\_get\_html\_text() : Function to get HTML text without
   tags using Beautiful soup

   cf.browser\_get\_html\_tabular\_data\_from\_website() : Web Scrape
   HTML Tables : Gets Website Table Data Easily as an Excel using
   Pandas. Just pass the URL of Website having HTML Tables.

   cf.browser\_navigate\_h() : Navigates to Specified URL.

   cf.browser\_write\_h() : Write a string on the given element.

   cf.browser\_mouse\_click\_h() : Click on the given element.

   cf.browser\_mouse\_double\_click\_h() : Doubleclick on the given
   element.

   cf.browser\_locate\_element\_h() : Find the element by Xpath, id or
   css selection.

   cf.browser\_locate\_elements\_h() : Find the elements by Xpath, id or
   css selection.

   cf.browser\_wait\_until\_h() : Wait until a specific element is
   found.

   cf.browser\_refresh\_page\_h() : Refresh the page.

   cf.browser\_quit\_h() : Close the Helium browser.

   cf.browser\_hit\_enter\_h() : Hits enter KEY using Browser Helium
   Functions

-  3 functions on Alert Messages:

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   cf.message\_counter\_down\_timer() : Function to show count-down
   timer. Default is 5 seconds.

   cf.message\_pop\_up() : Specified message will popup on the screen
   for a specified duration of time.

   cf.message\_flash() : Specified msg will popup for a specified
   duration of time with OK button.

-  3 functions on String Operations:

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   cf.string\_remove\_special\_characters() : Removes all the special
   character.

   cf.string\_extract\_only\_alphabets() : Returns only alphabets from
   given input string

   cf.string\_extract\_only\_numbers() : Returns only numbers from given
   input string

-  Loads of miscellaneous functions related to emoji, capture photo, flash (pop-up) messages etc:

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   cf.launch\_any\_exe\_bat\_application() : Launches any exe or batch
   file or excel file etc.

   cf.launch\_website\_h() : Internal function to launch browser.

   cf.schedule\_create\_task\_windows() : Schedules (weekly & daily
   options as of now) the current BOT (.bat) using Windows Task
   Scheduler. Please call create\_batch\_file() function before using
   this function to convert .pyw file to .bat

   cf.schedule\_delete\_task\_windows() : Deletes already scheduled
   task. Asks user to supply task\_name used during scheduling the task.
   You can also perform this action from Windows Task Scheduler.

   cf.show\_emoji() : Function which prints Emojis

   cf.message\_counter\_down\_timer() : Function to show count-down
   timer. Default is 5 seconds.

   cf.get\_long\_lat() : Function takes zip\_code as input (int) and
   returns longitude, latitude, state, city, county.

   cf.dismantle\_code() : This functions dis-assembles given function
   and shows you column-by-column summary to explain the output of
   disassembled bytecode.

   cf.ON\_semi\_automatic\_mode() : This function sets
   semi\_automatic\_mode as True => ON

   cf.OFF\_semi\_automatic\_mode() : This function sets
   semi\_automatic\_mode as False => OFF

   cf.camera\_capture\_image() : turn ON camera & take photo

   cf.convert\_csv\_to\_excel() : Function to convert CSV to Excel

   cf.capture\_snip\_now() : Captures the snip and stores in Image
   Folder of the BOT by giving continous numbering

   cf.take\_error\_screenshot() : Takes screenshot of an error popup
   parallely without waiting for the flow of the program. The screenshot
   will be saved in the log folder for reference.

   cf.find\_text\_on\_screen() : Clears previous search and finds the
   provided text on screen.

ClointFusion's function works in different modes:
=================================================

1) If you pass all the required parameters, function works silently. So,
   this is expert (Non-GUI) mode. This mode gives you more control over
   the function's parameters.

2) If you do not pass any parameter, GUI would pop-up asking you the required parameters. Next time, when you run the BOT, based upon your configuration, which you get to choose at the beginning of BOT run:

       A) If Semi-Automatic mode is OFF, GUI would pop-up again, showing you the previous entries, allowing you to modify the parameters.

       B) If Semi-Automatic mode in ON, BOT works silently taking your previous GUI entries.

   GUI Mode is for beginners. Anytime, if you are not getting how to use
   the function, just call an empty function (without parameters) and
   GUI would pop-up asking you for required parameters.

We love your contribution
=========================

Contribute by giving a star / writing article on ClointFusion / feedback
/ report issues / bug fixes / feature enhancement / add documentation /
many more ways as you please..

Participate in our monthly online hackathons & weekly meetups. Click
here for more details: https://sites.google.com/view/clointfusion-hackathon

Please visit our GitHub repository:
https://github.com/ClointFusion/ClointFusion

Contact us:
===========

Drop a mail to ClointFusion@cloint.com
