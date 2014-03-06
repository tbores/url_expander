url_expander
============

your.xls -> | url_expander.py | -> OUT_your.xls

Specification
-------------
URL Expander has been created for supporting a friend writing his Phd.
He wanted me to analyse Excel files generating with http://twitonomy.com/.
Each file contains a tweet in each row at the third column.
Goal of url_expander is to replace all shortened URL that can be found in a tweet with the real url.
url_expander doesn't change your twitonomy excel files. Instead it creates for each a new one named OUT_orginalfilename.

Usage:
------
The path to the folder containing the excel files is hardcoded. Default value is "./excels/".
So you just have to create a folder named "excels" and to put all your twitonomy excels files in it.
Then run "python url_expander.py" from the parent folder.

Dependencies
------------
- Python2.7
- xlrd (http://www.python-excel.org/)
- xlwt (http://www.python-excel.org/)
- xlutils (http://www.python-excel.org/)
- Expandulr online service (http://expandurl.appspot.com/)
