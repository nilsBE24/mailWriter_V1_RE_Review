 =================================================
 |                                               |
 |     |------------------|                      |
 |     |\_              _/|                      |
 |     |  \_          _/  |                      |
 |     |    \__    __/    |     E-Mail Writer    |
 |     |       \__/       |        V1 (RE)       |
 |     |                  |                      |
 |     |                  |                      |
 |     |------------------|                      |
 |                                               |
 =================================================
 |    Welcome to E-Mail Writer documentation!    |
 =================================================

© Nils Huber 2023

Version 1 RE (Review Edition)

================
| Installation |
================
1. Unzip the contents of the .zip file.
2. Make sure all components are in the same folder and not inside of a sub folder or similar.
3. Install Python3 if you have not already done so.
4. Install the following dependencies by using "pip install": pandas; openpyxl.


=============
| Execution |
=============
1. Edit mailWriterSettings.py first (please remember to change the senders @gmail.com address and app password for that address.)
2. Go to the folder directory and open the Terminal
3. Run mailWriter_v1.py.

============================
| Creating an app password |
============================

Tutorial: 
https://support.google.com/accounts/answer/185833?hl=de

============================
| Creating your .xlsx file |
============================
The following prerequisites must be met by any .xlsx file used by this program:
- The file must have one column filled exclusively with E-mails. (These can be any email address.)
- The file must have one column filled with names of the recipiants.
- The file must have one column filled with the gender of the recipiants. Should you wish to not provide a gender, please assign a random, empty column.

The program automatically assumes that:
-the addresses are in column A
-the names are in column B
-the genders are in column D

Should this not be the case, please change the values in mailWriterSettings.py.

You may optionally define a company name in column C or any column you choose.