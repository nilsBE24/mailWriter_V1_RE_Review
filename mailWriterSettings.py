
#  ====================
#  | LOGIN PARAMETERS |
#  ====================

gmail = "example@gmail.com" # Please enter the sender's E-mail address. Only @gmail.com addresses are accepted. Change this, as the email address filled in does not exist.
passkey = "examplePass"  # Enter the app password for the sender E-mail. Refer to "readme.txt" for more info.

#  ====================
#  | Program Settings |
#  ====================

fileLocation = "test.xlsx" # Please enter your file name here. Make sure it is in the same folder as mailWriter.py
status = True # Change this to show progress statuses. True = Show status; False = Do not show status. (Default = True)

#  ====================
#  | Content Settings |
#  ====================

addText = True # Choose if the E-Mail you are writing should contain plain text. (Default = True)
addSignature = True # Choose if the E-Mail has a signature (e.g. name, address, etc.) attached. (Default = True)
addImage = False # Not yet introduced
addFile = False # Not yet introduced

#  =================
#  | Mail Settings |
#  =================

subject = "Test" #the subject of your mail
message = "This is a test mail" # Please enter your desired E-Mail plain text message here.
signature = "This is a test signature" # Your personal signature goes here.

#  =======================
#  | .xlsx File Settings |
#  =======================

mailAmount = 0 # the amount of emails you would like to generate. If this value is zero, the program will automatically detect the amount.
mailAdress = 1 # the column where your mail addresses are stored (A=1; B=2; C=3 ...) (value can not be smaller than 1)
mailName = 2 #  the column where your recipients names are stored (A=1; B=2; C=3 ...) (value can not be smaller than 1)
mailCompany = 3 # the column where your recipients company names are stored (A=1; B=2; C=3 ...) (value can not be smaller than 1)
mailGender = 4 # the column where your recipients genders are stored (A=1; B=2; C=3 ...); (if there are no genders available, pick a random number with an empty column in the .xlsx file) (value can not be smaller than 1)
