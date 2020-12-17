# ExtractMsgContent
This PowerShell script extracts content from *.msg files and Active Directory, then writes that data to a CSV output file


1.	The script will iterate over msg files in a specified folder, defined by $msgPath
2.	Outlook must be closed… the script uses Outlook via a com object
3.	It creates a csv file with a randomly-generated filename in the source path to store information about each msg file
4.	The main script loops over each msg file in the source, extracting specific data from the message header
5.	The main script queries Active Directory for specific information about a user, and returns it in the $account variable
6.	Each set of information is added to an array as an object.  Array = @()  Object = @{}
7.	When the loop is complete, the script attempts to force close Outlook
8.	The array hash is exported to a .csv file


Sources:

https://mcpmag.com/articles/2017/06/08/creating-csv-files-with-powershell.aspx
http://vcloud-lab.com/entries/powershell/microsoft-powershell-generate-random-anything-filename--temppath--guid--password-
https://stackoverflow.com/questions/24074205/convertto-csv-output-without-quotes
https://stackoverflow.com/questions/47264561/how-to-get-email-address-from-the-emails-inside-an-oulook-folder-via-powershell
https://stackoverflow.com/questions/49693850/is-it-possible-to-extract-recipient-email-address-from-a-msg-file
https://stackoverflow.com/questions/43618494/get-contents-of-msg-file-into-string
https://stackoverflow.com/questions/17154825/renaming-msg-files-using-powershell
https://stackoverflow.com/questions/37932647/parse-body-of-msg-email-file
https://stackoverflow.com/questions/1954203/timestamp-on-file-name-using-powershell
