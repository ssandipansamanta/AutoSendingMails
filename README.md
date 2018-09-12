Sending Email to Multiple Recipients with attachments
===================================================

## Description
This code will send mails to multiple person with same message and attachment. 
The usability would be to send a status mail to all team members. 
Currently, it's using gmail but can be modified for others as well.


## Approach
User need to mention email ids of recipeints in a excel file. 
Currently it present in Inputs folder viz. __NameList__. 
The file needs to be also kept in Inputs folder to be attached. 
__welcome-image__ file is used in the current set up to attach with mail. 

## Steps to Run the Code

1. Install the required jars.
2. Keep the excel file with Email Ids.
3. Keep the file need to be attached.
4. Open Config file and change the File names, Email sender details etc.

## Config File Fields and Implementation

* NameListFile - Name of the Excel File. ___Refer NameList.xlxs file in Inputs folder___
* SheetPositionwithinFile - Sheet position. ___Mention 0 if Ids are present in Sheet 1___
* IDPositionwithinSheet - Field position where Ids are present. ___Mentioned 2 if Ids are present in col-C___
* NameAttachFile - Name of the file to be attached. ___'welcome-impage.jpg' for current set up___
* EmailFrom - Sender email id
* Password - Sender password

## Requirements
```xml
1. poi
2. javax.mail
3. javax.Json
```

  