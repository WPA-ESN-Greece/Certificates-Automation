# Certificates-Automation
An automated way to generate and send certificates of an event.

---

[Here](https://docs.google.com/spreadsheets/d/1qqix5plIvhFJW3dv1L45dQ7M_nlUvUWsx8BSsBd9M48/template/preview), you will find the template link to create your copy. 

After creating your copy, you can go to the settings sheet and click the Generate Template button. this will create a Google Slide file in the same folder as the Spreadsheet, which is going to be the template for the certificates.

<br>

### In this template, all you need to change is: 

- (obviously) the background (the default is the Windows XP wallpaper). The background is essentially the main design of the certificate without any text.
- the ESN logo (currently ESN Greece's logo)
- and lastly, the signature itself and the text next to it.
- Everything in {{ }} will be replaced by the code based on the information in the spreadsheet.

<br>

### After changing the template to your needs, you can add the participants' data. For this, you need to keep in mind the following:
- Either the first and last names or the full name is needed. That depends on the way you collected them. The code will combine the first and last names to create the full name if the full name is empty.
- In the email field, if it is filled with an email address, the generated certificate will be sent as an attachment to that email. After the email is sent, the cell will turn green to let you know.
- When you click on either button, if a Google authentication window pops up, authenticate the code for it to run. After that, you must click the button to run the code with the proper permissions.
- Note that any email sent will be from the email that authenticated the script.
- Under the generate Certificates button, after it is clicked, a link for a folder will be created. This folder will be in the same folder as the spreadsheet and contain a folder for every country present in the ESN Country field in the Data sheet.
