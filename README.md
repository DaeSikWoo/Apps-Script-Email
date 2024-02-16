# Apps-Script-Email
This code utilizes Google Apps Script in Google Spreadsheet to automatically send specific details from a table to relevant accounts for verification. This approach was emerged due to the need to protect personal information. When dealing with over 1,000 different accounts, it becomes time-consuming and labor-intensive for data admins to send individual emails. Consequently, the spreadsheet was made publicly accessible for account holders to review their information. With this code, emails can be sent to multiple accounts simultaneously with just a click, saving the admin the time and effort of sending them one by one.

### Note
There is a quota for the number of emails that can be sent depending on the Google account type. Please check the link (https://developers.google.com/apps-script/guides/services/quotas?hl=ko) and adjust your usage according to your account type.

## Running the code
mailApp.gs consists of two functions : The first function sends all the information simultaneously, while the second function is triggered to send information whenever a new row is added.

### sendEmailsAll()
![image](https://github.com/DaeSikWoo/Apps-Script-Email/assets/35883117/47c1f4d7-5429-441a-a4d2-0a9f44f55aa5)
![image](https://github.com/DaeSikWoo/Apps-Script-Email/assets/35883117/9641a3f1-31c8-4d1c-b907-1a5edf71ff80)
![image](https://github.com/DaeSikWoo/Apps-Script-Email/assets/35883117/744e441f-8253-43df-a362-5716ea5a4fc2)

### sendEmailsNew()
![image](https://github.com/DaeSikWoo/Apps-Script-Email/assets/35883117/bdcb8221-2e21-4fa0-a662-97802c328207)
![image](https://github.com/DaeSikWoo/Apps-Script-Email/assets/35883117/3b285ebd-70e9-4af6-acfc-0198c97bea5a)






 
