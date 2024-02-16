![image](https://github.com/DaeSikWoo/Apps-Script-Email/assets/35883117/7b8a1e00-c7da-4cc1-88e0-dc50f69bfa39)# Apps-Script-Email
This code utilizes Google Apps Script in Google Spreadsheet to automatically send specific details from a table to relevant accounts for verification. This approach was emerged due to the need to protect personal information. When dealing with over 1,000 different accounts, it becomes time-consuming and labor-intensive for data admins to send individual emails. Consequently, the spreadsheet was made publicly accessible for account holders to review their information. With this code, emails can be sent to multiple accounts simultaneously with just a click, saving the admin the time and effort of sending them one by one.

### Note
There is a quota for the number of emails that can be sent depending on the Google account type. Please check the link (https://developers.google.com/apps-script/guides/services/quotas?hl=ko) and adjust your usage according to your account type.

## Running the code
mailApp.gs consists of two functions : The first function sends all the information simultaneously, while the second function is triggered to send information whenever a new row is added.

### sendEmailsAll()

![image](https://github.com/DaeSikWoo/Apps-Script-Email/assets/35883117/e7df57bb-df8c-4481-a921-9c47cf9f00fe)



 
