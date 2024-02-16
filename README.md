# Apps-Script-Email
This code utilizes Google Apps Script in Google Spreadsheet to automatically send specific details from a table to relevant accounts for verification. This approach was emerged due to the need to protect personal information. When dealing with over 1,000 different accounts, it becomes time-consuming and labor-intensive for data admins to send individual emails. Consequently, the spreadsheet was made publicly accessible for account holders to review their information. With this code, emails can be sent to multiple accounts simultaneously with just a click, saving the admin the time and effort of sending them one by one.

### Note
There is a quota for the number of emails that can be sent depending on the Google account type. Please check the link (https://developers.google.com/apps-script/guides/services/quotas?hl=ko) and adjust your usage according to your account type.

## Running the code
mailApp.gs consists of two functions : The first function sends all the information simultaneously, while the second function is triggered to send information whenever a new row is added.

### sendEmailsAll()
- Click the button labeled '메일발송'. A popup will appear - select '전체발송' to send data to multiple users.
![image](https://github.com/DaeSikWoo/Apps-Script-Email/assets/35883117/47c1f4d7-5429-441a-a4d2-0a9f44f55aa5)
- The code gathers email addresses from column D and sends emails to those addresses. Once the emails are sent, the 'update datetime' in the table will be updated.
![image](https://github.com/DaeSikWoo/Apps-Script-Email/assets/35883117/9641a3f1-31c8-4d1c-b907-1a5edf71ff80)
- Only pertinent information will be sent to the corresponding user. The information will appear in the email as it does in the table.
![image](https://github.com/DaeSikWoo/Apps-Script-Email/assets/35883117/744e441f-8253-43df-a362-5716ea5a4fc2)

### sendEmailsNew()
- Click the button labeled '메일발송'. A popup will appear - select '신규발송' to send the new data.
![image](https://github.com/DaeSikWoo/Apps-Script-Email/assets/35883117/bdcb8221-2e21-4fa0-a662-97802c328207)
- Only the new data will be sent and the 'update datetime' for the email will be applied only to this new row.
![image](https://github.com/DaeSikWoo/Apps-Script-Email/assets/35883117/3b285ebd-70e9-4af6-acfc-0198c97bea5a)






 
