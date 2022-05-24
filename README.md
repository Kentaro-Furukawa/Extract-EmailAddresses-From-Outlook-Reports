## This code is to extract email addresses from Outlook undelivered reports to an Excel file.

 **\*Please note that, this code is for Outlook VBA.**

You might have been in a situation where you sent emails to hundreds of users/customers on Outlook, but inevitably many of the mail addresses were not reachable. As a result, you received a lot of undelivered reports, and you probably want to remove those mail addresses from the list for the future. But, how would you extract the addresses from the reports? Well, this code can be useful in this situation, so please feel free to use my code to avoid this tedious and time consuming task.

### Settings

You can set exceptional address you want to be exclude in ExceptionAddress array.

`Line 15: ExceptionAddress = Array("example1@mail.com","example2@mail.com")`
