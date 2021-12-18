<h2>This code is to extract email addresses from Outlook undelivered reports to an Excel file.</h2>

*Please note that, this code is for Outlook VBA.

<lead>You might have been in a situation that you sent mails to hundreds of users/customers on Outlook, but numbers of the mail addresses were not reachable, and you received a lot of undelivered reports, then you probably might want to remove those mail addresses from the list for the future. But, how would you extract the addresses from the reports? So, this code can be useful in the situation, please feel free to use the code and avoid the tedious and time consuming task.</lead>

<section>
<p>You can set exceptional address you want to be exclude in ExceptionAddress array.</p>
<code>
Line 15: ExceptionAddress = Array("example1@mail.com","example2@mail.com")
</code>
</section>
