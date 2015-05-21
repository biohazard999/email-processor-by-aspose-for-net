# Introduction #

**Email Processor by Aspose.Email for .NET** lets you view your emails from standard Exchange, IMAP and POP3 servers. You can export or convert the emails to PDF, DOCX, PST, TIFF or XPS. It is developed in ASP.NET, HTML5 and jQuery and supports the  latest versions of major browsers like Firefox, Chrome, Internet Explorer.


# Features #

  * Connect to Exchange, IMAP or POP3 mail servers like hosted Exchange, Gmail etc
  * List folders and sub folders
  * List email messages from any folder
  * Search emails
  * View the email message with HTML formatting
  * View and download the attachments in email
  * Export a folder or selected messages to PST
  * Convert selected messages to variety of formats like PDF, XPS, DOCX and TIFF.

## Login ##

You can connect to any Exchange, IMAP or POP3 server.


### Exchange Server ###

For connection to Exchange server, following details are required.


  1. **Server URL:** Usually, this is in the format of https://exchange.domain.com/ews/exchange.asmx and you can replace domain.com with your company's domain name
  1. **Username:** Your Exchange server email address
  1. **Password:** Your password
  1. **Domain:** This is optional, you can leave it empty if you provided email address in username

### IMAP Server ###

For connection to an IMAP server, it requires the following details


  1. **Server URL:** Your IMAP server URL.
  1. **Username:** Your email address
  1. **Password:** Your password
  1. **SSL Enabled:** If your server requires SSL, then enable it and specify the port number. Typically, this is 993

### POP3 Server ###

POP3 also requires the similar information as the IMAP server like URL, username, password and SSL port (optional).


**Note:** _POP3 servers are very rarely used these days, as it is based on older protocol. If you download a message, it gets deleted on the server by default. IMAP or Exchange servers are commonly used, as they provide better email sync  and email management features._

![https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131708/1/01-login.png](https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131708/1/01-login.png)

## Instructions for Gmail Users ##

You can also access your Gmail account using the IMAP or POP3 option. But Gmail does not allow IMAP or POP3 access by default, you need to enable it in **Settings**. Moreover, it does not allow external apps or websites to access the emails, to solve this, you also need to allow access using [this link](https://www.google.com/settings/security/lesssecureapps).


### Step 1: Enable IMAP/POP3 ###

Open the Gmail settings.

![https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131709/1/01-settings.png](https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131709/1/01-settings.png)

On the Settings page, click on Forwarding and POP/IMAP. Enable IMAP, if you want to access Gmail using IMAP method (recommended). Enable POP3, if you want to access with this method.

![https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131710/1/02-enable-imap.png](https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131710/1/02-enable-imap.png)

### Step 2: Allow access to less secure apps ###

Open https://www.google.com/settings/security/lesssecureapps and choose **Enable**.

![https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131711/1/04-enable-sign-in-access.png](https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131711/1/04-enable-sign-in-access.png)

## List Folders ##

Once you are connected to a mail server, it will show you the list of folders and sub folders (if any).

![https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131712/1/02-list-folders.png](https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131712/1/02-list-folders.png)

## Get Emails ##

To get the latest emails, click on any folder e.g. Inbox, Sent Items etc. it will show you the list of emails.

![https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131713/1/03-list-emails.png](https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131713/1/03-list-emails.png)

## Export to PST ##

You can export the emails to a PST file. The PST will be downloaded to your local drive and you can open it in Microsoft Outlook. You can either add the whole folder or select specific emails, to be included in the PST file.

![https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131714/1/04-pst-export.png](https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131714/1/04-pst-export.png)

## Convert to PDF, DOCX, XPS and TIFF ##

You can also convert the selected messages to PDF, DOCX, XPS and TIFF. After conversion, the file will be downloaded to your local drive.


If you select multiple emails, they will be concatenated in one output document. And in case of TIFF conversion, it will have multiple pages, if you select large or multiple emails.

![https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131715/1/05-convert-pdf-docx-tiff-xps.png](https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131715/1/05-convert-pdf-docx-tiff-xps.png)

## View Email ##

Click on any email from the list to view it. The message will retain the HTML formatting, CSS styles and embedded images.

![https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131716/1/06-view-email.png](https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131716/1/06-view-email.png)

## Download Email Attachments ##

Attachments (if any) will be listed on top of the message. Click on an attachment to download it to your local drive.

![https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131717/1/07-download-attachments.png](https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131717/1/07-download-attachments.png)

## Search Emails ##

You can search the emails by entering the text in the top textbox. It matches the search text in subject, to, from and body of the email.

![https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131718/1/08-search-emails.png](https://i1.code.msdn.s-msft.com/email-processor-by-aspose-ebed3ee0/image/file/131718/1/08-search-emails.png)

# Support and Feedback #

Please feel free to ask any questions or share feedback here.


If you have any query related to the [Aspose components](http://www.aspose.com/.net/total-component.aspx), please use our [forums](http://www.aspose.com/community/forums/default.aspx).
