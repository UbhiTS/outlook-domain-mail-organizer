# Outlook Domain Mail Organizer

Outlook VSTO plugin to help you organize your mails if you are working with a lot of different customers

The main motivation for this plugin was the following issue:
In outlook, even after I created all the rules to move messages from a customer into a specific folder, if anyone else from my organization replied to the mail thread, it landed in my inbox. The repititive task of dragging and dropping these reply mails inspired this plugin.

To use this:
1. Download the code, compile it in Visual Studio (you need to have VSTO toolkit installed). and Start the project. This need to be done only once.
2. Once Outlook starts, Create a "Customers" folder in the mailbox root
3. And create a subfolder for each of the Customer domains that you wish to organize
   e.g. If I work with these 4 companies, I will have my folder structure as follows
   
* Customers
  * ebay.com
  * expedia.com
  * offerup.com
  * contoso.com

4. Restart Outlook

When any new email arrives, all unread emails in the inbox will be evaluated and moved to matching folders. A quick note and a warning. If you have more than 20000 emails, it could take up to 15 to 30 minutes depending on your PC configuration. If you want to process only new emails, ensure that you mark all existing mails in the inbox as read.

On any email, even if one of the recipients (or the sender) is of the domain @ebay.com, it will land in the ebay.com folder regardless of who sent that email and to whom.
There is a catch, that if on the email you have recipients from multiple of these companies above, it would depend on the order of the recpients. The first recipient that matches any of the domains above is the folder that email, or meeting, or appointment will go to. But normally that would not be a use case :)

Hopefully this would be useful to someone as much as it was valuable to me.

Thanks for your time!
  
