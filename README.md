# Outlook Domain Mail Organizer

Outlook VSTO plugin to help you organize your mails if you are working with a lot of different customers

The main motivation for this plugin was the following issue:
In outlook, even after I created all the rules to move messages from a customer into a specific folder, if anyone else from my organization replied to the mail thread, it landed in my inbox. The repititive task of dragging and dropping these reply mails inspired this plugin.

To use this:
1. Create a "Customers" folder in the mailbox root
2. And create a subfolder for each of the Customer domains that you wish to organize
e.g. If I work with these 4 companies, I will have my folder structure as follows
  * Customers
    * ebay.com
    * expedia.com
    * offerup.com
    * contoso.com
  
On any email, even if one of the recipients (or the sender) is of the domain @ebay.com, it will land in the ebay.com folder regardless of who sent that email and to whom.
There is a catch, that if on the email you have recipients from multiple of these companies above, it would depend on the order of the recpients. The first recipient that matches any of the domains above is the folder that email, or meeting, or appointment will go to.

Hopefully this would be useful to someone as much as it was valuable to me.

Thanks for your time!
  
