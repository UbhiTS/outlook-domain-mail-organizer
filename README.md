# Outlook Domain Mail Organizer
A Microsoft Outlook plugin to help organize your mails when you are working with many customers
![image](https://user-images.githubusercontent.com/3799525/229867671-80e553de-3d66-4778-860b-6def867fabb0.png)

## Motivation
1. Having a very **noisy inbox** and **missing important customer mails** due to promotions, trainings, dl groups etc.
2. **Creating and managing rules** for each and every customer to land mails in specific folders
3. Reduce the repetitive task of dragging and dropping these specific mails that land in your inbox regardless of rules
   - **Someone else from your organization replies to the mail thread**
   - **You are cc'd on a customer thread**
4. Hard to **keep a track of most active customers/requests**.
   - The latest customer with a request **automatically moves to the top of your list**

## How to install
1. Download the code, compile it in Visual Studio (with VSTO development plugin)
2. Start the project **(steps 1 and 2 need to be done one time only)**
4. Once Outlook starts
   -  Create a "Customers" folder in the **mailbox root**
   -  Create a subfolder for each of your Customer domains you wish to auto organize
   -  (e.g. if I work with these 4 companies, I will have my folder structure as follows)
```
      Inbox
      Outbox
      Sent Items
      Drafts
      Deleted Items
      Customers
        |- ebay.com
        |- expedia.com
        |- offerup.com
        |- contoso.com
```
4. Restart Outlook

## Experience
When a new mail arrives, **all unread emails** in the inbox will be evaluated and moved to matching folders. 
Even **if one of the recipients (to, cc) or the sender** is of the domain @ebay.com, it will land in the ebay.com folder. This is useful when you are in the cc list.

## Caution
If you have more than 20000 emails, it could take up to 15 to 30 minutes depending on your PC configuration. If you want to process only new emails, **ensure that you mark all existing mails in the inbox as read**.

If you have recipients from multiple companies, it would depend on the order of the recpients. **The first recipient that matches** any of the domains is the folder that email, or meeting, or appointment will go to. But normally that's not really a use case, is it :) ?

This plugin is not ready for primetime and is not production ready so I do not claim any performance guarantees. In fact, you might have to reenable the plugin if outlook disables it citing performance. If you plugin disappears from your toolbar, you can reenable it via the following process -

Go to File >

![image](https://user-images.githubusercontent.com/3799525/229868171-575e6d09-9411-4577-8939-afb08db0db2f.png)

Go to File > Options

![image](https://user-images.githubusercontent.com/3799525/229867079-941f21b6-271a-463f-90bb-f322260778fa.png)

# Thank you!
Hopefully this is useful to you and help you keep your outlook organized. 
Thanks for your time!
