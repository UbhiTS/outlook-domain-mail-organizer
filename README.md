# Outlook Domain Mail Organizer
A Microsoft Outlook plugin to help organize your mails when you are working with many customers
![image](https://github.com/UbhiTS/outlook-domain-mail-organizer/assets/3799525/f9561809-ae2c-43ec-825b-e9a3c6f82bee)


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
   -  Create a "Domains" folder in the **mailbox root**
   -  Create a subfolder for each of your domains you get email from that you wish to auto organize
   -  (e.g. if I work with these 4 companies, I will have my folder structure as follows)
```
      Inbox
      Outbox
      Sent Items
      Drafts
      Deleted Items
      Domains
        |- ebay.com
        |- expedia.com
        |- offerup.com
        |- contoso.com
```
4. And click the "Process Inbox" Button
5. Once you have all inbox mails organized, you can archive all inbox items by clicking "Archive Mails" button
6. In case you overlook an email and add a customer folder after moving all items to the archive, you can always process the Archive folder to re-organize mails just like from the Inbox folder

I hope this helps keep your inboxes clean!

## Experience
   - If the **sender, or one of the recipients (to, cc)** is from the domain @xyz.com, it will land in the xyz.com folder.
   - if the **subject** contains the customers domain, or any **keyword defined in the description of the folder**, mail will be moved to that folder
   - if the **body** contains the customers domain, or any **keyword defined in the description of the folder**, mail will be moved to that folder
   - else leave it in inbox for you to review and archive later

# Thank you!
Hopefully this is useful to you and helps you keep your outlook organized. Thanks for your time.
Cheers!
