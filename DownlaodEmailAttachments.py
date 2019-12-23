import os
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
inbox = inbox.Folders['apitool']
messages = inbox.Items
message = messages.GetFirst()

while message is not None:
    try:
        print (message)
        attachments = message.Attachments
        attachment = attachments.Item(1)
        attachment.SaveASFile(os.getcwd() + '\\' + str(attachment)) #Saves to the attachment to current folder
        print (attachment)
        message = messages.GetNext()

    except:
        message = messages.GetNext()
