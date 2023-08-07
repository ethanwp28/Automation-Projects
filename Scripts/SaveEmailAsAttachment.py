# I use this script to automatically save down emails we receive that have a specific subject (and associated attachments) to a specified file location.

import os
import win32com.client


# sets up the file path to save the file. uncomment and replace the text in quotes with the file path
# path_to_save = r'destination file path here'

#setup outlook
application = win32com.client.Dispatch("Outlook.Application")
namespace = application.GetNamespace("MAPI")

# Default folder should just be inbox, in order to access subfolders. This part should be amended if you wish to use the "Sent" inbox, which I also use.
inbox = namespace.GetDefaultFolder(6)

# Uncomment and rename to whatever folder/subfolder you get these emails in
#subfolder = inbox.Folders['name here']

for message in subfolder.Items:
    if 'Subject' in message.Subject:
        if message.Attachments.Count > 0: #for my use case, I always wanted emails with an attachment. This piece may or may not be relevant
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path_to_save, attachment.FileName)) # this script saves the email as an attachment.

print('done')