import os
import win32com.client

# sets up the file path to save the file
path_to_save = r'file path here'  # replace with your actual path

# setup outlook
application = win32com.client.Dispatch("Outlook.Application")
namespace = application.GetNamespace("MAPI")

# Get the Sent folder
sent_folder = namespace.GetDefaultFolder(5)  # 5 corresponds to the Sent folder

# Iterate through the emails to find the most recent with desired subject
# Assuming the Items are sorted by received date
for item in sent_folder.Items:
    if "subject here" in item.Subject: # replace with email subject
        recent_email = item

        # Save the email itself as an attachment
        email_filename = recent_email.Subject + ".msg"
        recent_email.SaveAs(os.path.join(path_to_save, email_filename))

        # If you also want to save the other attachments of the email
        if recent_email.Attachments.Count > 0:
            for attachment in recent_email.Attachments:
                attachment.SaveAsFile(os.path.join(path_to_save, attachment.FileName))
        
        print('done')
else:
    print("No recent email found with the subject containing 'Delta Review'")
