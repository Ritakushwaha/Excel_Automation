import win32com.client
import os
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

for account in mapi.Accounts:
  print(account.DeliveryStore.DisplayName) #outlook account

inbox = mapi.GetDefaultFolder(6) #Inbox folder
inbox = inbox.Folders["your folder"] #Folder inside Inbox Folder
messages = inbox.Items 
received_dt = datetime.now() - timedelta(days=1)
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
email_sender = 'sender@outlook.com'
email_subject = 'Subject of mail'
messages = messages.Restrict("[ReceivedTime] >= '"+received_dt+"'")

#save to current directory
outputDir = os.getcwd()

try:
  for message in list(messages):
    if email_subject == message.subject and message.SenderEmailAddress == email_sender and message.ReceivedTime.strftime('%Y-%m-%d') == _date:
      try:
        s = message.sender
        for attachment in message.Attachments:
          attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))
          print(f"attachment {attachment.FileName} from {s} saved")
      except Exception as e:
          print("Error when saving the attachment:" + str(e))
except Exception as e:
    print("Error when processing emails messages:" + str(e))
finally:
    print("Done.")
