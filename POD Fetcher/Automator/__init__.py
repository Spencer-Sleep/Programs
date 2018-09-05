import email
import imaplib
import os
from exchangelib import DELEGATE, Account, Credentials
from exchangelib.attachments import FileAttachment, ItemAttachment
from exchangelib.items import Message
from time import sleep
import datetime
import img2pdf
    
if __name__ == '__main__':
#     mapi/emsmdb/?MailboxId=8b3c9bfe-7977-46dd-aeaa-fac511e675d2@seaportint.com
#     fetcher = FetchEmail(r"https://exchange.mobinet.ca/", "ssleep", "ss#99PASS")
    
    
    credentials = Credentials(
    username='mprewer@seaportint.com',  # Or myusername@example.com for O365
    password='mp#99PASS'
    )
    account = Account(
        primary_smtp_address='POD@seaportint.com', 
        credentials=credentials, 
        autodiscover=True, 
        access_type=DELEGATE
    )
    
    
    
#     credentials = Credentials(
#     username='ssleep@seaportint.com',  # Or myusername@example.com for O365
#     password='ss#99PASS'
#     )
#     account = Account(
#         primary_smtp_address='spencer.sleep@seaportint.com', 
#         credentials=credentials, 
#         autodiscover=True, 
#         access_type=DELEGATE
#     )


#     print("here")
    # Print first 100 inbox messages in reverse order
    while(True):
#         print(account.inbox.unread_count)
#         print(datetime.datetime.now())
        for item in account.inbox.filter(is_read=False, sender="wordpress@seaportint.com"):
            for attachment in item.attachments:
                if isinstance(attachment, FileAttachment):
                    contNumIndex = item.subject.find("Cont#")+5
                    fileName = item.subject[:contNumIndex] + item.subject[contNumIndex:].replace(" ", "").upper()
                    local_path = os.path.join("J:\PODs\\", fileName + ".pdf")
                    with open(local_path, 'wb') as f:
                        if "pdf" in attachment.name:
                            f.write(attachment.content)
                        else:    
                            f.write(img2pdf.convert(attachment.content))
                    print('Saved attachment to', local_path)
                elif isinstance(attachment, ItemAttachment):
                    if isinstance(attachment.item, Message):
                        print(attachment.item.subject, attachment.item.body)
            item.is_read=True
            item.save()
#         print(datetime.datetime.now())
        sleep(30)
        account.inbox.refresh()
#     for item in account.inbox.all().order_by('-datetime_received')[:100]:
#         print("1")
#         print(item.subject, item.body, item.attachments)
#     print("here2")

#     pyinstaller "C:\Users\ssleep\workspace\POD Fetcher\Automator\__init__.py" --distpath "J:\Spencer\POD Fetcher" -y