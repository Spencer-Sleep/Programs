import email
import imaplib
import os
from exchangelib import DELEGATE, Account, Credentials, EWSDateTime
from exchangelib.attachments import FileAttachment, ItemAttachment
from exchangelib.items import Message
from time import sleep
import datetime
import img2pdf
from PyPDF2.pdf import PdfFileWriter, PdfFileReader
from _io import BytesIO
import warnings
from datetime import timedelta


if __name__ == '__main__':
#     mapi/emsmdb/?MailboxId=8b3c9bfe-7977-46dd-aeaa-fac511e675d2@seaportint.com
#     fetcher = FetchEmail(r"https://exchange.mobinet.ca/", "ssleep", "ss#99PASS")
    warnings.filterwarnings("ignore")

    
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
        print(account.inbox.unread_count)
        print(datetime.datetime.now())
        for item in account.inbox.filter(is_read=False, sender="wordpress@seaportint.com")[:10]:
#         for item in account.inbox.filter(subject="POD-PB 451519-Book# -Cont# Oolu 3000188"):
#             for attachment in item.attachments:
#             if isinstance(attachment, FileAttachment):
            contNumIndex = item.subject.find("Cont#")+5
            fileName = item.subject[:contNumIndex] + item.subject[contNumIndex:].replace(" ", "").upper()
            local_path = os.path.join("J:\PODs\\", fileName + ".pdf")
            if os.path.exists(local_path):
                i=1
                while os.path.exists(local_path[:local_path.find(".")]+"("+str(i)+")"+local_path[local_path.find("."):]):
                    i+=1
                local_path=local_path[:local_path.find(".")]+"("+str(i)+")"+local_path[local_path.find("."):]
#             for attachment in item.attachments:
#                 if attachment.name=="book1.pdf":
#                     item.is_read=True
#                     item.save()
            with open(local_path, 'wb') as f:
                attachmentsContent=[]
                for attachment in item.attachments:
#                     print(attachment.name)
                    attach = attachment.content
                    if not "pdf" in attachment.name:
                        attach = img2pdf.convert(attach)
                    attachmentsContent.append(attach)
#                 if "pdf" in item.attachments[0].name:
#                     f.write(attachment.content)
                input_streams = []
                try:
                    # First open all the files, then produce the output file, and
                    # finally close the input files. This is necessary because
                    # the data isn't read from the input files until the write
                    # operation. Thanks to
                    # https://stackoverflow.com/questions/6773631/problem-with-closing-python-pypdf-writing-getting-a-valueerror-i-o-operation/6773733#6773733
                    i=0
                    if not os.path.exists("J:\PODs\\Temp Files\\"):
                        os.mkdir("J:\PODs\\Temp Files\\")
                    for input_file in attachmentsContent:
#                             input_file=input_file.encode('utf8').decode('utf8')
                        f1=open("J:\PODs\\Temp Files\\"+str(i), 'w+b')
                        f1.write(input_file)
#                             print(input_file)
                        input_streams.append(f1)
                        i+=1
                    writer = PdfFileWriter()
                    for reader in map(PdfFileReader, input_streams):
                        for n in range(reader.getNumPages()):
                            writer.addPage(reader.getPage(n))
                    writer.write(f)
                finally:
                    for f in input_streams:
                        f.close()
#                 else:
#                     f.write(img2pdf.convert(attachmentsContent))
            print('Saved attachment to', local_path)
#                 elif isinstance(attachment, ItemAttachment):
#                     if isinstance(attachment.item, Message):
#                         print(attachment.item.subject, attachment.item.body)
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