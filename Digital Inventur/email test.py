
from exchangelib import DELEGATE, Account, Credentials, Configuration, Message, FileAttachment


# Kontodatei und Server
creds = Credentials(username="inventur@argen.de", password="Argonaut_51%")
configur = Configuration(server='mail.argen.de', credentials=creds)

# Account create
account = Account(primary_smtp_address="inventur@argen.de", config=configur,
                  autodiscover=False, access_type=DELEGATE)

m = Message(account=account,
            subject='Inventur',
            body='Inventur',
            to_recipients=["d.bondarenko@argen.de"])
# Files
#with open('file.xlsx', 'rb') as f:
#    file_content = f.read()
#file = FileAttachment(name='file.xlsx', content=file_content)
#m.attach(file
m.send()