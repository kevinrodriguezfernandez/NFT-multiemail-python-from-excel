# importing pandas as pd
import smtplib
import ssl
from email.message import EmailMessage
import pandas as pd

# dictionary of lists
excel_file= 'prueba.xlsx'



class Person:
    def __init__(self, cliente, codigoDesc, foto, mail):
        self.cliente = cliente
        self.codigoDesc = codigoDesc
        self.foto = foto
        self.mail = mail

    def __repr__(self):
        return f"{self.cliente} {self.codigoDesc} {self.foto} {self.mail}"



# creating a dataframe from a dictionary
df=pd.read_excel(excel_file, sheet_name=0)
persons = []
# iterating over rows using iterrows() function
for i in df.itertuples():
    person = Person(i[1],i[2],i[3],i[4])
    persons.append(person)
   
print(persons)

print(persons[0].mail)

# create a template for send email with python
port = 465  # For SSL
email_sender = 'yourmail'
password = 'yourpassword'
contador=0

for person in persons:
    email_receiver = person.mail
    contador=contador+1
    print(contador)
    subject = f"{person.cliente} aqui tienes tu NFT \n"

    body = f"""\n Hola {person.cliente}, has ganado el NFT {person.foto} con el codigo {person.codigoDesc},
     envianos tu cartera web 3.0 y te lo enviaremos

    """
    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = email_receiver
    em['Subject'] = subject
    em.set_content(body)
    fullBody =em.as_string().replace("""Content-Type: text/plain; charset="utf-8"
Content-Transfer-Encoding: 7bit
MIME-Version: 1.0""","")
    # Create a secure SSL context
    context = ssl.create_default_context()
    print(fullBody)
    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as smtp:
        smtp.login(email_sender, password)
        smtp.sendmail(email_sender, email_receiver, fullBody)

