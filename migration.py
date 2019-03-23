import datetime
import pandas as pd
import pyodbc

#CONNECTING TO ACCESS
#conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\\Users\\LHA\\Documents\\import_be.accdb;')
#cursor = conn.cursor()
#cursor.execute('select * from AKTIVT')
#rows = cursor.fetchall()
print("Hello before connection to sql")

server = 'DESKTOP-R6J8JHM\SQLEXPRESS'
database = '***Change***'
username = '***Change***'
password = '***Change***'
driver = '{ODBC driver 13 for SQL Server}' # Driver you need to connect to the database
port = '1433'
conn = pyodbc.connect('DRIVER='+driver+';SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password)
cursor = conn.cursor()

#Få fat i kursus_navn
cursor.execute("select * from kurser")
kur = cursor.fetchone()

val3 = {}

while kur:
    val3.update({kur[3]:kur[0]})
    kur = cursor.fetchone()


#Få fat i alle jobcentre og deres Id'er
cursor.execute("select * from kommuner")
kom = cursor.fetchone()
val2 = {}

while kom:
    val2.update({kom[1]: kom[0]})
    kom = cursor.fetchone()

#Få fat i alle kursisterne
cursor.execute('select [Fornavn],[Efternavn],[Tlf],[Mail],[Udd niveau],[Køn],[Fødselsdato],[Indgang],[Kontaktperson],[Note],[Status],[Blanket sendt],[A-kasse],[Jobcenter],[Startdato], [Kursustitel] FROM AKTIVT')
row = cursor.fetchone()

val = []

while row:
    køn = ""
    if (row[5] == "Kvinde"):
        køn = "0"
    elif (row[5] == "Mand"):
        køn = "1"

    #KONVERTER MEDARBEJDER INITIALER TIL INT
    if (row[8] == "Kro"):
        med = "2"
    elif (row[8] == "sme"):
        med = "3"
    elif (row[8] == "cha"):
        med = 1
    elif (row[8] == "sny"):
        med = "4"

    if (row[1] is None):
        row[1] = 'ukendt'
    if (row[2] is None):
        row[2] = 'ukendt'
    if (row[3] is None):
        row[3] = 'ukendt'
    if (row[4] is None):
        row[4] = 'ukendt'
    startDato = row[14]
    startDato = startDato.strftime("%Y-%m-%d")
    kursus_navn = str(row[15]+' - '+startDato)
    #print(kursus_navn)
    dato = datetime.datetime.today()

    #print(kursus_navn)
    jc = val2.get(row[13])

    default_value = "****----------------666-----------------****"

    kursus_id = val3.get(kursus_navn, default_value)
    #print (val3.get(kursus_navn))

    val.append((row[0], row[1], row[2], row[3], row[4], køn, row[6], row[7], row[8], row[9], row[10], med, kursus_id, dato, '0', row[12], jc, kursus_navn))
    row = cursor.fetchone()

sql = "insert into kursister (fornavn, efternavn, tlf, mail, uddannelsesniveau, køn, fødselsdato, indgang, kontaktperson, note, status, medarbejder_id, kursus_id, oprettet, [Uddannelsesbevis modtaget], [a-kasse], jobcenter, [kursus_navn-migration]) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
cursor.executemany(sql,val)
conn.commit()

print(val)

#Få fat i alle kursus-oplysninger om kursister, så de kan indsættes i dbo.kursus_kursist
cursor.execute("select [kursus_id], [kursist_id], [status] from kursister")
kk = cursor.fetchone()
kk1 = []

while kk:
    kk1.append((kk[0], kk[1], kk[2]))
    kk = cursor.fetchone()


sql = "insert into kursus_kursist (kursus_id, kursist_id, status) values (?, ?, ?)"
cursor.executemany(sql,kk1)
conn.commit()

Print("Indholdet af kk1 er nu sat ind i dbo.kursus_kursist. Indholdet er som følger \n")
print(kk1)
print("Hello after sql execution")
