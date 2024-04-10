#
#   Questo script estrae in massa su formato CSV informazioni dalle sole email di reportistica COFENSE
#       (.msg) contenute nella directory di esecuzione. 
#   Può dare errore se i messaggi risultano aperti in outlook durante la fase di raccolta.
#
#   Output di default: mail_report.csv , URL_report.csv
#


from pathlib import Path
from datetime import datetime
import win32com.client
import os
import csv
import re
import base64

#Contatori globali per definizione CSV
max_attachment=0
max_nodes=0
no_localhost=0
# Strutture di raccolta
URLs=[]
lettura=[]


def checkBase64 (item):
    #In presenza di base64 effettua la decodifica, altrimenti restituisce la stringa originale
    try:
        base64_str = re.search(r'=\?[uU][tT][fF]-8\?B\?([^?]+)\?=', item).group(1)
        decoded_bytes = base64.b64decode(base64_str)
        decoded_string = decoded_bytes.decode('utf-8')
        return decoded_string
    except:
        return item
    
def makeHeader():
    #Crea una intestazione dinamica per il CSV di riepilogo mail in base al numero di allegati e di delivery nodes
    header=['Mittente','Destinatario','Oggetto','Data']
    for i in range(max_attachment):
        header.append("Allegato" + str(i+1))
        header.append("Hash" + str(i+1))
    for i in range (max_nodes):
        header.append("Nodo" + str(max_nodes-i))
    return header

def removeDuplicatesUrls():
    #Metodo di supporto per eliminare URL ridondanti
    global URLs
    URLs=list(set(URLs))
    
def getURLs(body):
    #Metodo di supporto per la raccolta delle URL dalle email e contestuale ordinamento alfabetico
    sites = re.findall('URL: (.*)',body,re.M)
    for site in sites:
        URLs.append(site.replace('\r', ''))
    URLs.sort()
       

def writeMailCSV():
    #Scrive un CSV di riepilogo delle email Cofense
    with open(r'mail_report.csv', mode='w', newline='',encoding="utf-8-sig" ) as csvfile:
        header = makeHeader()
        writer = csv.DictWriter(csvfile,header)
        writer.writeheader()
        for mail in lettura:
            to_write={'Mittente': mail['Mittente'],'Destinatario': mail['Destinatario'],'Oggetto': mail['Oggetto'],'Data': mail['Data']}
            mail_fields=mail.keys()
            for field in header:
                if field in mail_fields:
                    to_write[field]=mail[field]
                else:
                    to_write[field]="#"
            writer.writerow(to_write)
            
def writeURLCSV():
    #Scrive un CSV di sole URL individuate nelle email
    with open(r'URL_report.csv', mode='w') as csvfile:
        writer=csv.writer(csvfile)
        for site in URLs:
            writer.writerow([site])

def setDateFormat(date):
    #Cambia il formato della data nell'header in dd-mm-YY HH:MM
        #N.B. : l'header Date si presenta in formati differenti
    date_string_parts = date.split()
    if(len(date_string_parts))>5:
        date_string_clean = ' '.join(date_string_parts[:5])
        date_object = datetime.strptime(date_string_clean, '%a, %d %b %Y %H:%M:%S')
    else:
        date_string_clean = ' '.join(date_string_parts[:4])
        date_object = datetime.strptime(date_string_clean, '%d %b %Y %H:%M:%S')
    date = date_object.strftime('%d-%m-%y %H:%M')  
    return date
        
  
def populate():
    #Popola le strutture dati di raccolta, da cui viene creato successivamente il rispettivo CSV
    global max_attachment
    global max_nodes
    global no_localhost
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    for p in Path(os.getcwd()).iterdir():
        if p.is_file() and p.suffix == '.msg':
            mail={}
            msg = outlook.OpenSharedItem(p)
            
                    ##Mittente##
            
            # La raccolta del mittente è strettamente dipendente dal formato dell'intestazione From. 
            sender_address=re.search('From:.*?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6})',msg.body)
            if(sender_address):
                sender_address=sender_address.group(1)
            else:
                #Eventualità in cui l'header From riporti anche il nome dell'autore
                sender_address=re.search('From:\s*[^<]*<([^\s<@]+@[^\s<@]+\.[^\s<@]+)>',msg.body)
                if (sender_address):
                    sender_address=sender_address.group(1)
                else:
                # Nota: Alcune intestazioni sono manipolate e il sender non è immediatamente individuato. In quel caso verifica manuale
                    sender_address="VERIFICARE MANUALMENTE"
                    
                    ##Destinatario##
                    
            recipient=re.search(' To: .*?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6})',msg.body)
            # Il destinatario non compare nei casi di Undisclosed Recipients
            if (recipient):
                recipient=recipient.group(1)
            else:
                recipient="UNDISCLOSED?"
                
                    ##Oggetto##
                    
                    
            # L'oggetto della mail può essere codificato o non presente. Si verifica
            subject=re.search('Subject: (.*)',msg.body)
            if (subject):
                subject=checkBase64(subject.group(1)).replace('\r', '')
            else:
                subject="#"
                
                    ##Data##
                    
                    
            # Il formato della data nell'header viene sostituito
            date=re.search('(?s:.*)Date: (.*)',msg.body).group(1).replace('\r', '')
            date=setDateFormat(date)
            
                    ##Links##
            
            getURLs(msg.body)
            
                    ##Allegati e relativi HASH SHA1 ##
                    
                    
            mail={'Mittente':sender_address,'Destinatario':recipient,'Oggetto':subject,'Data':date}
            attachments=re.findall('File Name:(.*)',msg.body,re.M)
            filehash=re.findall('SHA1 File Checksum:(.*)',msg.body,re.M)
            count_attachments = len(attachments)    
            if count_attachments > 0:
                if count_attachments>max_attachment:
                    max_attachment=count_attachments
                for item in range(count_attachments):
                    #il dizionario si estende dinamicamente in base alla quantità di allegati
                    mail['Allegato'+str(item+1)]=attachments[item].replace('\r', '')
                    mail['Hash'+str(item+1)]=filehash[item].replace('\r', '')          
                    
                    ##IP dei nodi di consegna##
                    
            sender_ip=re.findall('Received: .*?(\d{1,3}(?:\.\d{1,3}){3})',msg.body,re.M)
            delivery_nodes=len(sender_ip)
            if(delivery_nodes)>0:
                for i in range (delivery_nodes):
                    if (sender_ip[i])!="127.0.0.1":
                        mail['Nodo'+str(i+1)]=sender_ip[i]
                        no_localhost+=1
                if delivery_nodes-no_localhost>max_nodes:
                    max_nodes=delivery_nodes-no_localhost
                else:
                    max_nodes=delivery_nodes
            lettura.append(mail)            
            del msg
    del outlook
    

populate()
writeMailCSV()
writeURLCSV()

#################v1 29/2/24
########### FdF
######

###Changelog
# 5/4/24 - Modificata regex di cattura Data per nuova variante headers 
