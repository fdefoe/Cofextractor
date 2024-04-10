**Premessa sul prodotto PhishMe**

Cofense (PhishMe!) è una affermata soluzione enterprise di protezione dal fenomeno del Phishing, che integrandosi
con il client di posta elettronica consente una immediata segnalazione del messaggio sospetto agli addetti della sicurezza IT.

Al seguente link, scelto casualmente dal web, è riportata una buona panoramica di come l'add-on Cofense si collochi
ad esempio sul client Outlook attraverso il famoso Fish-button:
https://www.oa.pa.gov/Documents/Cofense-Report-Phishing-User-Guide.pdf


**Funzionamento PhishMe**

La singola email segnalata dall'utente viene compressa in un file zip dal prodotto PhishMe ed inviata in allegato ad un
indirizzo di posta prescelto: per limitare le interazioni con l'originale l'archivio sarà protetto da password
configurata dagli amministratori, ed al fine di facilitare le analisi, il body della mail conterrà copia degli headers di quella
segnalata.

![image](https://github.com/fdefoe/Cofextractor/assets/166450568/f21ce68d-39cb-4042-9402-511ae2e8f995)

Un ulteriore valore aggiunto di PhishMe sta nell'inclusione di statistiche aggregate per report,
oltre ad un efficace censimento di link ed allegati:

![image](https://github.com/fdefoe/Cofextractor/assets/166450568/61031b14-4c8f-4846-931e-14f22c531c41)
![image](https://github.com/fdefoe/Cofextractor/assets/166450568/a6f23cf5-beb9-46be-b175-efff41da50de)
![image](https://github.com/fdefoe/Cofextractor/assets/166450568/f31458a9-8ca8-49c4-892d-8b1bbde48b34)

Lo svantaggio si trova nel dover analizzare singolarmente ogni email segnalata tramite il prodotto, senza una console
interattiva che consenta correlazioni e fornisca un quadro globale del trend del periodo.


**Cofextractor**

Cofextractor è un tool scritto in Python che effettua il parsing di massa degli headers riportati nei messaggi di segnalazione,
riversando su CSV tutti gli elementi utili ad identificare il singolo invio ed il relativo contenuto.
Alla versione attuale si innesca semplicemente avviandolo da una folder nella quale sono stati copiati i messaggi da analizzare.
Come da immagine a seguire, gli output sono:
* _URL_report_ contenente un elenco di tutti i link contenuti nelle mail esaminate, senza duplicati
* _mail_report_ contenente mittente, destinatario,data di invio, nomi di eventuali allegati e relativi hash, indirizzi IP sulla
delivery chain (dove minore è la posizione più ci si avvicina al client di invio)


![image](https://github.com/fdefoe/Cofextractor/assets/166450568/d80513da-4429-4977-867c-cad60e1646b6)


Un esempio del mail report:
![image](https://github.com/fdefoe/Cofextractor/assets/166450568/5759cdd4-ea34-4e67-85c8-dcf1429e72fb)


Dipendenze:
* Pywin32 (https://pypi.org/project/pywin32/)

**Avvertenze**

Lo script è pensato per lavorare unicamente con i report del prodotto PhishMe, in quanto strettamente dipendente dal formato.
