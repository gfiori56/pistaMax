Gestione gare su pista per 2 automobiline.
Si basa su 2 microinterruttori che segnalano il passaggio della singola macchinina; essi sono acquisiti da un Arduino Uno R4 WiFi.

Su Arduino sono rilevati i fronti dei segnali, poi sono filtrati per eliminare i rimbalzi.
Arduino invia l'informazione dei passaggi al PC tramite un cavo USB che nello stesso tempo lo alimenta. 
Con questo modello di Arduino sfrutto la sua matrice di LED per diagnosticare lo stato dei microinterruttori.
Il programma di interfaccia sul PC è scritto in VB6 e funziona con Windows 10.

