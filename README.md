DistributionListUpdater
=======================

Outlook-Add-In zum automatischen Erstellen von Verteilerlisten

Autor: Christoph Fischer, christoph.fischer@elkw.de

## Einf�hrung
Verteilerlisten unter Outlook leiden unter zwei grunds�tzlichen Problemen:

1. Bestehende Outlookkontakte k�nnen zwar zu einer Verteilerliste hinzugef�gt werden, nachfolgende �nderungen an den Kontaktdaten werden aber nicht automatisch �bernommen. Die M�glichkeit, eine Verteilerliste zur Bearbeitung zu �ffnen und anschlie�end zu "aktualisieren" ist bie einer gro�en Anzahl von Verteilerlisten keine zufriedenstellende L�sung.
2. Verteilerlisten k�nnen nicht zu einer Kontaktliste auf Sharepoint hinzugef�gt und damit auf mehreren Rechnern synchron gehalten werden. Verteilerlisten k�nnen nur in lokalen (bzw. auf Exchange gespeicherten) Kontaktordnern angelegt werden.

Das vorliegende Add-In bietet einen L�sungsansatz dazu. Mit einem Klick wird ein vordefinierter Kontaktordner nach Kategorien durchsucht. Zu jeder Kategorie wird eine Verteilerliste "VL.Kategoriename" mit den
entsprechenden Kontakten erstellt.

## Installation

1. Legen Sie einen neuen, leeren Kontaktordner f�r Ihre Verteilerlisten an.
2. Laden Sie das Setup-Programm zum Add-In hier unter [Releases](https://github.com/potofcoffee/DistributionListUpdater/releases) herunter.
3. Beenden Sie Outlook, falls es aktuell ausgef�hrt wird.
4. F�hren Sie das Setup-Programm aus und �bernehmen Sie die vorgeschlagenen Einstellungen.
5. Starten Sie Outlook. Falls Sie die folgende Sicherheitsabfrage (nur beim ersten Start nach der Installation) sehen, klicken Sie auf "Installieren":

![Security Message](DistributionListUpdater/docs/SecurityWarningOnInstall.png)
6. Nach dem Start fragt das Add-In einmalig nach dem zu durchsuchenden Kontaktordner, sowie nach einem Ordner f�r die Verteilerlisten. *Bitte beachten: Der Ordner f�r die 
Verteilerlisten wird bei jeder Aktualisierung durch das Add-In komplett gel�scht und neu angelegt. Legen Sie hier keine eigenen Eintr�ge an!* Die beim ersten Start getroffene 
Zuweisung kann jederzeit �ber die Schaltfl�chen "Ko" und "List" ver�ndert werden.


## Bedienung

1. Ordnen sie den Kontakten im ausgew�hlten Kontaktordner beliebige Kategorien zu.

![Categorize Contact](DistributionListUpdater/docs/CategorizeContact.png)



2. Klicken Sie im Bereich "Verteilerlisten" des Men�bands auf die Schaltfl�che "Alle aktualisieren". (Der Bereich "Verteilerlisten" erscheint unter dem Reiter "Start", wenn Sie sich im Bereich E-Mails oder Kontakte von Outlook befinden).

![Ribbon With Button](DistributionListUpdater/docs/RibbonWithButton.png)



3. Im Verteilerlistenordner finden Sie ihre neuen Verteilerlisten. Zum einfacheren Auffinden bei der Adresseingabe wird dem Titel der Kategorie jeweils "VL." vorangestellt.

![Distribution Lists](DistributionListUpdater/docs/DistributionLists.png)

## Wichtige Informationen

1. Der Verteilerlistenordner wird bei jedem Klick auf die Schaltfl�che "Alle aktualisieren" gel�scht und neu erstellt. Manuelle �nderungen an der Verteilerliste gehen dabei verloren. Bitte nehmen Sie �nderungen nur direkt an den Kontakten bzw. deren Kategoriezuweisungen vor.
2. Kontakte mit Kategorien k�nnen �ber Sharepoint zwischen mehreren Benutzern synchron gehalten werden. Dies ist grunds�tzlich auch auf dem ELKW-Sharepoint m�glich. Dazu muss die Einrichtung einer "Kontaktliste" beantragt werden. Bitte beachten Sie, dass in der Standardeinstellung nur bestimmte Kontaktdaten synchronisiert werden. Das Feld "Kategorien" geh�rt nicht dazu. Auf dem Sharepoint ist es daher n�tig, dieses (und evtl. weitere ben�tigte Felder) erst hinzuzuf�gen. Im Fall des ELKW-Sharepoints muss dies durch die Datagroup geschehen.


## Danke
Das Installationsprogramm zum Add-In wurde mit Daniel Kraus' ausgezeichnetem [VstoAdd-InInstaller](https://github.com/bovender/VstoAdd-InInstaller) erstellt.

## Lizenz
Dieses Add-In wird unter der GNU GPLv3-Lizenz angeboten. N�here Informationen dazu finden sich in der Datei [LICENSE](LICENSE). Ausf�hrliche deutschsprachige Lizenzinformationen finden sich [hier](http://www.gnu.de/documents/gpl.de.html).
