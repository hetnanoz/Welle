DALLAS CASH TRANSACTIONS - VBA 1.0.0
====================================

1. ZAWARTOSC PAKIETU
--------------------

modInfo.bas
modTypes.bas
modConst.bas
modMain.bas
modConfig.bas
modFileSystem.bas
modOutlook.bas
modInput.bas
modTransactionData.bas
modFondsliste.bas
modStaticData.bas
modReports.bas
modHistorical.bas

Makro startowe:

    RunDallasCashTransactions

Pliki nie zawieraja wierszy Attribute VB_Name. Moduly nalezy utworzyc/importowac pod nazwami zgodnymi z nazwami plikow.


2. WYMAGANA ZALEZNOSC BAT-CORE
-----------------------------

Projekt uzywa firmowego ErrorManager z bat-core.xlam.

W edytorze VBA:

    Tools -> References -> zaznacz wlasciwa biblioteke BAT-core

Nie dodawaj lokalnej klasy ErrorManager i nie tworz arkusza Error_Log.

Po imporcie wszystkich modulow uruchom:

    Debug -> Compile VBAProject


3. ARKUSZE W SKOROSZYCIE NARZEDZIA
----------------------------------

Mapping
    Zawiera wszystkie komorki nazwane wymienione w punkcie 4.

Static_Data
    Wiersz 1: naglowki.
    Kolumna B: IDENTIFIER / Account.
    Kolumna C: Fonds / Fund Number.
    Kolumna D: Master.
    Kolumna E: Team.

Podczas kazdego uruchomienia kolumna E jest aktualizowana na podstawie najnowszej Fondslisty. Dla funduszu, ktory nie wystepuje jako Open, Team jest czyszczony. W rezultacie transakcja bez aktywnego Teamu otrzymuje UNKNOWN TEAM.


4. WYMAGANE KOMORKI NAZWANE W ARKUSZU MAPPING
---------------------------------------------

FLAG_SAVE_LOCAL
    TRUE/FALSE. Zapis raportow do OUTPUT_LOCAL_BASE.

FLAG_SAVE_SHAREPOINT
    TRUE/FALSE. Zapis raportow do OUTPUT_SHAREPOINT_BASE.

FLAG_UPDATE_HISTORICAL
    TRUE/FALSE. Aktualizacja plikow Master w kazdej aktywnej lokalizacji.

OUTLOOK_MAILBOX
    Nazwa wyswietlana lub adres wspoldzielonej skrzynki.

OUTLOOK_SOURCE_FOLDER
    Sciezka wzgledem korzenia skrzynki, np.:
    Reports & Checks\Dallas - Cash transactions

OUTLOOK_ARCHIVE_FOLDER
    Sciezka wzgledem korzenia skrzynki, np.:
    Reports & Checks\Dallas - Cash transactions\Archive

FONDSLISTE_FOLDER
    Lokalny/UNC/WebDAV folder lub adres HTTPS SharePoint.

FONDSLISTE_PATTERN
    Prefiks albo wzorzec nazwy. Przyklady:
    Fondsliste_
    Fondsliste_*
    Fondsliste_*.xlsx

FONDSLISTE_EXTENSION
    Rozszerzenie, np. .xlsx albo .xls.

OUTPUT_LOCAL_BASE
    Bazowy folder lokalny lub sieciowy. Wymagany, gdy FLAG_SAVE_LOCAL = TRUE.

OUTPUT_SHAREPOINT_BASE
    Bazowy folder SharePoint. Wymagany, gdy FLAG_SAVE_SHAREPOINT = TRUE.

SUPPORTED_TEAMS
    Lista rozdzielona srednikami, np.:
    3;6;7;11

Kazda nazwa musi wskazywac dokladnie jedna komorke w arkuszu Mapping.


5. SHAREPOINT
-------------

Najbardziej niezawodne warianty konfiguracji to:

- lokalna sciezka biblioteki zsynchronizowanej przez OneDrive,
- sciezka UNC/WebDAV,
- adres HTTPS, ktory kod konwertuje do postaci WebDAV:
  \\tenant.sharepoint.com@SSL\DavWWWRoot\sites\...

Dla adresu HTTPS komputer musi miec aktywna obsluge WebDAV/WebClient i uwierzytelnienie do SharePoint. Gdy adres zawiera zakodowane znaki spoza ASCII, zalecana jest sciezka zsynchronizowana albo gotowa sciezka WebDAV.


6. OUTLOOK
----------

Kod uzywa late binding, wiec osobna referencja do biblioteki Outlook nie jest wymagana.

Przetwarzane sa wiadomosci, ktorych Subject zawiera:

    Cash transactions Fonds Welle

Zapisywane sa tylko zalaczniki .xlsx i .xls.

Mail jest przenoszony do Archive dopiero po poprawnym zapisaniu wszystkich obslugiwanych zalacznikow z tej wiadomosci. Jezeli pozniejszy etap procesu zakonczy sie bledem, wiadomosc bedzie juz w Archive i stamtad nalezy ja przywrocic lub ponownie uruchomic proces na jej kopii.

Makra w automatycznie otwieranych plikach sa wymuszonym sposobem wylaczone. Po zakonczeniu procesu poprzedni poziom Application.AutomationSecurity jest przywracany.


7. PLIKI WEJSCIOWE
------------------

Kod czyta pierwszy arkusz kazdego zalacznika.

Wiersz 1 musi zawierac dokladnie, bez wzgledu na wielkosc liter:

Account | Account Name | v/d | t/d | Journal | BLZ | Ref. 1 | Ref. 2 |
Ref. 3 | Ref. 4 | Ref. 5 | Ref. 6 | EUR amount | Amount | CCY | Reversal

Wszystkie rekordy sa scalane w pamieci. W raporcie oryginalne A:P trafiaja do F:U, a A:E zawiera:

A MF
B Master
C VD
D TD
E Team

VD i TD sa zapisywane jako DD.MM.YYYY.


8. FONDSLISTA
-------------

Kod wybiera leksykograficznie najwyzsza pasujaca nazwe pliku. Oznacza to, ze data/czas w nazwie Fondslisty powinny miec format sortowalny, np. YYYYMMDD.

Czytany jest pierwszy arkusz, dane od wiersza 10:

B Status
D Fund Number
O Team

Uwzgledniane sa tylko rekordy Status = Open, bez wzgledu na wielkosc liter.

Przed otwarciem pliku kod sprawdza Application.Workbooks:

- najpierw porownuje znormalizowana pelna sciezke,
- nazwa pliku jest dopuszczalnym fallbackiem,
- skoroszyt juz otwarty przez uzytkownika jest ponownie wykorzystywany i pozostaje otwarty,
- skoroszyt otwarty przez makro jest ReadOnly i zostaje zamkniety bez zapisu.

Slownik Fund Number -> Team jest budowany tylko raz na uruchomienie.


9. RAPORTY DZIENNE
------------------

Folder docelowy:

    \YYYY\MM\YYYY-MM-DD\

Plik laczny:

    Dallas_Transactions_All_Teams_Combined_YYYYMMDD_hhmmss.xlsx

Plik Teamu z danymi:

    Dallas_Transactions_Team_<TEAM>_YYYYMMDD_hhmmss.xlsx

Plik Teamu bez danych:

    Dallas_Transactions_Team_<TEAM>_YYYYMMDD_hhmmss_EMPTY.xlsx

Kazdy raport zawiera widoczny arkusz Transactions oraz ukryty arkusz Static_Data. Static_Data jest kopiowany jako wartosci, co zapobiega tworzeniu zewnetrznych polaczen do skoroszytu narzedzia.

Data folderu i nazwy pliku jest pobierana z czasu uruchomienia komputera.


10. PLIKI HISTORYCZNE
---------------------

Pliki sa przechowywane bezposrednio w aktywnym folderze bazowym:

    Master_Historical_Combined.xlsx
    Master_Historical_Team_<TEAM>.xlsx

Deduplikacja jest wykonywana w pamieci na kluczu:

    F|H|I|L|M|N|O|P|R|S|T

czyli dokladnie na kolumnach wskazanych w wymaganiach.

Istniejacy Master jest ladowany do tablicy, klucze trafiaja do Scripting.Dictionary, a do arkusza dopisywany jest jednym zapisem tylko blok nowych rekordow.

Jezeli Master jest otwarty i ma niezapisane zmiany uzytkownika, proces przerywa aktualizacje tego pliku zamiast automatycznie zapisywac cudze zmiany. Otwarty i zapisany Master moze zostac zaktualizowany, zapisany i pozostaje otwarty.


11. SZYBKI TEST ODBIORCZY
-------------------------

1. Debug -> Compile VBAProject.
2. Ustaw wszystkie trzy flagi na FALSE i potwierdz kontrolowany blad konfiguracji.
3. Wlacz tylko FLAG_SAVE_LOCAL i wykonaj test z jednym mailem i jednym zalacznikiem.
4. Sprawdz folder YYYY\MM\YYYY-MM-DD.
5. Sprawdz Combined, Team z rekordami i Team _EMPTY.
6. Sprawdz, ze Static_Data jest ukryty.
7. Otworz najnowsza Fondsliste przed uruchomieniem i potwierdz, ze po procesie pozostaje otwarta.
8. Zamknij Fondsliste, uruchom ponownie i potwierdz, ze plik otwarty przez makro zostaje zamkniety.
9. Wlacz historie, uruchom dwa razy na tych samych danych i potwierdz brak duplikatow.
10. Wykonaj osobny test SharePoint/WebDAV.
11. Potwierdz przeniesienie przetworzonego maila do Archive.
12. Ponownie uruchom Debug -> Compile VBAProject po kazdej zmianie.
