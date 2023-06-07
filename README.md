# telefonlista
Ez egy gyakorlásként készített powershell script, AD ból Excel alapú telefonlista készítéséhez. (A szintaxis szépítgetése még folyamatban, de működik) Ütemezett feladatként lehet automatizálni és mindig frissül.

Betölti az Active Directory PowerShell modult és az Excel objektumot.
Definiál egy tömböt ($tartomanyok), amely tartalmazza a kívánt tartományok neveit.
Létrehoz egy új Excel munkafüzetet és lapokat a tartományok számára.
Minden tartományhoz végrehajt egy iterációt és lekérdezi az Active Directory-ból az adott tartományhoz tartozó felhasználókat, akiknél van valamilyen telefonszám megadva.
Az adatokat feltölti az Excel táblázatba a megfelelő oszlopokba.
Beállítja a táblázatstílust és a sorok színét az adott lapokon.
Létrehoz másik lapokat és feltölti azokat a felhasználókkal, akiknél az "Iroda" mező üres vagy nem üres.
Beállítja az oszlopok szélességét az egész munkafüzetben.
Törli az "Munkalap1" nevű lapot, ha létezik.
Törli az "Oszlop1" és "Oszlop2" nevű oszlopokat a "munkalapneve" nevű lapról.
Beállítja a fejléc színét és betűszínét a megfelelő lapokon.
A szkript célja tehát az Active Directory-ból származó felhasználói adatok lekérdezése és ezek Excel táblázatba való feltöltése, valamint az Excel táblázat formázása és a lapokhoz tartozó fejléc színezése.