# Szükséges powershell modul betöltése, ha előtte telepíteni kell a modult akkor a következő sorral megteheted: WIN10= Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online     SERVER= Install-WindowsFeature -Name "RSAT-AD-PowerShell" -IncludeAllSubFeature
Import-Module ActiveDirectory

# Tartományok
$tartomanyok = @("tartomany1", "tartomany2", "tartomany3")

# Excel létrehozása
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()

# Minden iterációban egy switch utasítás határozza meg a $tartomany értékétől függően a $searchBase változó értékét. Ezután az if feltételvizsgálat megnézi, hogy a $searchBase változó értéke nem-nulla-e, majd ennek függvényében lekérdezi a felhasználókat az Active Directory-ból
foreach ($tartomany in $tartomanyok) {
    switch ($tartomany) {
        "tartomany1" {
            $searchBase = "OU=valami1,DC=pelda,DC=hu"
        }
        "tartomany2" {
            $searchBase = "OU=valami2,OU=valami2,DC=pelda2,DC=hu"
        }
        "tartomany3" {
            $searchBase = "OU=valami3,DC=pelda3,DC=hu"
        }
        default {
            $searchBase = $null
        }
    }

    if ($searchBase) {
        # Felhasználók lekérése az Active Directory-ból az adott tartományban, OU alapján, szűrve csak olyanokra, akiknél van valamilyen telefonszám megadva
        $users = Get-ADUser -Filter {TelephoneNumber -like "*" -or Mobile -like "*" -or IPPhone -like "*"} -Property Name, Title, Company, Department, Office, TelephoneNumber, Mobile, IPPhone, FacsimileTelephoneNumber, Pager, EmailAddress, DisplayName -Server $tartomany -SearchBase $searchBase |
                 Select-Object Name, Title, Company, Department, Office, TelephoneNumber, Mobile, IPPhone, FacsimileTelephoneNumber, Pager, EmailAddress, DisplayName | Sort-Object DisplayName

    }
   
    # Új lap létrehozása a tartomány nevével
    $sheet = $workbook.Worksheets.Add()
    if ($tartomany -eq "tartomany1") {
        $sheet.Name = "lap neve a"
    }
    elseif ($tartomany -eq "tartomany2") {
        $sheet.Name = "lap neve b"
    }
    elseif ($tartomany -eq "tartomany3") {
        $sheet.Name = "lap neve c"
    }


    # Fejlécek beállítása
    $sheet.Cells.Item(1, 1) = "Név"
    $sheet.Cells.Item(1, 2) = "Beosztás"
    $sheet.Cells.Item(1, 3) = "Szervezet"
    $sheet.Cells.Item(1, 4) = "Osztály"
    $sheet.Cells.Item(1, 5) = "Iroda"
    $sheet.Cells.Item(1, 6) = "Mellék"
    $sheet.Cells.Item(1, 7) = "Mobil"
    $sheet.Cells.Item(1, 8) = "Vezetékes"
    $sheet.Cells.Item(1, 9) = "Faxszám"
    $sheet.Cells.Item(1, 10) = "Rendszám"
    $sheet.Cells.Item(1, 11) = "Emailcím"

    # Adatok feltöltése
    $row = 2
    foreach ($user in $users) {
        $sheet.Cells.Item($row, 1) = $user.DisplayName
        $sheet.Cells.Item($row, 2) = $user.Title
        $sheet.Cells.Item($row, 3) = $user.Company
        $sheet.Cells.Item($row, 4) = $user.Department
        $sheet.Cells.Item($row, 5) = $user.Office
        $sheet.Cells.Item($row, 6) = $user.IPPhone
        $sheet.Cells.Item($row, 7) = $user.Mobile
        $sheet.Cells.Item($row, 8) = $user.TelephoneNumber
        $sheet.Cells.Item($row, 9) = $user.FacsimileTelephoneNumber
        $sheet.Cells.Item($row, 10) = $user.Pager
        $sheet.Cells.Item($row, 11) = $user.EmailAddress
        $row++
    }
    # Táblázatstílus alkalmazása az adott lapon
    $tableRange = $sheet.Range("A1:K$row")
    $table = $sheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $tableRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
    $table.TableStyle = "TableStyleMedium1"

    # Táblázat sorok színezése az adott lapon
    $dataRows = $table.DataBodyRange.Rows.Count

    for ($i = 1; $i -le $dataRows; $i++) {
        if ($i % 2 -eq 0) {
            $table.DataBodyRange.Rows.Item($i).Interior.Color = 16770229
        } else {
            $table.DataBodyRange.Rows.Item($i).Interior.Color = -1  # Fehér szín (-1)
        }
    }
}

# másik lap1 létrehozása
$teruletekSheet = $workbook.Worksheets.Add()
$teruletekSheet.Name = "lap neve"

# másik lap1 fejlécek beállítása
$teruletekSheet.Cells.Item(1, 1) = "Név"
$teruletekSheet.Cells.Item(1, 2) = "Beosztás"
$teruletekSheet.Cells.Item(1, 3) = "Szervezet"
$teruletekSheet.Cells.Item(1, 4) = "Osztály"
$teruletekSheet.Cells.Item(1, 5) = "Mobile"
$teruletekSheet.Cells.Item(1, 6) = "Vezetékes"
$teruletekSheet.Cells.Item(1, 7) = "Faxszám"
$teruletekSheet.Cells.Item(1, 8) = "Rendszám"
$teruletekSheet.Cells.Item(1, 9) = "Emailcím"

# másik lap1 feltöltése
$row = 2
foreach ($user in $users) {
    if ([string]::IsNullOrWhiteSpace($user.Office)) {
        $teruletekSheet.Cells.Item($row, 1) = $user.DisplayName
        $teruletekSheet.Cells.Item($row, 2) = $user.Title
        $teruletekSheet.Cells.Item($row, 3) = $user.Company
        $teruletekSheet.Cells.Item($row, 4) = $user.Department
        $teruletekSheet.Cells.Item($row, 5) = $user.Mobile
        $teruletekSheet.Cells.Item($row, 6) = $user.TelephoneNumber
        $teruletekSheet.Cells.Item($row, 7) = $user.FacsimileTelephoneNumber
        $teruletekSheet.Cells.Item($row, 8) = $user.Pager
        $teruletekSheet.Cells.Item($row, 9) = $user.EmailAddress
        $row++
    }
}


# másik lap2 létrehozása
$fogarasiSheet = $workbook.Worksheets.Add()
$fogarasiSheet.Name = "lap neve2"

# másik lap2 fejlécek beállítása
$fogarasiSheet.Cells.Item(1, 1) = "Név"
$fogarasiSheet.Cells.Item(1, 2) = "Beosztás"
$fogarasiSheet.Cells.Item(1, 3) = "Szervezet"
$fogarasiSheet.Cells.Item(1, 4) = "Osztály"
$fogarasiSheet.Cells.Item(1, 5) = "Iroda"
$fogarasiSheet.Cells.Item(1, 6) = "Mellék"
$fogarasiSheet.Cells.Item(1, 7) = "Mobil"
$fogarasiSheet.Cells.Item(1, 8) = "Vezetékes"
$fogarasiSheet.Cells.Item(1, 9) = "Faxszám"
$fogarasiSheet.Cells.Item(1, 10) = "Rendszám"
$fogarasiSheet.Cells.Item(1, 11) = "Emailcím"

# másik lap2 adatok feltöltése
$row = 2
foreach ($user in $users) {
    if (![string]::IsNullOrWhiteSpace($user.Office)) {
        $fogarasiSheet.Cells.Item($row, 1) = $user.DisplayName
        $fogarasiSheet.Cells.Item($row, 2) = $user.Title
        $fogarasiSheet.Cells.Item($row, 3) = $user.Company
        $fogarasiSheet.Cells.Item($row, 4) = $user.Department
        $fogarasiSheet.Cells.Item($row, 5) = $user.Office
        $fogarasiSheet.Cells.Item($row, 6) = $user.IPPhone
        $fogarasiSheet.Cells.Item($row, 7) = $user.Mobile
        $fogarasiSheet.Cells.Item($row, 8) = $user.TelephoneNumber
        $fogarasiSheet.Cells.Item($row, 9) = $user.FacsimileTelephoneNumber
        $fogarasiSheet.Cells.Item($row, 10) = $user.Pager
        $fogarasiSheet.Cells.Item($row, 11) = $user.EmailAddress
        $row++
    }
}


# Oszlopok szélességének beállítása az egész munkafüzetben
$usedRange = $workbook.Worksheets.Item(1).UsedRange
$columnCount = $usedRange.Columns.Count

for ($i = 1; $i -le $columnCount; $i++) {
    $column = $usedRange.Columns.Item($i)
    $column.AutoFit() | Out-Null
    $columnWidth = $column.ColumnWidth
    $column.ColumnWidth = $columnWidth + 2
}
# "Munkalap1" nevű lap ellenőrzése és törlése, ha létezik
$worksheet = $workbook.Worksheets | Where-Object {$_.Name -eq "Munka1"}
if ($worksheet -ne $null) {
    $worksheet.Delete()
}

$worksheet = $workbook.Worksheets["munkalapneve"]  # a munkalap, amelyen törölni szeretnéd az oszlopokat

$columnHeaders = $worksheet.Range("A1", "Z1").Value2  # Az oszlopfejlécek tartományának beolvasása
$columnsToDelete = @()  # Törlendő oszlopok tárolására szolgáló tömb

# Oszlopok keresése, amelyekben "Oszlop1" vagy "Oszlop2" szerepel a fejlécben: erre azért van szükség, mert ahol kevesebb oszlop generálódik, ott kipótolhatja
for ($i = 1; $i -le $columnHeaders.Length; $i++) {
    if ($columnHeaders[1, $i] -eq "Oszlop1" -or $columnHeaders[1, $i] -eq "Oszlop2") {
        $columnsToDelete += $i
    }
}

# Oszlopok törlése fordított sorrendben, hogy a törlés ne befolyásolja a többi oszlop indexét
for ($i = $columnsToDelete.Length - 1; $i -ge 0; $i--) {
    $columnIndex = $columnsToDelete[$i]
    $columnLetter = [char](65 + $columnIndex - 1)  # Oszlop betűjele a sorszám alapján
    $columnRange = $worksheet.Range("${columnLetter}:${columnLetter}")
    $columnRange.Delete() | Out-Null
}
# Fejléc betűszíne
$fontColor = 0  # fekete
# Fejléc színe hexadecimális alakban
$colorMapping = @{
    "lap1neve" = 0xeba134   
    "lap2neve" = 0xeba134  
    "lap3neve" = 0xeba134  
    "lap4neve" = 0xeba134
    "lap5neve" = 0xeba134
}

# minden munkalapon végigmegy (iterate), majd beállítja a megadott fejléc színt
foreach ($worksheet in $workbook.Worksheets) {
    # lekéri a lapok neveit
    $sheetName = $worksheet.Name

    # ellenőrzi, hogy van e beállítás
    if ($colorMapping.ContainsKey($sheetName)) {
        
        $color = $colorMapping[$sheetName]

        
        $headerRange = $worksheet.Range("A1", $worksheet.Cells.Item(1, $worksheet.UsedRange.Columns.Count))

        
        $headerRange.Interior.Color = $color
        
        # fejléc betűszínének beállítása a megadott színre
        $headerRange.Font.Color = $fontColor

        # fejléc illesztése
        $headerRange.EntireColumn.AutoFit()
    }
}

# rácsvonalak színének meghatározása
$gridlineColors = @{
    "lap1neve" = 0xF5F5F5
    "lap2neve" = 0xF5F5F5  
    "lap3neve" = 0xF5F5F5  
    "lap4neve" = 0xF5F5F5 
    "lap5neve" = 0xF5F5F5
}

# minden munkalapon végigmegy (iterate) a rácsvonalak színezéséhez
foreach ($worksheet in $workbook.Worksheets) {
    
    $sheetName = $worksheet.Name

    
    if ($gridlineColors.ContainsKey($sheetName)) {
        
        $gridlineColor = $gridlineColors[$sheetName]

        
        $range = $worksheet.UsedRange
        $rowCount = $range.Rows.Count
        
        for ($i = 1; $i -le $rowCount; $i++) {
            
            if ($i % 2 -eq 0) {
                continue  # nem minden sor szélét színezi, csak minden másodikat
            }

            
            $row = $range.Rows.Item($i)
            $row.Borders.LineStyle = 0  # xlNone
            $row.Borders.Color = $gridlineColor

            
            $lastCell = $row.Cells.Item($row.Cells.Count)
            $lastCell.Borders.Item(2).LineStyle = 1  # xlContinuous
            $lastCell.Borders.Item(2).Color = 0x808080  
        }
    }
}



# lapok alsó fülének színezése
$tabColors = @{
    "lap1neve" = 16770229  
    "lap2neve" = 16770229  
    "lap3neve" = 16770229  
    "lap4neve" = 16770229 
    "lap5neve" = 16770229
}

# minden munkalapon végigmegy (iterate)
foreach ($worksheet in $workbook.Worksheets) {
    # Get the worksheet name
    $sheetName = $worksheet.Name

    
    if ($tabColors.ContainsKey($sheetName)) {
        
        $tabColor = $tabColors[$sheetName]

        
        $worksheet.Tab.Color = $tabColor
    }
}
# Iterálás a munkafüzet lapjain
foreach ($sheet in $workbook.Worksheets) {
    # Ellenőrizze, hogy a lap tartalmaz-e táblázatot
    if ($sheet.ListObjects.Count -gt 0) {
        # Táblázatok iterálása a lapon
        foreach ($table in $sheet.ListObjects) {
            # Autofilter letiltása a táblázaton
            $table.ShowAutoFilterDropDown = $false
        }
    }
}
# A sorok magasságának beállítása minden lapon
foreach ($sheet in $workbook.Worksheets) {
    $rows = $sheet.UsedRange.Rows
    $rows.RowHeight = 18  # Állítsd be a kívánt sor magasságot
    $rows.VerticalAlignment = -4108  # -4108 az Excel konstans kódja a függőleges középre igazításnak
}


# Fejléc középre igazítása és rögzítése minden lapon
foreach ($sheet in $workbook.Sheets) {
    $table = $sheet.ListObjects.Item(1)
    $headerRowRange = $table.HeaderRowRange
    
    # Fejléc középre igazítása
    $headerRowRange.VerticalAlignment = -4108  # Középre igazítás függőlegesen (-4108)
    $headerRowRange.HorizontalAlignment = -4108  # Középre igazítás vízszintesen (-4108)

    # Fejléc formázásának rögzítése és további formázás
    $headerRowRange.Locked = $true
    $headerRowRange.EntireColumn.AutoFit()
    $headerRowRange.Rows.WrapText = $true
    $headerRowRange.Rows.RowHeight = 24

    # Rögzített ablak beállítása a fejléchez
    $sheet.Activate()
    $sheet.Application.ActiveWindow.SplitRow = 1
    $sheet.Application.ActiveWindow.FreezePanes = $true
}
# Utolsó sor törlése minden lapon, ez azért kell mert +1 sor üresen generálódik a végére
foreach ($sheet in $workbook.Sheets) {
    $table = $sheet.ListObjects.Item(1)
    $dataBodyRange = $table.DataBodyRange
    
    # Utolsó sor törlése
    if ($dataBodyRange.Rows.Count -gt 0) {
        $dataBodyRange.Rows.Item($dataBodyRange.Rows.Count).Delete()
    }
}
# Aktív lap beállítása az első lapra 
$workbook.Sheets.Item(1).Activate()
# Excel mentése, itt adjuk meg a mentéshez tartozó elérési utat

$savePath = "sajat\eleresi\utad"
$excel.DisplayAlerts = $false  # Ne jelenjenek meg figyelmeztető üzenetek

# Munkafüzet minden lapjának írásvédelme
foreach ($worksheet in $workbook.Worksheets) {
    $worksheet.Protect("Jelszo123")
}
$protectionPassword = "Jelszo123"  # Ide írd be a kívánt jelszót
$saveAsOptions = [Microsoft.Office.Interop.Excel.XlSaveAsAccessMode]::xlExclusive
$workbook.SaveAs($savePath, [Type]::Missing, [Type]::Missing, $protectionPassword, $saveAsOptions)
$workbook.Close()
#Leállítja az Excel folyamatot és felszabadítja a hozzá tartozó erőforrásokat. Végül a workbook és excel változókat törli a memóriából. Az a célja, hogy megszabaduljon a fel nem használt Excel folyamatoktól és erőforrásoktól, és tisztább állapotba hozza a környezetet.
$excel.Quit()
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable workbook, excel

# PowerShell script bezárása
$host.SetShouldExit(0)
Write-Host "Az adatok sikeresen exportálva lettek a következő fájlba: $savePath"