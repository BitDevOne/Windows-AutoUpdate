<#
.SYNOPSIS
    Skrypt do automatyzacji instalacji i zarządzania aktualizacjami systemu Windows oraz konfigurowania zadań harmonogramu do regularnego pobierania plików z GitHuba.

.DESCRIPTION
    Ten skrypt w PowerShell służy do automatyzacji procesu wyszukiwania, pobierania i instalacji aktualizacji systemu Windows. Skrypt obsługuje różne scenariusze aktualizacji, w tym codzienne aktualizacje, aktualizacje co dwa tygodnie oraz aktualizacje wykonywane po restarcie systemu.

    Główne funkcje skryptu:
    - Automatyczne sprawdzanie dostępności nowszej wersji skryptu na GitHubie oraz jego aktualizacja.
    - Wyszukiwanie aktualizacji dostępnych do zainstalowania za pomocą usługi Windows Update.
    - Pobieranie oraz instalacja dostępnych aktualizacji, łącznie z akceptowaniem warunków licencji (EULA).
    - Wysyłanie e-maila z informacjami o zainstalowanych aktualizacjach, z danymi pobranymi z pliku XML ("settings.xml").
    - Pobieranie plików z GitHuba, w tym głównego skryptu oraz pliku ustawień.
    - Konfigurowanie zadań harmonogramu do automatycznego uruchamiania skryptu w określonych odstępach czasu:
      1. Uruchamianie codziennie o 1 w nocy.
      2. Uruchamianie co 2 tygodnie o 2 w nocy w zależności od konfiguracji (środa lub czwartek).
      3. Uruchamianie skryptu po uruchomieniu komputera.
    - Obsługa trzech trybów pracy:
      1. `-WindowsUpdateDayly` - Codzienne sprawdzanie i instalacja aktualizacji.
      2. `-WindowsUpdate2Week` - Instalacja aktualizacji co dwa tygodnie, z możliwością automatycznego restartu komputera.
      3. `-WindowsUpdateAfterReboot` - Instalacja aktualizacji po restarcie systemu.
    - Sprawdzanie dostępności połączenia internetowego przed sprawdzeniem i pobraniem aktualizacji.

.PARAMETERS
    -WindowsUpdateDayly
        Przełącznik, który uruchamia proces codziennego sprawdzania i instalacji aktualizacji systemu Windows.

    -WindowsUpdate2Week
        Przełącznik, który uruchamia proces aktualizacji wykonywany co dwa tygodnie. W razie potrzeby, po instalacji aktualizacji może być wymagane automatyczne ponowne uruchomienie komputera.

    -WindowsUpdateAfterReboot
        Przełącznik, który uruchamia proces instalacji aktualizacji po restarcie systemu Windows.

.NOTES
    - Przed uruchomieniem skryptu, wymagane jest posiadanie pliku "settings.xml" z odpowiednimi danymi konfiguracyjnymi dotyczącymi serwera SMTP oraz informacji o firmie i komputerze.
    - Skrypt wymaga uprawnień administratora, aby móc instalować aktualizacje i ponownie uruchamiać komputer.
    - Skrypt łączy się z internetem w celu pobrania nowszej wersji samego siebie oraz skryptów pomocniczych.
    - Skrypt umożliwia również konfigurację zadań harmonogramu do automatyzacji uruchamiania aktualizacji systemu Windows w określonych odstępach czasu.

.EXAMPLE
    .\WindowsUpdateAutomation.ps1 -WindowsUpdateDayly
        Uruchamia codzienne sprawdzanie dostępnych aktualizacji i ich instalację.

    .\WindowsUpdateAutomation.ps1 -WindowsUpdate2Week
        Uruchamia proces instalacji aktualizacji co dwa tygodnie, z możliwością automatycznego restartu systemu w razie potrzeby.

    .\WindowsUpdateAutomation.ps1 -WindowsUpdateAfterReboot
        Uruchamia instalację aktualizacji po restarcie systemu.

    .\WindowsUpdateAutomation.ps1
        Skrypt zapyta o ścieżkę do zapisania plików pobranych z GitHuba, skonfiguruje pobranie skryptu oraz pliku XML i utworzy odpowiednie zadania harmonogramu zadań.
#>

# Określenie URL do pobrania plików z GitHuba
$scriptUrl = "https://raw.githubusercontent.com/BitDevOne/Windows-AutoUpdate/refs/heads/main/WindowsUpdate.ps1"
$xmlUrl = "https://raw.githubusercontent.com/BitDevOne/Windows-AutoUpdate/refs/heads/main/settings.xml"

# Sprawdzenie, czy skrypt działa jako administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Host "Skrypt nie został uruchomiony z uprawnieniami administratora. Uruchamianie ponownie jako administrator..."
    Start-Sleep -Seconds 3
    Start-Process -FilePath PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Zapytanie użytkownika o ścieżkę do zapisania pobranych plików
$savePath = Read-Host "Podaj ścieżkę, gdzie zapisać pliki"

# Sprawdzenie, czy podana ścieżka jest pusta
while ([string]::IsNullOrWhiteSpace($savePath)) {
    Write-Host "Podana ścieżka jest pusta. Proszę podaj poprawną ścieżkę."
    $savePath = Read-Host "Podaj ścieżkę, gdzie zapisać pliki"
}

# Sprawdzenie, czy folder istnieje
if (-not (Test-Path -Path $savePath -PathType Container)) {
    $createFolder = Read-Host "Folder nie istnieje. Czy utworzyć folder? (tak/nie)"
    if ($createFolder -eq "tak") {
        New-Item -Path $savePath -ItemType Directory | Out-Null
        Write-Host "Folder utworzony."
    } else {
        Write-Host "Nie utworzono folderu. Skrypt zakończył działanie."
        exit
    }
}

# Pobieranie plików z GitHuba
Invoke-WebRequest -Uri $scriptUrl -OutFile "$savePath\WindowsUpdate.ps1"
Invoke-WebRequest -Uri $xmlUrl -OutFile "$savePath\settings.xml"

# Wczytanie danych z pliku XML
[xml]$xmlData = Get-Content "$savePath\settings.xml"

# Pytanie użytkownika o wartości zmiennych z XML
Write-Host "Podaj wartości zmiennych z pliku XML:"
foreach ($node in $xmlData.DocumentElement.ChildNodes) {
    $newValue = Read-Host "Podaj wartość dla $($node.Name) (obecna: $($node.InnerText))"
    
    # Sprawdzenie, czy wartość jest pusta, jeśli nie podano nowej, zachowujemy starą wartość
    while ([string]::IsNullOrWhiteSpace($newValue) -and [string]::IsNullOrWhiteSpace($node.InnerText)) {
        Write-Host "Podana wartość jest pusta. Proszę podaj poprawną wartość dla $($node.Name)."
        $newValue = Read-Host "Podaj wartość dla $($node.Name) (obecna: $($node.InnerText))"
    }
    
    if (-not [string]::IsNullOrWhiteSpace($newValue)) {
        $node.InnerText = $newValue
    }
}

# Zapisanie zaktualizowanego pliku XML
$xmlData.Save("$savePath\settings.xml")

# Definiowanie akcji dla zadania
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DontStopOnIdleEnd
$principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# Zadanie 1: Uruchamianie codziennie o 1 w nocy
$actionDaily = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File `"$savePath\WindowsUpdate.ps1`" -WindowsUpdateDayly" -WorkingDirectory $savePath
$triggerDaily = New-ScheduledTaskTrigger -Daily -At "01:00AM"
Register-ScheduledTask -Action $actionDaily -Trigger $triggerDaily -Settings $settings -Principal $principal -TaskName "WindowsUpdateDayly" -Description "Uruchamianie skryptu codziennie o 1 w nocy"

# Zadanie 2: Uruchamianie co 2 tygodnie od podanej daty przez użytkownika
# Wczytanie pliku XML
$xmlFilePath = "$savePath\settings.xml"
$xmlContent = Get-Content -Path $xmlFilePath
$xmlDocument = [xml]$xmlContent

# Pobranie danych z pliku XML
$Companyname = $xmlDocument.settings.CompanyName
if($Companyname -like "*HOST*") {
    $triggerBiweekly = New-ScheduledTaskTrigger -Weekly -WeeksInterval 2 -At "02:00" -DaysOfWeek Thursday
}else{
    $triggerBiweekly = New-ScheduledTaskTrigger -Weekly -WeeksInterval 2 -At "02:00" -DaysOfWeek Wednesday
}
$actionWindowsUpdate2Week = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File `"$savePath\WindowsUpdate.ps1`" -WindowsUpdate2Week" -WorkingDirectory $savePath
Register-ScheduledTask -Action $actionWindowsUpdate2Week -Trigger $triggerBiweekly -Settings $settings -Principal $principal -TaskName "WindowsUpdate2Week" -Description "Uruchamianie skryptu co 2 tygodnie od podanej daty"

# Zadanie 3: Uruchamianie po starcie komputera
$actionWindowsUpdateAfterReboot = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File `"$savePath\WindowsUpdate.ps1`" -WindowsUpdateAfterReboot" -WorkingDirectory $savePath
$triggerAtStartup = New-ScheduledTaskTrigger -AtStartup
Register-ScheduledTask -Action $actionWindowsUpdateAfterReboot -Trigger $triggerAtStartup -Settings $settings -Principal $principal -TaskName "WindowsUpdateAfterReboot" -Description "Uruchamianie skryptu po uruchomieniu komputera"

Write-Host "Zadania zostały dodane do harmonogramu zadań."
