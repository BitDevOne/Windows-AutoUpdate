<#
.SYNOPSIS
    Skrypt do automatyzacji instalacji i zarządzania aktualizacjami systemu Windows.

.DESCRIPTION
    Ten skrypt w PowerShell służy do automatyzacji procesu wyszukiwania, pobierania i instalacji aktualizacji systemu Windows. Skrypt obsługuje różne scenariusze aktualizacji, w tym codzienne aktualizacje, aktualizacje co dwa tygodnie oraz aktualizacje wykonywane po restarcie systemu.

    Główne funkcje skryptu:
    - Automatyczne sprawdzanie dostępności nowszej wersji skryptu na GitHubie oraz jego aktualizacja.
    - Wyszukiwanie aktualizacji dostępnych do zainstalowania za pomocą usługi Windows Update.
    - Pobieranie oraz instalacja dostępnych aktualizacji, łącznie z akceptowaniem warunków licencji (EULA).
    - Wysyłanie e-maila z informacjami o zainstalowanych aktualizacjach, z danymi pobranymi z pliku XML ("settings.xml").
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

.EXAMPLE
    .\WindowsUpdateAutomation.ps1 -WindowsUpdateDayly
        Uruchamia codzienne sprawdzanie dostępnych aktualizacji i ich instalację.

    .\WindowsUpdateAutomation.ps1 -WindowsUpdate2Week
        Uruchamia proces instalacji aktualizacji co dwa tygodnie, z możliwością automatycznego restartu systemu w razie potrzeby.

    .\WindowsUpdateAutomation.ps1 -WindowsUpdateAfterReboot
        Uruchamia instalację aktualizacji po restarcie systemu.
#>


[CmdletBinding()]
param (
    [switch]$WindowsUpdateDayly,
    [switch]$WindowsUpdate2Week,
    [switch]$WindowsUpdateAfterReboot
)

# Definiowanie funkcji

#AutoUpdate
$githubver = "https://raw.githubusercontent.com/BitDevOne/Windows-AutoUpdate/refs/heads/main/Version"
$githubscript = "https://raw.githubusercontent.com/BitDevOne/Windows-AutoUpdate/refs/heads/main/WindowsUpdate.ps1"
$version = "1.2"

function UpdatesAvailable()
{
    $updateavailable = $false
    $nextversion = $null
    try
    {
        $nextversion = (New-Object System.Net.WebClient).DownloadString($githubver).Trim([Environment]::NewLine)
    }
    catch [System.Exception] 
    {
        Write-Host $_
    }
    
    Write-Host "Aktualna wersja: $version" 
    Write-Host "Nowa wersja: $nextversion"

    if ($nextversion -ne $null -and $version -ne $nextversion)
    {
        # Sprawdzenie czy aktualizacja jest rzeczywiście dostępna
        $updateavailable = $false
        $curr = $version.Split('.')
        $next = $nextversion.Split('.')
        for($i=0; $i -le ($curr.Count - 1); $i++)
        {
            if ([int]$next[$i] -gt [int]$curr[$i])
            {
                $updateavailable = $true
                break
            }
        }
    }
    return $updateavailable
}

function ProcessUpdates()
{
    if (Test-Connection 8.8.8.8 -Count 1 -Quiet)
    {
        $updatepath = "$($PWD.Path)\WindowsUpdate.ps1"
        
        if (UpdatesAvailable)
        {
            # Jeśli istnieje stary skrypt, nazwiemy go inaczej dla celów bezpieczeństwa
            if (Test-Path -Path $updatepath)
            {
                Rename-Item -Path $updatepath -NewName "WindowsUpdate_Old.ps1" -Force
            }
            
            # Pobrać i zapisać nową wersję skryptu
            try
            {
                (New-Object System.Net.WebClient).DownloadFile($githubscript, $updatepath)
                Write-Host "Pobrano nową wersję skryptu. Aktualizowanie..."
                
                # Wykonanie nowego skryptu
                Start-Process PowerShell -ArgumentList "-File `"$updatepath`"" -Wait
                
                # Usunięcie starej wersji po udanej aktualizacji
                if (Test-Path -Path "$($PWD.Path)\WindowsUpdate_Old.ps1")
                {
                    Remove-Item "$($PWD.Path)\WindowsUpdate_Old.ps1" -Force
                }
                exit
            }
            catch [System.Exception]
            {
                Write-Host "Błąd podczas pobierania nowej wersji skryptu: $_"

                # Przywrócenie starego skryptu w przypadku błędu
                if (Test-Path -Path "$($PWD.Path)\WindowsUpdate_Old.ps1")
                {
                    Rename-Item -Path "$($PWD.Path)\WindowsUpdate_Old.ps1" -NewName "WindowsUpdate.ps1" -Force
                }
            }
        }
        else
        {
            Write-Host "Brak aktualizacji."
        }
    }
    else
    {
        Write-Host "Brak połączenia z internetem."
    }
}

# Definiowanie funkcji do wysyłania e-maila
function SendEmail {
    [CmdletBinding()]
    param (
        [string]$Body
    )

    # Wczytanie pliku XML
    $xmlFilePath = ".\settings.xml"
    $xmlContent = Get-Content -Path $xmlFilePath
    $xmlDocument = [xml]$xmlContent

    # Pobranie danych z pliku XML
    $HostName = $xmlDocument.settings.HostName
    $Companyname = $xmlDocument.settings.CompanyName
    $Username = $xmlDocument.settings.MailboxFrom
    $Password = $xmlDocument.settings.MailboxFromPass
    $message = new-object Net.Mail.MailMessage;
    $message.From = $Username;
    $Recipient = $xmlDocument.settings.MailboxTo
    $message.To.Add($Recipient);
    $message.Subject = "[Windows Update] - [$Companyname] - [$HostName]"
    $message.Body = $Body
    $SMTPServer = $xmlDocument.settings.SMTPserver
    $SMTPport = $xmlDocument.settings.SMTPport
    $smtp = new-object Net.Mail.SmtpClient($SMTPServer, $SMTPport);
    $smtp.EnableSSL = $true;
    $smtp.Credentials = New-Object System.Net.NetworkCredential($Username, $Password);
    $smtp.send($message);
}

# Definiowanie funkcji do aktualizacji
function WindowsUpdateDayly {
    [CmdletBinding()]
    param ()

    # Zainicjowanie sesji z Windows Update
    $objSession = New-Object -ComObject 'Microsoft.Update.Session'
    $objSearcher = $objSession.CreateUpdateSearcher()

    # Wyszukiwanie dostępnych aktualizacji
    $search = 'IsInstalled=0 and IsHidden=0'
    $objResults = $objSearcher.Search($search)
    $Updates = $objResults.Updates

    if ($Updates.Count -eq 0) {
        Write-Host "Brak dostępnych aktualizacji do zainstalowania."
        return
    }   

    # Przygotowanie do pobrania i instalacji
    $objCollectionDownload = New-Object -ComObject 'Microsoft.Update.UpdateColl'
    $objCollectionInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'

    foreach ($Update in $Updates) {
        $Update.AcceptEula()
        $objCollectionDownload.Add($Update)
    }

    # Pobieranie aktualizacji
    $Downloader = $objSession.CreateUpdateDownloader()
    $Downloader.Updates = $objCollectionDownload
    $DownloadResult = $Downloader.Download()

    # Filtracja aktualizacji gotowych do instalacji
    foreach ($Update in $objCollectionDownload) {
        if ($Update.IsDownloaded) {
            $objCollectionInstall.Add($Update)
        }
    }

    # Instalacja aktualizacji
    $Installer = $objSession.CreateUpdateInstaller()
    $Installer.Updates = $objCollectionInstall
    $InstallResult = $Installer.Install()

    # Tworzenie listy zainstalowanych aktualizacji
    $InstalledUpdates = @()

    for ($i = 0; $i -lt $objCollectionInstall.Count; $i++) {
        $Update = $objCollectionInstall.Item($i)
        $status = $InstallResult.GetUpdateResult($i).ResultCode
        if ($status -eq 2) {  # Kod 2 oznacza 'Succeeded'
            $InstalledUpdates += $Update.Title
        }
    }

    # Wyświetlanie zainstalowanych aktualizacji
    if ($InstalledUpdates.Count -gt 0) {
        Write-Host "Zainstalowano następujące aktualizacje:"
        $InstalledUpdates | ForEach-Object { Write-Host $_ }

        # Przygotowanie treści e-maila
        $EmailBody = "Zainstalowano następujące aktualizacje:\n" + ($InstalledUpdates -join "`n")

        # Wysyłanie e-maila z listą aktualizacji
        SendEmail -Recipient $EmailRecipient -Body $EmailBody
    } else {
        Write-Host "Nie znaleziono żadnych aktualizacji do zainstalowania."
    }
}

function WindowsUpdate2Week {
    [CmdletBinding()]
    param ()

    # Zainicjowanie sesji z Windows Update
    $objSession = New-Object -ComObject 'Microsoft.Update.Session'
    $objSearcher = $objSession.CreateUpdateSearcher()

    # Wyszukiwanie dostępnych aktualizacji
    $search = 'IsInstalled=0 and IsHidden=0'
    $objResults = $objSearcher.Search($search)
    $Updates = $objResults.Updates

    if ($Updates.Count -eq 0) {
        Write-Host "Brak dostępnych aktualizacji do zainstalowania."
        return
    }   

    # Przygotowanie do pobrania i instalacji
    $objCollectionDownload = New-Object -ComObject 'Microsoft.Update.UpdateColl'
    $objCollectionInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'

    foreach ($Update in $Updates) {
        $Update.AcceptEula()
        $objCollectionDownload.Add($Update)
    }

    # Pobieranie aktualizacji
    $Downloader = $objSession.CreateUpdateDownloader()
    $Downloader.Updates = $objCollectionDownload
    $DownloadResult = $Downloader.Download()

    # Filtracja aktualizacji gotowych do instalacji
    foreach ($Update in $objCollectionDownload) {
        if ($Update.IsDownloaded) {
            $objCollectionInstall.Add($Update)
        }
    }

    # Instalacja aktualizacji
    $Installer = $objSession.CreateUpdateInstaller()
    $Installer.Updates = $objCollectionInstall
    $InstallResult = $Installer.Install()

    # Tworzenie listy zainstalowanych aktualizacji
    $InstalledUpdates = @()

    for ($i = 0; $i -lt $objCollectionInstall.Count; $i++) {
        $Update = $objCollectionInstall.Item($i)
        $status = $InstallResult.GetUpdateResult($i).ResultCode
        if ($status -eq 2) {  # Kod 2 oznacza 'Succeeded'
            $InstalledUpdates += $Update.Title
        }
    }

    # Sprawdzenie, czy jest wymagane ponowne uruchomienie
    if ($InstallResult.RebootRequired) {
        Write-Host "Wymagane jest ponowne uruchomienie systemu."

        # Wyświetlanie zainstalowanych aktualizacji
        if ($InstalledUpdates.Count -gt 0) {
            Write-Host "Zainstalowano następujące aktualizacje:"
            $InstalledUpdates | ForEach-Object { Write-Host $_ }
    
            # Przygotowanie treści e-maila
            $EmailBody = "Zainstalowano następujące aktualizacje:`n" + ($InstalledUpdates -join "`n") + "`nWymagane jest ponowne uruchomienie systemu."
    
            # Wysyłanie e-maila z listą aktualizacji
            SendEmail -Body $EmailBody

            Restart-Computer -Force
        }

    }else {
        
        # Wyświetlanie zainstalowanych aktualizacji
        if ($InstalledUpdates.Count -gt 0) {
            Write-Host "Zainstalowano następujące aktualizacje:"
            $InstalledUpdates | ForEach-Object { Write-Host $_ }
    
            # Przygotowanie treści e-maila
            $EmailBody = "Zainstalowano następujące aktualizacje:\n" + ($InstalledUpdates -join "`n")
    
            # Wysyłanie e-maila z listą aktualizacji
            SendEmail -Body $EmailBody
        }
    }
}

function WindowsUpdateAfterReboot {
    [CmdletBinding()]
    param ()

    # Zainicjowanie sesji z Windows Update
    $objSession = New-Object -ComObject 'Microsoft.Update.Session'
    $objSearcher = $objSession.CreateUpdateSearcher()

    # Wyszukiwanie dostępnych aktualizacji
    $search = 'IsInstalled=0 and IsHidden=0'
    $objResults = $objSearcher.Search($search)
    $Updates = $objResults.Updates

    if ($Updates.Count -eq 0) {
        Write-Host "Brak dostępnych aktualizacji do zainstalowania."
        return
    }   

    # Przygotowanie do pobrania i instalacji
    $objCollectionDownload = New-Object -ComObject 'Microsoft.Update.UpdateColl'
    $objCollectionInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'

    foreach ($Update in $Updates) {
        $Update.AcceptEula()
        $objCollectionDownload.Add($Update)
    }

    # Pobieranie aktualizacji
    $Downloader = $objSession.CreateUpdateDownloader()
    $Downloader.Updates = $objCollectionDownload
    $DownloadResult = $Downloader.Download()

    # Filtracja aktualizacji gotowych do instalacji
    foreach ($Update in $objCollectionDownload) {
        if ($Update.IsDownloaded) {
            $objCollectionInstall.Add($Update)
        }
    }

    # Instalacja aktualizacji
    $Installer = $objSession.CreateUpdateInstaller()
    $Installer.Updates = $objCollectionInstall
    $InstallResult = $Installer.Install()

    # Tworzenie listy zainstalowanych aktualizacji
    $InstalledUpdates = @()

    for ($i = 0; $i -lt $objCollectionInstall.Count; $i++) {
        $Update = $objCollectionInstall.Item($i)
        $status = $InstallResult.GetUpdateResult($i).ResultCode
        if ($status -eq 2) {  # Kod 2 oznacza 'Succeeded'
            $InstalledUpdates += $Update.Title
        }
    }

    # Sprawdzenie, czy jest wymagane ponowne uruchomienie
    if ($InstallResult.RebootRequired) {
        Write-Host "Wymagane jest ponowne uruchomienie systemu."

        # Wyświetlanie zainstalowanych aktualizacji
        if ($InstalledUpdates.Count -gt 0) {
            Write-Host "Zainstalowano następujące aktualizacje:"
            $InstalledUpdates | ForEach-Object { Write-Host $_ }
    
            # Przygotowanie treści e-maila
            $EmailBody = "Zainstalowano następujące aktualizacje:\n" + ($InstalledUpdates -join "`n") + "\nWymagane jest ponowne uruchomienie systemu."
    
            # Wysyłanie e-maila z listą aktualizacji
            SendEmail -Body $EmailBody

            Restart-Computer -Force
        }

    }else {
        
        # Wyświetlanie zainstalowanych aktualizacji
        if ($InstalledUpdates.Count -gt 0) {
            Write-Host "Zainstalowano następujące aktualizacje:"
            $InstalledUpdates | ForEach-Object { Write-Host $_ }
    
            # Przygotowanie treści e-maila
            $EmailBody = "Zainstalowano następujące aktualizacje:\n" + ($InstalledUpdates -join "`n")
    
            # Wysyłanie e-maila z listą aktualizacji
            SendEmail -Body $EmailBody
        }
    }
}

# Uruchomienie funkcji, jeśli przełącznik jest ustawiony
ProcessUpdates

if ($WindowsUpdateDayly) {
    WindowsUpdateDayly
} elseif ($WindowsUpdate2Week) {
    WindowsUpdate2Week
} elseif($WindowsUpdateAfterReboot) {
    WindowsUpdateAfterReboot
}else {
    Write-Host "Nie podano parametru, więc funkcja nie zostanie uruchomiona."
}



