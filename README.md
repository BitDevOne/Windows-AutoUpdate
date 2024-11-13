# Windows Update Automation

## Opis projektu

Ten projekt zawiera dwa skrypty PowerShell, których celem jest automatyzacja procesu instalacji i zarządzania aktualizacjami systemu Windows oraz konfiguracja harmonogramu zadań do regularnego pobierania i uruchamiania tych skryptów.

Główne funkcjonalności obejmują:

- **Automatyczne pobieranie aktualizacji Windows**: Codziennie, co dwa tygodnie, lub po restarcie systemu, skrypty wyszukują i instalują dostępne aktualizacje.
- **Konfiguracja zadań w harmonogramie zadań**: Automatyczne uruchamianie skryptu w wybranych terminach przy użyciu harmonogramu zadań systemu Windows.
- **Aktualizacje z GitHuba**: Skrypt sprawdza dostępność nowszych wersji na GitHubie i automatycznie pobiera je, jeśli są dostępne.
- **Wysyłanie e-maila z wynikami**: Po instalacji aktualizacji, skrypt wysyła raport na e-mail, korzystając z konfiguracji zapisanej w pliku XML.

## Skrypty

### 1. `WindowsUpdate.ps1`

Ten skrypt służy do zarządzania procesem aktualizacji systemu Windows. Obsługuje trzy różne scenariusze instalacji aktualizacji:

1. **Codzienne aktualizacje** (`-WindowsUpdateDayly`) - Skrypt codziennie sprawdza dostępność nowych aktualizacji i instaluje je, jeśli są dostępne.
2. **Co dwa tygodnie** (`-WindowsUpdate2Week`) - Skrypt instaluje aktualizacje co dwa tygodnie. Jeśli po instalacji wymagany jest restart, system jest automatycznie uruchamiany ponownie.
3. **Po restarcie** (`-WindowsUpdateAfterReboot`) - Skrypt uruchamia się i instaluje aktualizacje po restarcie komputera.

Dodatkowo skrypt konfiguruje zadania w harmonogramie zadań, aby automatycznie uruchamiać skrypty w określonych porach, np. codziennie o godzinie 1:00 w nocy, co dwa tygodnie lub po starcie systemu.

### 2. `Setup.ps1`

Ten skrypt odpowiada za pobieranie skryptu `WindowsUpdate.ps1` oraz pliku konfiguracyjnego XML z repozytorium GitHub. Użytkownik jest proszony o podanie ścieżki, w której pliki mają zostać zapisane. Skrypt sprawdza, czy użytkownik ma odpowiednie uprawnienia administratora oraz czy folder do zapisania plików istnieje. Jeśli nie istnieje, zostaje utworzony.

Dodatkowo, `Setup.ps1` konfiguruje zadania w harmonogramie zadań do uruchamiania skryptu `WindowsUpdate.ps1` w określonych odstępach czasu.

## Wymagania

- **System operacyjny Windows**: Skrypty są przeznaczone dla systemu Windows.
- **PowerShell**: Wymagane jest użycie PowerShell z uprawnieniami administratora.
- **Połączenie z internetem**: Wymagane do pobierania skryptów z GitHuba oraz do wyszukiwania aktualizacji Windows.
- **Plik konfiguracyjny `settings.xml`**: Plik konfiguracyjny, który zawiera informacje dotyczące serwera SMTP, wymagane do wysyłania raportów e-mail.

## Instrukcja uruchomienia

1. **Pobierz repozytorium**: Skopiuj repozytorium z GitHuba na lokalny komputer.

2. **Uruchom `Setup.ps1`**: Skrypt poprosi o podanie ścieżki do zapisania pobranych plików oraz skonfiguruje zadania harmonogramu do automatyzacji procesu aktualizacji.

   ```powershell
   .\Setup.ps1
   ```

3. **Uruchom `WindowsUpdate.ps1`**: Możesz ręcznie uruchomić skrypt, podając odpowiedni przełącznik (np. `-WindowsUpdateDayly`, `-WindowsUpdate2Week`, `-WindowsUpdateAfterReboot`). Skrypt można też skonfigurować do automatycznego uruchamiania przez zadania harmonogramu.

   ```powershell
   .\WindowsUpdate.ps1 -WindowsUpdateDayly
   ```

4. **Konfiguracja e-maila**: Upewnij się, że plik `settings.xml` zawiera poprawne dane konfiguracyjne dotyczące serwera SMTP, aby możliwe było wysyłanie raportów o zainstalowanych aktualizacjach.

## Przykłady użycia

- Uruchomienie codziennego sprawdzania i instalacji aktualizacji:
  ```powershell
  .\WindowsUpdate.ps1 -WindowsUpdateDayly
  ```

- Uruchomienie aktualizacji co dwa tygodnie, z możliwością automatycznego restartu:
  ```powershell
  .\WindowsUpdate.ps1 -WindowsUpdate2Week
  ```

- Uruchomienie aktualizacji po restarcie systemu:
  ```powershell
  .\WindowsUpdate.ps1 -WindowsUpdateAfterReboot
  ```

## Uwagi

- Skrypt wymaga uruchomienia z uprawnieniami administratora.
- Zadania harmonogramu zadań są konfigurowane automatycznie przez skrypt `Setup.ps1` i obejmują codzienne oraz dwutygodniowe aktualizacje, jak również uruchamianie po starcie systemu.
- Konfiguracja SMTP w pliku `settings.xml` jest niezbędna, jeśli chcesz otrzymywać raporty o zainstalowanych aktualizacjach.


## Wkład

Jeśli chcesz pomóc w rozwoju tego projektu, proszę o przesłanie `pull requestów` lub zgłoszenie problemów (`issues`). Każdy wkład jest mile widziany!

## Autor

Projekt został stworzony przez BitDevOne. W razie pytań proszę o kontakt poprzez GitHuba lub adres e-mail podany w pliku `settings.xml`.

---
Dziękuję za zainteresowanie tym projektem! Mam nadzieję, że ułatwi on zarządzanie aktualizacjami w systemie Windows w sposób automatyczny i wygodny.