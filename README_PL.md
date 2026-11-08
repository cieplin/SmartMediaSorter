# SmartMediaSorter & Find-MediaDuplicates

Zestaw narzędzi w PowerShellu do automatycznego sortowania, zmiany nazw zdjęć na podstawie daty wykonania z metadanych **EXIF** (`SmartMediaSorter.ps1`) oraz wykrywania i izolowania duplikatów plików (`Find-MediaDuplicates.ps1`), współdzielących wspólny moduł hashowania i bazę danych (`SmartMediaCommon.ps1` / `.smartmediasorter_db.json`).

---

## Wymagania

- **PowerShell 7+** (`pwsh`). Windows PowerShell 5.1 (`powershell.exe`) **nie jest obsługiwany**, a skrypty uruchomione w nim kończą się błędem parsowania.
- Opcjonalnie `ExifTool` w `PATH`, potrzebny tylko `SmartMediaSorter.ps1`. Bez niego skrypt używa wbudowanego czytnika EXIF.

Skrypty uruchamiaj w `pwsh`, np. `pwsh -File .\Remove-DuplicateFolders.ps1 -WhatIf`.

---

## Architektura projektu

- **`SmartMediaCommon.ps1`** – współdzielony moduł zawierający funkcje obliczania skrótów `SHA-256` plików (`Get-CachedFileHash`) i stanu katalogów (`Get-DirectoryStateHash`), normalizacji nazw plików (`Get-NormalizedBaseName`), logowania (`Write-SorterLog`) oraz odczytu/zapisu bazy danych JSON (`.smartmediasorter_db.json`).
- **`SmartMediaSorter.ps1`** – główny skrypt sortujący i przenoszący (lub kopiujący) zdjęcia z katalogu źródłowego (`X:\Photos`) do struktury docelowej (`X:\Pictures\YYYY\YYYY-MM\YYYY-MM-DD_HHMMSS_XXX.<ext>`).
- **`Find-MediaDuplicates.ps1`** – skrypt do wykrywania duplikatów (najpierw po znormalizowanej nazwie pliku i tym samym rozszerzeniu, a następnie po rozmiarze i sumie kontrolnej `SHA-256`) oraz przenoszenia wykrytych duplikatów do lokalnego podfolderu `_Duplicates` w tym samym katalogu.
- **`Remove-DuplicateFolders.ps1`** – skrypt rekurencyjnie usuwający foldery `_Duplicates` do Kosza Windows (dyski lokalne) lub do datowanego katalogu `_DuplicatesTrash` (dyski sieciowe/wymienne).

### Automatyczne ignorowanie Kosza, folderów systemowych i katalogów zaczynających się od kropki
Wszystkie skrypty (`SmartMediaSorter.ps1`, `Find-MediaDuplicates.ps1`, `Remove-DuplicateFolders.ps1`) w oparciu o wspólny moduł `SmartMediaCommon.ps1` automatycznie ignorują:
- **Katalogi ukryte/konfiguracyjne (zaczynające się od `.`):** `.git`, `.thumbnails`, `.cache`, `.sync`, `.idea`, `.vscode` itp.
- **Kosz Windows:** `$RECYCLE.BIN`, `RECYCLER`, `RECYCLED`
- **Kosz sieciowy / NAS:** `#recycle` (Synology DSM), `@Recycle` (QNAP QTS)
- **Kosz uniksowy / macOS:** `.Trash`, `.Trash-*`, `.Trashes`
- **Folder kosza projektu:** `_DuplicatesTrash`
- **Foldery systemowe:** `System Volume Information`

Foldery te są pomijane na poziomie drzewa katalogów (skrypty nie wchodzą do nich, co zapobiega sortowaniu plików usuniętych i błędom uprawnień). Próba podania takiego katalogu bezpośrednio jako parametru wejściowego (`-Path`, `-SourcePath`, `-DestinationPath`) kończy się kontrolowaną odmową uruchomienia.

---

## 1. Skrypt sortujący zdjęcia (`SmartMediaSorter.ps1`)

### Główne funkcje
1. **Struktura docelowa i nazewnictwo:**
   - Katalog źródłowy (domyślnie): `X:\Photos`
   - Katalog docelowy (domyślnie): `X:\Pictures`
   - Struktura podkatalogów: `\YYYY\YYYY-MM\`
   - Format nazwy pliku: `YYYY-MM-DD_HHMMSS_XXX.<oryginalne_rozszerzenie>` (małymi literami)
   - Przykład: `X:\Pictures\2015\2015-08\2015-08-30_114023_000.jpg`
2. **Odczyt daty wykonania (3-stopniowy mechanizm):**
   - **Krok 1 (ExifTool):** Wsadowy odczyt tagów `DateTimeOriginal`, `CreateDate`, `MediaCreateDate`, `TrackCreateDate` przez `ExifTool` (obsługa formatów JPEG, HEIC, RAW, TIFF, PNG itd.).
   - **Krok 2 (Natywny czytnik EXIF .NET):** Jeśli `ExifTool` nie jest dostępny w systemie, skrypt automatycznie odczytuje binarne nagłówki EXIF (`APP1` / `TIFF`: `DateTimeOriginal 0x9003`, `CreateDate 0x9004`, `ModifyDate 0x0132`).
   - **Krok 3 (Fallback):** Jeśli plik nie posiada metadanych EXIF, skrypt używa daty modyfikacji pliku (`LastWriteTime`) i odnotowuje ten fakt w logu.
3. **Wykrywanie duplikatów i obsługa zdjęć seryjnych (`_XXX`):**
   - Jeśli w docelowym katalogu istnieje już plik z tej samej sekundy (`YYYY-MM-DD_HHMMSS_*.ext`), skrypt porównuje rozmiar pliku oraz sumę kontrolną **SHA-256**:
     - **Identyczna suma SHA-256 (Duplikat):** Plik jest pomijany (pozostaje w katalogu źródłowym `X:\Photos`), a zdarzenie jest zapisywane w logu jako `DUPLICATE`.
     - **Różna suma SHA-256 (np. seria zdjęć w tej samej sekundzie):** Plik otrzymuje kolejny wolny trzycyfrowy numer sekwencyjny (`_001`, `_002`, itd.).
4. **Baza danych hashy katalogów źródłowych i docelowych (`.smartmediasorter_db.json`):**
   - Skrypt domyślnie tworzy i aktualizuje plik bazy danych w katalogu docelowym (`<DestinationPath>\.smartmediasorter_db.json`).
   - Przy ponownym uruchomieniu niezmienione pliki i podkatalogi korzystają z zapisanych w bazie danych informacji (pomijając ponowne wywołanie `ExifTool` oraz ponowne obliczanie `SHA-256` z dysku).

### Przykłady użycia (`SmartMediaSorter.ps1`)
```powershell
# Symulacja działania (-WhatIf)
.\SmartMediaSorter.ps1 -WhatIf

# Domyślne przenoszenie z X:\Photos do X:\Pictures
.\SmartMediaSorter.ps1

# Kopiowanie zamiast przenoszenia
.\SmartMediaSorter.ps1 -CopyOnly

# Wymuszenie pełnej przebudowy bazy danych hashy
.\SmartMediaSorter.ps1 -RebuildDatabase
```

---

## 2. Skrypt wykrywania duplikatów (`Find-MediaDuplicates.ps1`)

### Zasada działania
1. **Różne rozszerzenia = różne pliki:** Pliki o różnych rozszerzeniach (np. `.jpg` vs `.png` lub `.cr2` vs `.jpg`) nigdy nie są uznawane za duplikaty.
2. **Krok 1 – Podpowiedź z nazwy pliku:** Pliki o tym samym rozszerzeniu są grupowane według znormalizowanej nazwy bazowej (`Get-NormalizedBaseName`), która łączy oryginał z typowymi przyrostkami kopii i serii (np. `DSC_5000.jpg`, `DSC_5000 (1).jpg`, `DSC_5000 - kopia.jpg`, `DSC_5000_copy.jpg`, a także `2015-08-30_114023_000.jpg` i `2015-08-30_114023_001.jpg`).
3. **Krok 2 – Weryfikacja zawartości (Rozmiar + SHA-256):** Wewnątrz grupy o pasującej nazwie porównywany jest rozmiar w bajtach (`Length`), a następnie obliczana lub pobierana z bazy danych `.smartmediasorter_db.json` suma kontrolna `SHA-256`.
4. **Krok 3 – Przeniesienie do folderu `_Duplicates` w tym samym katalogu:** Plik o najczystszej/najkrótszej nazwie (np. `DSC_5000.jpg`) zostaje zachowany na swoim miejscu jako oryginał, natomiast wszystkie potwierdzone duplikaty są przenoszone do podfolderu `_Duplicates` utworzonego bezpośrednio w katalogu, w którym znajdował się dany duplikat (np. `X:\Photos\AlbumA\_Duplicates\`).

### Przykłady użycia (`Find-MediaDuplicates.ps1`)
```powershell
# Symulacja wykrywania duplikatów (-WhatIf) w domyślnym katalogu X:\Photos
.\Find-MediaDuplicates.ps1 -WhatIf

# Wykrycie i przeniesienie duplikatów w X:\Photos do lokalnych podfolderów _Duplicates
.\Find-MediaDuplicates.ps1 -Path "X:\Photos"

# Skanowanie katalogu X:\Pictures z własną nazwą podfolderu na duplikaty
.\Find-MediaDuplicates.ps1 -Path "X:\Pictures" -DuplicateFolderName "_Duplikaty"
```

### Parametry `Find-MediaDuplicates.ps1`

| Parametr | Typ | Wartość domyślna | Opis |
| :--- | :--- | :--- | :--- |
| `-Path` | `string` | `"X:\Photos"` | Katalog do przeskanowania pod kątem duplikatów. |
| `-DuplicateFolderName` | `string` | `"_Duplicates"` | Nazwa podfolderu tworzonego w katalogu duplikatu, do którego przenoszone są duplikaty. |
| `-Recurse` | `bool` | `$true` | Przeszukuje również podkatalogi (z automatycznym pominięciem folderów `_Duplicates`). |
| `-DatabasePath` | `string` | `<Path>\.smartmediasorter_db.json` | Ścieżka do współdzielonej bazy danych hashy JSON. |
| `-NoDatabase` | `switch` | `$false` | Wyłącza odczyt i zapis bazy danych hashy. |
| `-RebuildDatabase` | `switch` | `$false` | Ignoruje zapisane w bazie hashe i przelicza je na nowo. |
| `-IncludeExtensions` | `string[]` | `@()` (wszystkie pliki) | Opcjonalna lista rozszerzeń plików do filtrowania. |
| `-ThrottleLimit` | `int` | `8` | Maksymalna liczba równoległych wątków (`RunspacePool`) skanujących i wyliczających hashe dla poszczególnych podkatalogów (`1` = tryb jednowątkowy). |
| `-DirectoryTimeoutSeconds` | `int` | `300` | Maksymalny czas przetwarzania pojedynczego podkatalogu (w sekundach), po którym zawieszony wątek jest przerywany i logowany z nazwą pliku (`0` = brak limitu). |
| `-LogPath` | `string` | `FindMediaDuplicates_<timestamp>.log` | Ścieżka do pliku logu. |
| `-WhatIf` | `switch` | `$false` | Tryb symulacji – pokazuje wykryte duplikaty bez przenoszenia plików. |

---

## 3. Skrypt usuwania folderów duplikatów (`Remove-DuplicateFolders.ps1`)

Rekurencyjnie wyszukuje w podanym katalogu wszystkie foldery `_Duplicates` (utworzone przez `Find-MediaDuplicates.ps1`) i usuwa je razem z zawartością **w sposób umożliwiający odzyskanie**.

### Zasada działania
1. **Wyszukiwanie:** przejście drzewa katalogów bez wchodzenia w dowiązania (junction/symlink). Nazwa folderu porównywana jest bez rozróżniania wielkości liter, a skrypt nie wchodzi do wnętrza znalezionych folderów `_Duplicates`.
2. **Sposób usuwania (`-Mode Auto`, domyślnie):**
   - **Dysk lokalny (`Fixed`)** – foldery trafiają do **Kosza Windows**.
   - **Dysk sieciowy (np. `X:`), wymienny lub ścieżka UNC** – Windows nie ma na nich Kosza (pliki zostałyby usunięte trwale), więc foldery są **przenoszone** (szybka zmiana nazwy w obrębie woluminu) do:
     `<root>\_DuplicatesTrash\<yyyyMMdd_HHmmss>\<ścieżka względem roota>`, np. `X:\_DuplicatesTrash\20261108_142152\Photos\2020\Album\_Duplicates`.
     Po sprawdzeniu zawartości folder `X:\_DuplicatesTrash` usuwa się ręcznie.
3. **Bezpieczeństwo:**
   - Przed każdym folderem skrypt pyta o potwierdzenie (`ConfirmImpact = High`; odpowiedź `A` oznacza „Tak dla wszystkich”). Przełącznik `-Force` pomija pytania.
   - `-WhatIf` wyświetla listę folderów z liczbą plików i rozmiarem, niczego nie usuwając.
   - Folder kosza jest pomijany przy skanowaniu, a foldery `_Duplicates` będące dowiązaniami są pomijane z ostrzeżeniem.
   - Skrypt odmawia działania, jeśli sam `-Path` jest folderem `_Duplicates`.
   - Usunięcie tych folderów nie wymaga zmian w bazie `.smartmediasorter_db.json`, bo `Find-MediaDuplicates.ps1` usuwa rekordy plików już w chwili przenoszenia ich do `_Duplicates`.

### Przykłady użycia (`Remove-DuplicateFolders.ps1`)
```powershell
# Podgląd folderów do usunięcia w X:\Photos (domyślnie)
.\Remove-DuplicateFolders.ps1 -WhatIf

# Usunięcie z potwierdzeniem dla każdego folderu
.\Remove-DuplicateFolders.ps1

# Usunięcie bez pytań w innym katalogu
.\Remove-DuplicateFolders.ps1 -Path "X:\Pictures" -Force
```

### Parametry `Remove-DuplicateFolders.ps1`

| Parametr | Typ | Wartość domyślna | Opis |
| :--- | :--- | :--- | :--- |
| `-Path` | `string` | `"X:\Photos"` | Katalog przeszukiwany rekurencyjnie. |
| `-DuplicateFolderName` | `string` | `"_Duplicates"` | Nazwa folderów do usunięcia. |
| `-Mode` | `Auto` / `RecycleBin` / `TrashFolder` | `Auto` | Sposób usuwania (`RecycleBin` jest dozwolony tylko na dyskach lokalnych). |
| `-TrashPath` | `string` | `<root>\_DuplicatesTrash` | Katalog kosza dla trybu `TrashFolder` (musi leżeć na tym samym woluminie). |
| `-Force` | `switch` | `$false` | Pomija pytania o potwierdzenie. |
| `-ThrottleLimit` | `int` | `8` | Maksymalna liczba równoległych wątków skanujących podkatalogi pierwszego poziomu i zliczających rozmiar folderów `_Duplicates`. |
| `-LogPath` | `string` | `RemoveDuplicateFolders_<timestamp>.log` | Ścieżka do pliku logu. |
| `-PassThru` | `switch` | `$false` | Zwraca obiekt ze statystykami. |
| `-WhatIf` | `switch` | `$false` | Tryb symulacji. |

---

## Uruchamianie testów automatycznych

Skrypt testów integracyjnych weryfikuje wszystkie narzędzia (`SmartMediaSorter.ps1`, `Find-MediaDuplicates.ps1`, `Remove-DuplicateFolders.ps1`) wraz ze wspólnym modułem `SmartMediaCommon.ps1`:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-SmartMediaSorter.ps1
```
