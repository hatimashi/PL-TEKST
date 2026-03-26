# Changelog

Wszystkie istotne zmiany w projekcie są dokumentowane w tym pliku.

Format oparty na [Keep a Changelog](https://keepachangelog.com/pl/1.0.0/).
Projekt stosuje [Semantic Versioning](https://semver.org/lang/pl/).

---
## [2.0.0] - 2026-03-26

### Dodano
- Obsługa wielu walut — nowy opcjonalny parametr: `=PL_TEKST(kwota; "EUR")`
- Obsługa EUR — euro/euro/euro + cent/centy/centów
- Obsługa USD — dolar/dolary/dolarów + cent/centy/centów
- Obsługa GBP — funt/funty/funtów + pens/pensy/pensów
- Komunikat błędu dla nieznanej waluty z listą dostępnych kodów
- Rozbudowany arkusz testowy — 76 testów w 4 zakładkach (PLN, EUR, USD, GBP)

### Zmieniono
- Całkowita refaktoryzacja kodu — podział na 5 czytelnych sekcji
- Nowa funkcja `LiczbaSlownie()` — reużywalna, niezależna od waluty
- Nowa funkcja `PobierzWalute()` — słownik walut, łatwe dodawanie nowych
- Zapis groszy zawsze jako ułamek z formą dopełniaczową (groszy/centów/pensów)
- Usunięto `Private Type` — zastąpiono tablicami dla lepszej kompatybilności VBA

---

## [1.0.1] - 2026-03-25

### Naprawiono
- Błędna odmiana jedności przy złotych — "jedna/dwie" zmienione na "jeden/dwa" (rodzaj męski)
- Błędna odmiana przy tysiącach — "dwie tysiące" zmienione na "dwa tysiące"
- Błędna odmiana grosza dla wartości 0 i 1 — zawsze "groszy" dla zapisu ułamkowego
- Błędne obliczanie groszy dla liczb zmiennoprzecinkowych (np. 0.01 dawało 00)
- Overflow dla maksymalnych wartości (999 999 999,99) przy obliczaniu groszy
- Błędne zaokrąglenie dla wartości typu 0.005 — teraz poprawnie 01/100 groszy
- Błędna odmiana milionów — "dwie miliony" zmienione na "dwa miliony"
- Błędna odmiana tysięcy dla 12000 — teraz poprawnie "dwanaście tysięcy"

### Zmieniono
- Obliczanie groszy przez typ `Currency` zamiast `Double` — eliminuje błędy zmiennoprzecinkowe
- Każdy wyraz języka polskiego jako osobna funkcja — eliminuje problemy z kodowaniem ChrW()
- Dodano 41 testów automatycznych w pliku `tests/testy_PL_TEKST.xlsx`

---

## [1.0.0] - 2026-03-24

### Dodano
- Pierwsza publiczna wersja funkcji `PL_TEKST`
- Obsługa kwot od 0 do 999 999 999,99 PLN
- Poprawna odmiana: złoty/złote/złotych
- Poprawna odmiana: grosz/grosze/groszy
- Poprawna odmiana: tysiąc/tysiące/tysięcy
- Poprawna odmiana: milion/miliony/milionów
- Obsługa liczebników żeńskich (dwie, jedna) przy tysiącach
- Grosze jako ułamek (np. 67/100 groszy)
- Polskie znaki budowane przez ChrW() — brak problemów z kodowaniem
- Plik instalacyjny .xlam
- Dokumentacja instalacji i użycia
