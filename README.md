# PL-TEKST

> Polska funkcja VBA dla Microsoft Excel — zamiana liczb na tekst słowny w języku polskim.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-1.0.1-blue.svg)](CHANGELOG.md)
[![Excel](https://img.shields.io/badge/Microsoft%20Excel-2010%2B-green.svg)]()

---

## 📋 Opis

**PL-TEKST** to darmowy dodatek do Microsoft Excel, który dodaje funkcję `=PL_TEKST()` zamieniającą dowolną liczbę na jej zapis słowny w języku polskim — wraz z poprawną odmianą (złoty/złote/złotych, grosz/grosze/groszy, tysiąc/tysiące/tysięcy itd.).

Funkcja przydatna przy generowaniu faktur, umów, czeków i innych dokumentów finansowych.

### Przykłady

| Formuła | Wynik |
|---|---|
| `=PL_TEKST(1)` | Jeden złoty (00/100 groszy) |
| `=PL_TEKST(1234.67)` | Jeden tysiąc dwieście trzydzieści cztery złote (67/100 groszy) |
| `=PL_TEKST(21)` | Dwadzieścia jeden złotych (00/100 groszy) |
| `=PL_TEKST(1000000)` | Jeden milion złotych (00/100 groszy) |
| `=PL_TEKST(0)` | Zero złotych (00/100 groszy) |

---

## 🚀 Instalacja

### Metoda 1 — Dodatek .xlam (zalecana)

1. Pobierz plik `PL_TEKST.xlam` z sekcji [Releases](../../releases)
2. W Excelu: **Plik → Opcje → Dodatki → Przejdź**
3. Kliknij **Przeglądaj** i wskaż pobrany plik `.xlam`
4. Zaznacz checkbox przy **PL_TEKST** → **OK**
5. Funkcja `=PL_TEKST()` jest teraz dostępna we wszystkich plikach

### Metoda 2 — Import modułu VBA

1. Pobierz plik `src/PL_TEKST.bas`
2. Otwórz Excel i naciśnij `ALT + F11`
3. W edytorze VBA: prawy klik na projekt → **Import File**
4. Wskaż plik `PL_TEKST.bas`
5. Zapisz plik jako `.xlsm`

### Metoda 3 — Personal Macro Workbook (funkcja globalna)

Jeśli chcesz używać `=PL_TEKST()` we **wszystkich** plikach bez instalowania dodatku:

1. Otwórz Excel i naciśnij `ALT + F11`
2. W lewym panelu znajdź **PERSONAL.XLSB** → prawy klik → **Insert → Module**
3. Wklej zawartość pliku `src/PL_TEKST.bas`
4. Zapisz (`CTRL + S`)

---

## 📖 Użycie

```
=PL_TEKST(liczba)
```

| Parametr | Typ | Opis |
|---|---|---|
| `liczba` | Liczba | Kwota w złotych (max 999 999 999,99) |

### Obsługiwane wartości
- Liczby od `0` do `999 999 999,99`
- Grosze zaokrąglane do 2 miejsc po przecinku
- Liczby całkowite i dziesiętne

### Kody błędów

| Komunikat | Przyczyna |
|---|---|
| `Blad: nieprawidlowa wartosc` | Komórka zawiera tekst lub jest pusta |
| `Blad: ujemna liczba` | Podano liczbę ujemną |
| `Blad: liczba zbyt duza` | Przekroczono limit 999 999 999,99 |

---

## 🗺️ Roadmap

- [x] v1.0 — podstawowa funkcja PLN
- [ ] v1.1 — obsługa EUR, USD, GBP
- [ ] v1.2 — format prawny dla faktur
- [ ] v2.0 — API webowe (Python/FastAPI)
- [ ] v2.1 — prosta aplikacja webowa

---

## 🤝 Współpraca

Pull requesty są mile widziane! Jeśli chcesz dodać nową walutę, poprawić odmianę lub dodać testy — śmiało.

1. Zrób **Fork** repozytorium
2. Stwórz branch: `git checkout -b feature/nowa-waluta`
3. Zatwierdź zmiany: `git commit -m 'Dodaj obsługę EUR'`
4. Wypchnij: `git push origin feature/nowa-waluta`
5. Otwórz **Pull Request**

---

## 📄 Licencja

[MIT](LICENSE) — używaj swobodnie, także komercyjnie.

---

## 🇬🇧 English summary

**PL-TEKST** is a free VBA add-in for Microsoft Excel that converts numbers to Polish words, including correct grammatical inflection. Useful for invoices, contracts, and financial documents in Polish.

Install via `.xlam` add-in or import `PL_TEKST.bas` directly into your VBA project.
