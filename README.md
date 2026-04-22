# PL-TEKST

> Darmowy dodatek VBA dla Microsoft Excel zamieniający kwoty pieniężne na pełny zapis słowny w języku polskim — z poprawną odmianą złotych, euro, dolarów i funtów. Przydatny przy wystawianiu faktur, umów i dokumentów finansowych.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-3.0.0-blue.svg)](CHANGELOG.md)
[![Excel](https://img.shields.io/badge/Microsoft%20Excel-2010%2B-green.svg)]()
[![API](https://img.shields.io/badge/API-Railway-blueviolet.svg)](https://pl-tekst-production.up.railway.app/docs)

---

## Opis

**PL-TEKST** to darmowy projekt zawierający:

- **=PL_TEKST()** — funkcja Excel, zapis kwoty z groszami jako ułamek
- **=PL_TEKST_FAKTURA()** — funkcja Excel, pełny zapis słowny dla dokumentów prawnych
- **REST API** — publiczne API dostępne online

Wszystko obsługuje waluty PLN, EUR, USD i GBP z poprawną polską odmianą.

---

## Funkcja PL_TEKST

Zapis kwoty z groszami jako ułamek (XX/100).

```
=PL_TEKST(kwota)
=PL_TEKST(kwota; "waluta")
```

| Parametr | Opis |
|---|---|
| kwota | Kwota do zamiany (max 999 999 999,99) |
| waluta | Opcjonalnie: PLN (domyślnie), EUR, USD, GBP |

### Przykłady

| Formuła | Wynik |
|---|---|
| =PL_TEKST(1) | Jeden złoty (00/100 groszy) |
| =PL_TEKST(1234,67) | Tysiąc dwieście trzydzieści cztery złote (67/100 groszy) |
| =PL_TEKST(21) | Dwadzieścia jeden złotych (00/100 groszy) |
| =PL_TEKST(1000000) | Jeden milion złotych (00/100 groszy) |
| =PL_TEKST(0) | Zero złotych (00/100 groszy) |
| =PL_TEKST(1234,67; "EUR") | Tysiąc dwieście trzydzieści cztery euro (67/100 centów) |
| =PL_TEKST(1234,67; "USD") | Tysiąc dwieście trzydzieści cztery dolary (67/100 centów) |
| =PL_TEKST(1234,67; "GBP") | Tysiąc dwieście trzydzieści cztery funty (67/100 pensów) |

---

## Funkcja PL_TEKST_FAKTURA

Pełny zapis słowny kwoty z groszami zapisanymi słownie — przeznaczony do faktur, umów i dokumentów prawnych.

```
=PL_TEKST_FAKTURA(kwota)
=PL_TEKST_FAKTURA(kwota; "waluta")
```

| Parametr | Opis |
|---|---|
| kwota | Kwota do zamiany (max 999 999 999,99) |
| waluta | Opcjonalnie: PLN (domyślnie), EUR, USD, GBP |

### Przykłady

| Formuła | Wynik |
|---|---|
| =PL_TEKST_FAKTURA(1234,67) | Tysiąc dwieście trzydzieści cztery złote i sześćdziesiąt siedem groszy |
| =PL_TEKST_FAKTURA(1000) | Tysiąc złotych i zero groszy |
| =PL_TEKST_FAKTURA(0,01) | Zero złotych i jeden grosz |
| =PL_TEKST_FAKTURA(1,01) | Jeden złoty i jeden grosz |
| =PL_TEKST_FAKTURA(1234,67; "EUR") | Tysiąc dwieście trzydzieści cztery euro i sześćdziesiąt siedem centów |
| =PL_TEKST_FAKTURA(1234,67; "USD") | Tysiąc dwieście trzydzieści cztery dolary i sześćdziesiąt siedem centów |
| =PL_TEKST_FAKTURA(1234,67; "GBP") | Tysiąc dwieście trzydzieści cztery funty i sześćdziesiąt siedem pensów |

---

## API

PL-TEKST udostępnia publiczne REST API dostępne pod adresem:

**https://pl-tekst-production.up.railway.app**

Dokumentacja interaktywna: **https://pl-tekst-production.up.railway.app/docs**

### Endpointy

| Metoda | Endpoint | Opis |
|---|---|---|
| GET | /pl-tekst | Zapis z ułamkiem groszy |
| GET | /pl-tekst-faktura | Pełny zapis słowny |
| POST | /pl-tekst | Zapis z ułamkiem groszy |
| POST | /pl-tekst-faktura | Pełny zapis słowny |

### Parametry

| Parametr | Typ | Opis |
|---|---|---|
| kwota | liczba | Kwota do zamiany (max 999 999 999,99) |
| waluta | tekst | Opcjonalnie: PLN (domyślnie), EUR, USD, GBP |

### Przykłady GET

```
GET /pl-tekst?kwota=1234.67&waluta=PLN
{"kwota":1234.67,"waluta":"PLN","wynik":"Tysiąc dwieście trzydzieści cztery złote (67/100 groszy)"}

GET /pl-tekst?kwota=1234.67&waluta=EUR
{"kwota":1234.67,"waluta":"EUR","wynik":"Tysiąc dwieście trzydzieści cztery euro (67/100 centów)"}

GET /pl-tekst-faktura?kwota=1234.67&waluta=PLN
{"kwota":1234.67,"waluta":"PLN","wynik":"Tysiąc dwieście trzydzieści cztery złote i sześćdziesiąt siedem groszy"}
```

### Przykład POST

```json
POST /pl-tekst
Content-Type: application/json

{
  "kwota": 1234.67,
  "waluta": "PLN"
}
```

---

## Instalacja Excel

### Metoda 1 — Dodatek .xlam (zalecana)

1. Pobierz plik PL_TEKST.xlam z sekcji Releases
2. W Excelu: Plik → Opcje → Dodatki → Przejdź
3. Kliknij Przeglądaj i wskaż pobrany plik .xlam
4. Zaznacz checkbox przy PL_TEKST → OK
5. Obie funkcje są teraz dostępne we wszystkich plikach

### Metoda 2 — Import modułów VBA

1. Pobierz pliki src/PL_TEKST.bas i src/PL_TEKST_FAKTURA.bas
2. Otwórz Excel i naciśnij ALT + F11
3. Prawy klik na projekt → Import File
4. Zaimportuj najpierw PL_TEKST.bas, potem PL_TEKST_FAKTURA.bas
5. Zapisz plik jako .xlsm

### Metoda 3 — Personal Macro Workbook (funkcje globalne)

1. Otwórz Excel i naciśnij ALT + F11
2. W lewym panelu znajdź PERSONAL.XLSB → prawy klik → Insert → Module
3. Wklej zawartość PL_TEKST.bas, utwórz drugi moduł i wklej PL_TEKST_FAKTURA.bas
4. Zapisz (CTRL + S)

---

## Kody błędów

| Komunikat | Przyczyna |
|---|---|
| Blad: nieprawidlowa wartosc | Komórka zawiera tekst lub jest pusta |
| Blad: ujemna liczba | Podano liczbę ujemną |
| Blad: liczba zbyt duza | Przekroczono limit 999 999 999,99 |
| Blad: nieznana waluta 'XXX' | Podano nieobsługiwany kod waluty |

---

## Roadmap

- [x] v1.0.0 — podstawowa funkcja PLN
- [x] v1.0.1 — poprawki odmiany i zaokrąglenia
- [x] v2.0.0 — obsługa walut EUR, USD, GBP + refaktoryzacja
- [x] v2.1.0 — funkcja PL_TEKST_FAKTURA (pełny zapis słowny)
- [x] v3.0.0 — publiczne REST API (Python/FastAPI + Railway)
- [ ] v3.1.0 — prosta aplikacja webowa

---

## Współpraca

Pull requesty są mile widziane! Jeśli chcesz dodać nową walutę, poprawić odmianę lub dodać testy — śmiało.

1. Zrób Fork repozytorium
2. Stwórz branch: git checkout -b feature/nowa-waluta
3. Zatwierdź zmiany: git commit -m 'Dodaj obsługę CHF'
4. Wypchnij: git push origin feature/nowa-waluta
5. Otwórz Pull Request

---

## Licencja

MIT — używaj swobodnie, także komercyjnie.

---

## English summary

**PL-TEKST** is a free project that converts monetary amounts to Polish words, including correct grammatical inflection for PLN, EUR, USD and GBP currencies.

Available as:
- Excel VBA add-in (.xlam) with two functions: PL_TEKST() and PL_TEKST_FAKTURA()
- Public REST API: https://pl-tekst-production.up.railway.app
