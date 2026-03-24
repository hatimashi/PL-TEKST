# Instalacja PL-TEKST

## Wymagania

- Microsoft Excel 2010 lub nowszy (Windows)
- Włączona obsługa makr w Excelu

## Włączenie obsługi makr

Jeśli Excel blokuje makra:

1. **Plik → Opcje → Centrum zaufania → Ustawienia Centrum zaufania**
2. Wybierz **Ustawienia makr**
3. Zaznacz **Włącz wszystkie makra** (lub "Włącz makra z powiadomieniem")
4. Kliknij **OK**

---

## Metoda 1 — Dodatek .xlam (zalecana)

Najwygodniejsza metoda — funkcja dostępna we wszystkich plikach automatycznie.

1. Pobierz `PL_TEKST.xlam` z sekcji [Releases](../../releases)
2. Zapisz go w dowolnym miejscu na dysku (np. `C:\Users\TwojaNazwa\Documents\Dodatki Excel\`)
3. W Excelu: **Plik → Opcje → Dodatki**
4. Na dole wybierz **Zarządzaj: Dodatki programu Excel** → kliknij **Przejdź**
5. Kliknij **Przeglądaj** i wskaż plik `PL_TEKST.xlam`
6. Upewnij się że checkbox przy **PL_TEKST** jest zaznaczony → **OK**

Funkcja `=PL_TEKST()` jest teraz dostępna we wszystkich otwartych i przyszłych plikach.

---

## Metoda 2 — Import modułu VBA do konkretnego pliku

Jeśli chcesz mieć funkcję tylko w jednym pliku.

1. Pobierz `src/PL_TEKST.bas`
2. Otwórz swój plik Excel
3. Naciśnij `ALT + F11` (otwiera edytor VBA)
4. W lewym panelu (Project Explorer) znajdź swój plik
5. Kliknij prawym przyciskiem → **Import File**
6. Wskaż plik `PL_TEKST.bas` → **Otwórz**
7. Zapisz plik jako **Skoroszyt z obsługą makr (.xlsm)**

---

## Metoda 3 — Personal Macro Workbook

Funkcja globalna bez instalowania dodatku — działa na Twoim komputerze we wszystkich plikach.

1. Naciśnij `ALT + F11`
2. W lewym panelu znajdź **PERSONAL.XLSB**
   - Jeśli go nie ma: nagraj dowolne makro przez **Widok → Makra → Nagraj makro**, wybierz "Skoroszyt makr osobistych", zatrzymaj — plik się utworzy
3. Rozwiń **PERSONAL.XLSB** → prawy klik na **Modules** → **Insert → Module**
4. Skopiuj i wklej zawartość pliku `src/PL_TEKST.bas`
5. Zapisz (`CTRL + S`)

---

## Weryfikacja instalacji

W dowolnej komórce Excela wpisz:

```
=PL_TEKST(1234.67)
```

Oczekiwany wynik:
```
Jeden tysiąc dwieście trzydzieści cztery złote (67/100 groszy)
```

---

## Odinstalowanie

**Dodatek .xlam:**
Plik → Opcje → Dodatki → Przejdź → odznacz checkbox przy PL_TEKST → OK

**Moduł VBA:**
ALT+F11 → znajdź moduł PL_TEKST → prawy klik → Remove Module
