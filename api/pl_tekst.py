# =============================================================================
#  Modul:  pl_tekst.py
#  Opis:   Zamienia kwoty pieniezne na zapis slowny w jezyku polskim
#  Funkcje publiczne:
#    pl_tekst(kwota, waluta)         - zapis z ulamkiem (67/100 groszy)
#    pl_tekst_faktura(kwota, waluta) - pelny zapis slowny
# =============================================================================

# -----------------------------------------------------------------------------
#  SEKCJA 1: SLOWNIK LICZEBNIKOW
# -----------------------------------------------------------------------------

JEDNOSCI_M = ["", "jeden", "dwa", "trzy", "cztery", "pięć", "sześć",
              "siedem", "osiem", "dziewięć"]

JEDNOSCI_F = ["", "jedna", "dwie", "trzy", "cztery", "pięć", "sześć",
              "siedem", "osiem", "dziewięć"]

NASTKI = ["dziesięć", "jedenaście", "dwanaście", "trzynaście", "czternaście",
          "piętnaście", "szesnaście", "siedemnaście", "osiemnaście", "dziewiętnaście"]

DZIESIATKI = ["", "dziesięć", "dwadzieścia", "trzydzieści", "czterdzieści",
              "pięćdziesiąt", "sześćdziesiąt", "siedemdziesiąt", "osiemdziesiąt",
              "dziewięćdziesiąt"]

SETKI = ["", "sto", "dwieście", "trzysta", "czterysta", "pięćset",
         "sześćset", "siedemset", "osiemset", "dziewięćset"]


# -----------------------------------------------------------------------------
#  SEKCJA 2: ODMIANA
# -----------------------------------------------------------------------------

def odmiana(n: int, f1: str, f2_4: str, f5: str) -> str:
    """Zwraca poprawną formę odmiany rzeczownika po liczbie."""
    j = n % 10
    d = (n % 100) // 10
    if n == 1:
        return f1
    elif n == 0 or d == 1:
        return f5
    elif 2 <= j <= 4:
        return f2_4
    else:
        return f5


def odmiana_tysiac(n: int) -> str:
    j = n % 10
    d = (n % 100) // 10
    if n == 1:
        return "tysiąc"
    elif d == 1:
        return "tysięcy"
    elif 2 <= j <= 4:
        return "tysiące"
    else:
        return "tysięcy"


def odmiana_milion(n: int) -> str:
    j = n % 10
    d = (n % 100) // 10
    if n == 1:
        return "milion"
    elif d == 1:
        return "milionów"
    elif 2 <= j <= 4:
        return "miliony"
    else:
        return "milionów"


# -----------------------------------------------------------------------------
#  SEKCJA 3: ZAMIANA LICZBY NA SLOWA
# -----------------------------------------------------------------------------

def trojka_slownie(n: int, zenski: bool = False) -> str:
    """Zamienia liczbę 0-999 na słowa."""
    if n == 0:
        return ""

    jednosci = JEDNOSCI_F if zenski else JEDNOSCI_M
    czesci = []

    s = n // 100
    reszta = n % 100

    if s > 0:
        czesci.append(SETKI[s])

    if 10 <= reszta <= 19:
        czesci.append(NASTKI[reszta - 10])
    else:
        d = reszta // 10
        j = reszta % 10
        if d > 0:
            czesci.append(DZIESIATKI[d])
        if j > 0:
            czesci.append(jednosci[j])

    return " ".join(czesci)


def liczba_slownie(n: int) -> str:
    """Zamienia dowolną liczbę całkowitą (0-999999999) na słowa."""
    if n == 0:
        return "zero"

    mln = n // 1_000_000
    tys = (n % 1_000_000) // 1_000
    res = n % 1_000

    czesci = []

    if mln > 0:
        czesci.append(f"{trojka_slownie(mln)} {odmiana_milion(mln)}")

    if tys > 0:
        if tys == 1:
            czesci.append(odmiana_tysiac(1))
        else:
            czesci.append(f"{trojka_slownie(tys)} {odmiana_tysiac(tys)}")

    if res > 0:
        czesci.append(trojka_slownie(res))

    return " ".join(czesci)


# -----------------------------------------------------------------------------
#  SEKCJA 4: SLOWNIK WALUT
# -----------------------------------------------------------------------------

WALUTY = {
    "PLN": {
        "glowna":    ("złoty",   "złote",   "złotych"),
        "podrzedna": ("grosz",   "grosze",  "groszy"),
    },
    "EUR": {
        "glowna":    ("euro",    "euro",    "euro"),
        "podrzedna": ("cent",    "centy",   "centów"),
    },
    "USD": {
        "glowna":    ("dolar",   "dolary",  "dolarów"),
        "podrzedna": ("cent",    "centy",   "centów"),
    },
    "GBP": {
        "glowna":    ("funt",    "funty",   "funtów"),
        "podrzedna": ("pens",    "pensy",   "pensów"),
    },
}


def pobierz_walute(kod: str) -> dict | None:
    return WALUTY.get(kod.upper())


# -----------------------------------------------------------------------------
#  SEKCJA 5: FUNKCJE PUBLICZNE
# -----------------------------------------------------------------------------

def _podziel_kwote(kwota: float) -> tuple[int, int]:
    """Rozdziela kwotę na część całkowitą i grosze. Unika błędów float."""
    zaokraglona = round(kwota, 2)
    calkowita = int(zaokraglona)
    grosze = round((zaokraglona - calkowita) * 100)
    if grosze >= 100:
        calkowita += 1
        grosze = 0
    return calkowita, int(grosze)


def pl_tekst(kwota: float, waluta: str = "PLN") -> str:
    """
    Zamienia kwotę na zapis z ułamkiem groszy (XX/100).

    Przykład:
        pl_tekst(1234.67) -> "Tysiąc dwieście trzydzieści cztery złote (67/100 groszy)"
        pl_tekst(1234.67, "EUR") -> "Tysiąc dwieście trzydzieści cztery euro (67/100 centów)"
    """
    w = pobierz_walute(waluta)
    if w is None:
        raise ValueError(f"Nieznana waluta '{waluta}'. Dostępne: {', '.join(WALUTY.keys())}")

    calkowita, grosze = _podziel_kwote(kwota)

    tekst = liczba_slownie(calkowita)
    nazwa = odmiana(calkowita, *w["glowna"])
    nazwa_groszy = w["podrzedna"][2]  # zawsze forma dopełniaczowa

    wynik = f"{tekst} {nazwa} ({grosze:02d}/100 {nazwa_groszy})"
    return wynik[0].upper() + wynik[1:]


def pl_tekst_faktura(kwota: float, waluta: str = "PLN") -> str:
    """
    Zamienia kwotę na pełny zapis słowny dla dokumentów prawnych.

    Przykład:
        pl_tekst_faktura(1234.67) -> "Tysiąc dwieście trzydzieści cztery złote i sześćdziesiąt siedem groszy"
        pl_tekst_faktura(1234.67, "EUR") -> "Tysiąc dwieście trzydzieści cztery euro i sześćdziesiąt siedem centów"
    """
    w = pobierz_walute(waluta)
    if w is None:
        raise ValueError(f"Nieznana waluta '{waluta}'. Dostępne: {', '.join(WALUTY.keys())}")

    calkowita, grosze = _podziel_kwote(kwota)

    tekst_calkowita = liczba_slownie(calkowita)
    nazwa_calkowita = odmiana(calkowita, *w["glowna"])

    tekst_grosze = liczba_slownie(grosze)
    nazwa_grosze = odmiana(grosze, *w["podrzedna"])

    wynik = f"{tekst_calkowita} {nazwa_calkowita} i {tekst_grosze} {nazwa_grosze}"
    return wynik[0].upper() + wynik[1:]
