# =============================================================================
#  Modul:  main.py
#  Opis:   API dla projektu PL-TEKST
#  Autor:  hatimashi
#
#  Endpointy:
#    GET  /                              - info o API
#    POST /pl-tekst                      - zapis z ulamkiem (67/100 groszy)
#    POST /pl-tekst-faktura              - pelny zapis slowny
#    GET  /pl-tekst?kwota=&waluta=       - zapis z ulamkiem (GET)
#    GET  /pl-tekst-faktura?kwota=&waluta= - pelny zapis slowny (GET)
# =============================================================================

from fastapi import FastAPI, Query, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from pl_tekst import pl_tekst, pl_tekst_faktura

# --- Aplikacja ---
app = FastAPI(
    title="PL-TEKST API",
    description="Zamiana kwot pieniężnych na zapis słowny w języku polskim.",
    version="1.0.0",
)

# --- CORS (umozliwia wywolanie API z przegladarki / aplikacji webowej) ---
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# --- Modele danych ---
class ZapytanieKwota(BaseModel):
    kwota: float = Field(..., example=1234.67, description="Kwota do zamiany (max 999 999 999,99)")
    waluta: str  = Field("PLN", example="PLN", description="Kod waluty: PLN, EUR, USD, GBP")


class OdpowiedzTekst(BaseModel):
    kwota:  float
    waluta: str
    wynik:  str


# --- Endpointy ---

@app.get("/", tags=["Info"])
def root():
    return {
        "nazwa": "PL-TEKST API",
        "wersja": "1.0.0",
        "opis": "Zamiana kwot pieniężnych na zapis słowny w języku polskim",
        "endpointy": {
            "POST /pl-tekst":           "Zapis z ułamkiem groszy (67/100 groszy)",
            "POST /pl-tekst-faktura":   "Pełny zapis słowny dla dokumentów prawnych",
            "GET  /pl-tekst":           "Zapis z ułamkiem (parametry w URL)",
            "GET  /pl-tekst-faktura":   "Pełny zapis słowny (parametry w URL)",
        },
        "waluty": ["PLN", "EUR", "USD", "GBP"],
        "dokumentacja": "/docs",
    }


@app.post("/pl-tekst", response_model=OdpowiedzTekst, tags=["PL_TEKST"])
def endpoint_pl_tekst_post(dane: ZapytanieKwota):
    """
    Zamienia kwotę na zapis z ułamkiem groszy.

    Przykład wyniku: **Tysiąc dwieście trzydzieści cztery złote (67/100 groszy)**
    """
    try:
        wynik = pl_tekst(dane.kwota, dane.waluta)
        return OdpowiedzTekst(kwota=dane.kwota, waluta=dane.waluta.upper(), wynik=wynik)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/pl-tekst-faktura", response_model=OdpowiedzTekst, tags=["PL_TEKST_FAKTURA"])
def endpoint_pl_tekst_faktura_post(dane: ZapytanieKwota):
    """
    Zamienia kwotę na pełny zapis słowny dla dokumentów prawnych.

    Przykład wyniku: **Tysiąc dwieście trzydzieści cztery złote i sześćdziesiąt siedem groszy**
    """
    try:
        wynik = pl_tekst_faktura(dane.kwota, dane.waluta)
        return OdpowiedzTekst(kwota=dane.kwota, waluta=dane.waluta.upper(), wynik=wynik)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/pl-tekst", response_model=OdpowiedzTekst, tags=["PL_TEKST"])
def endpoint_pl_tekst_get(
    kwota:  float = Query(...,   example=1234.67, description="Kwota do zamiany"),
    waluta: str   = Query("PLN", example="PLN",   description="Kod waluty: PLN, EUR, USD, GBP"),
):
    """
    Zamienia kwotę na zapis z ułamkiem groszy (wersja GET).

    Przykład: `/pl-tekst?kwota=1234.67&waluta=PLN`
    """
    try:
        wynik = pl_tekst(kwota, waluta)
        return OdpowiedzTekst(kwota=kwota, waluta=waluta.upper(), wynik=wynik)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/pl-tekst-faktura", response_model=OdpowiedzTekst, tags=["PL_TEKST_FAKTURA"])
def endpoint_pl_tekst_faktura_get(
    kwota:  float = Query(...,   example=1234.67, description="Kwota do zamiany"),
    waluta: str   = Query("PLN", example="PLN",   description="Kod waluty: PLN, EUR, USD, GBP"),
):
    """
    Zamienia kwotę na pełny zapis słowny dla dokumentów prawnych (wersja GET).

    Przykład: `/pl-tekst-faktura?kwota=1234.67&waluta=PLN`
    """
    try:
        wynik = pl_tekst_faktura(kwota, waluta)
        return OdpowiedzTekst(kwota=kwota, waluta=waluta.upper(), wynik=wynik)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
