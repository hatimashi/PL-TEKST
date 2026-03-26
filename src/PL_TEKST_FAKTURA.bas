Option Explicit

' =============================================================================
'  Modul:  PL_TEKST_FAKTURA  v1.0
'  Autor:  hatimashi
'  Opis:   Zamienia liczbe na pelny zapis slowny dla dokumentow prawnych
'  Uzycie: =PL_TEKST_FAKTURA(kwota)
'          =PL_TEKST_FAKTURA(kwota, "EUR")
'          =PL_TEKST_FAKTURA(kwota, "USD")
'          =PL_TEKST_FAKTURA(kwota, "GBP")
'
'  Roznica vs PL_TEKST:
'    PL_TEKST         -> Tysiąc złotych (67/100 groszy)
'    PL_TEKST_FAKTURA -> tysiąc złotych i sześćdziesiąt siedem groszy
' =============================================================================


' =============================================================================
'  UWAGA: Ten modul wymaga rowniez modulu PL_TEKST (wspoldzielone funkcje)
'  Upewnij sie ze oba moduly sa zaimportowane do projektu VBA
' =============================================================================


' =============================================================================
'  SEKCJA 1: SLOWNIK WALUT DLA FAKTURY
'  Format: pelna odmiana dla czesci glownej i podrzednej
' =============================================================================

Private Function PobierzWaluteFaktura(kod As String, _
                                      ByRef glowna() As String, _
                                      ByRef podrzedna() As String) As Boolean
    ReDim glowna(2)
    ReDim podrzedna(2)

    Select Case UCase(kod)

        Case "PLN"
            glowna(0) = "z" & ChrW(322) & "oty"
            glowna(1) = "z" & ChrW(322) & "ote"
            glowna(2) = "z" & ChrW(322) & "otych"
            podrzedna(0) = "grosz"
            podrzedna(1) = "grosze"
            podrzedna(2) = "groszy"

        Case "EUR"
            glowna(0) = "euro"
            glowna(1) = "euro"
            glowna(2) = "euro"
            podrzedna(0) = "cent"
            podrzedna(1) = "centy"
            podrzedna(2) = "cent" & ChrW(243) & "w"

        Case "USD"
            glowna(0) = "dolar"
            glowna(1) = "dolary"
            glowna(2) = "dolar" & ChrW(243) & "w"
            podrzedna(0) = "cent"
            podrzedna(1) = "centy"
            podrzedna(2) = "cent" & ChrW(243) & "w"

        Case "GBP"
            glowna(0) = "funt"
            glowna(1) = "funty"
            glowna(2) = "funt" & ChrW(243) & "w"
            podrzedna(0) = "pens"
            podrzedna(1) = "pensy"
            podrzedna(2) = "pens" & ChrW(243) & "w"

        Case Else
            PobierzWaluteFaktura = False
            Exit Function

    End Select

    PobierzWaluteFaktura = True
End Function


' =============================================================================
'  SEKCJA 2: ODMIANA (lokalna kopia zeby modul byl niezalezny)
' =============================================================================

Private Function OdmianaF(n As Long, f1 As String, f2_4 As String, f5 As String) As String
    Dim j As Long: j = n Mod 10
    Dim d As Long: d = (n Mod 100) \ 10
    If n = 1 Then
        OdmianaF = f1
    ElseIf n = 0 Or d = 1 Then
        OdmianaF = f5
    ElseIf j >= 2 And j <= 4 Then
        OdmianaF = f2_4
    Else
        OdmianaF = f5
    End If
End Function


' =============================================================================
'  SEKCJA 3: GLOWNA FUNKCJA PUBLICZNA
' =============================================================================

Function PL_TEKST_FAKTURA(kwota As Variant, Optional waluta As String = "PLN") As String
    On Error GoTo Blad

    ' --- Walidacja ---
    If IsEmpty(kwota) Or Not IsNumeric(kwota) Then
        PL_TEKST_FAKTURA = "Blad: nieprawidlowa wartosc"
        Exit Function
    End If

    Dim liczba As Double
    liczba = CDbl(kwota)

    If liczba < 0 Then
        PL_TEKST_FAKTURA = "Blad: ujemna liczba"
        Exit Function
    End If

    If liczba >= 1000000000# Then
        PL_TEKST_FAKTURA = "Blad: liczba zbyt duza (max 999 999 999,99)"
        Exit Function
    End If

    ' --- Pobierz walute ---
    Dim glowna() As String
    Dim podrzedna() As String
    If Not PobierzWaluteFaktura(waluta, glowna, podrzedna) Then
        PL_TEKST_FAKTURA = "Blad: nieznana waluta '" & waluta & "' (dostepne: PLN, EUR, USD, GBP)"
        Exit Function
    End If

    ' --- Rozdziel na czesc calkowita i podrzedna ---
    Dim liczbaCur As Currency
    liczbaCur = CCur(liczba)

    Dim calkowita As Long
    Dim grosze As Integer
    calkowita = CLng(Int(liczbaCur))
    grosze = CInt(liczbaCur * 100 - CCur(Int(liczbaCur)) * 100)

    If grosze >= 100 Then
        calkowita = calkowita + 1
        grosze = 0
    End If

    ' --- Zamien na slowa ---
    Dim tekstCalkowita As String
    Dim nazwaCalkowita As String
    Dim tekstGrosze As String
    Dim nazwaGrosze As String

    tekstCalkowita = LiczbaSlownie(calkowita)
    nazwaCalkowita = OdmianaF(calkowita, glowna(0), glowna(1), glowna(2))

    tekstGrosze = LiczbaSlownie(CLng(grosze))
    nazwaGrosze = OdmianaF(CLng(grosze), podrzedna(0), podrzedna(1), podrzedna(2))

    ' --- Zloz wynik ---
    Dim rezultat As String
    rezultat = tekstCalkowita & " " & nazwaCalkowita & " i " & tekstGrosze & " " & nazwaGrosze

    ' Pierwsza litera wielka
    rezultat = UCase(Left(rezultat, 1)) & Mid(rezultat, 2)

    PL_TEKST_FAKTURA = rezultat
    Exit Function
Blad:
    PL_TEKST_FAKTURA = "Blad: " & Err.Description
End Function
