Option Explicit

' =============================================================================
'  Modul:  PL_TEKST  v2.0
'  Autor:  hatimashi
'  Opis:   Zamienia liczbe na zapis slowny w jezyku polskim
'  Uzycie: =PL_TEKST(kwota)
'          =PL_TEKST(kwota, "EUR")
'          =PL_TEKST(kwota, "USD")
'          =PL_TEKST(kwota, "GBP")
' =============================================================================


' =============================================================================
'  SEKCJA 1: SLOWNIK LICZEBNIKOW
' =============================================================================

Private Function Jednosci(n As Integer, zenski As Boolean) As String
    Select Case n
        Case 1: If zenski Then Jednosci = "jedna" Else Jednosci = "jeden"
        Case 2: If zenski Then Jednosci = "dwie"  Else Jednosci = "dwa"
        Case 3: Jednosci = "trzy"
        Case 4: Jednosci = "cztery"
        Case 5: Jednosci = "pi" & ChrW(281) & ChrW(263)
        Case 6: Jednosci = "sze" & ChrW(347) & ChrW(263)
        Case 7: Jednosci = "siedem"
        Case 8: Jednosci = "osiem"
        Case 9: Jednosci = "dziewi" & ChrW(281) & ChrW(263)
        Case Else: Jednosci = ""
    End Select
End Function

Private Function Nastki(n As Integer) As String
    Select Case n
        Case 10: Nastki = "dziesi" & ChrW(281) & ChrW(263)
        Case 11: Nastki = "jedena" & ChrW(347) & "cie"
        Case 12: Nastki = "dwana" & ChrW(347) & "cie"
        Case 13: Nastki = "trzyna" & ChrW(347) & "cie"
        Case 14: Nastki = "czterna" & ChrW(347) & "cie"
        Case 15: Nastki = "pi" & ChrW(281) & "tna" & ChrW(347) & "cie"
        Case 16: Nastki = "szesna" & ChrW(347) & "cie"
        Case 17: Nastki = "siedemna" & ChrW(347) & "cie"
        Case 18: Nastki = "osiemna" & ChrW(347) & "cie"
        Case 19: Nastki = "dziewi" & ChrW(281) & "tna" & ChrW(347) & "cie"
        Case Else: Nastki = ""
    End Select
End Function

Private Function Dziesiatki(n As Integer) As String
    Select Case n
        Case 20: Dziesiatki = "dwadzie" & ChrW(347) & "cia"
        Case 30: Dziesiatki = "trzydzie" & ChrW(347) & "ci"
        Case 40: Dziesiatki = "czterdzie" & ChrW(347) & "ci"
        Case 50: Dziesiatki = "pi" & ChrW(281) & ChrW(263) & "dziesi" & ChrW(261) & "t"
        Case 60: Dziesiatki = "sze" & ChrW(347) & ChrW(263) & "dziesi" & ChrW(261) & "t"
        Case 70: Dziesiatki = "siedemdziesi" & ChrW(261) & "t"
        Case 80: Dziesiatki = "osiemdziesi" & ChrW(261) & "t"
        Case 90: Dziesiatki = "dziewi" & ChrW(281) & ChrW(263) & "dziesi" & ChrW(261) & "t"
        Case Else: Dziesiatki = ""
    End Select
End Function

Private Function Setki(n As Integer) As String
    Select Case n
        Case 100: Setki = "sto"
        Case 200: Setki = "dwie" & ChrW(347) & "cie"
        Case 300: Setki = "trzysta"
        Case 400: Setki = "czterysta"
        Case 500: Setki = "pi" & ChrW(281) & ChrW(263) & "set"
        Case 600: Setki = "sze" & ChrW(347) & ChrW(263) & "set"
        Case 700: Setki = "siedemset"
        Case 800: Setki = "osiemset"
        Case 900: Setki = "dziewi" & ChrW(281) & ChrW(263) & "set"
        Case Else: Setki = ""
    End Select
End Function


' =============================================================================
'  SEKCJA 2: LOGIKA ZAMIANY LICZBY NA SLOWA
' =============================================================================

Private Function TrojkaSlownie(n As Integer, zenski As Boolean) As String
    Dim s As Integer, reszta As Integer, dz As Integer, j As Integer
    Dim wynik As String
    wynik = ""

    s = (n \ 100) * 100
    reszta = n Mod 100

    If s > 0 Then wynik = Setki(s)

    If reszta >= 10 And reszta <= 19 Then
        If Len(wynik) > 0 Then wynik = wynik & " "
        wynik = wynik & Nastki(reszta)
    Else
        dz = (reszta \ 10) * 10
        j = reszta Mod 10
        If dz > 0 Then
            If Len(wynik) > 0 Then wynik = wynik & " "
            wynik = wynik & Dziesiatki(dz)
        End If
        If j > 0 Then
            If Len(wynik) > 0 Then wynik = wynik & " "
            wynik = wynik & Jednosci(j, zenski)
        End If
    End If

    TrojkaSlownie = wynik
End Function

Private Function LiczbaSlownie(n As Long) As String
    Dim mln As Integer, tys As Integer, res As Integer
    Dim wynik As String
    wynik = ""

    mln = CInt(n \ 1000000)
    tys = CInt((n Mod 1000000) \ 1000)
    res = CInt(n Mod 1000)

    If mln > 0 Then
        wynik = TrojkaSlownie(mln, False) & " " & OdmianaMilion(mln)
    End If

    If tys > 0 Then
        If Len(wynik) > 0 Then wynik = wynik & " "
        If tys = 1 Then
            wynik = wynik & OdmianaTysiac(1)
        Else
            wynik = wynik & TrojkaSlownie(tys, False) & " " & OdmianaTysiac(tys)
        End If
    End If

    If res > 0 Then
        If Len(wynik) > 0 Then wynik = wynik & " "
        wynik = wynik & TrojkaSlownie(res, False)
    End If

    If n = 0 Then wynik = "zero"

    LiczbaSlownie = wynik
End Function


' =============================================================================
'  SEKCJA 3: ODMIANA
' =============================================================================

Private Function Odmiana(n As Long, f1 As String, f2_4 As String, f5 As String) As String
    Dim j As Long: j = n Mod 10
    Dim d As Long: d = (n Mod 100) \ 10
    If n = 1 Then
        Odmiana = f1
    ElseIf n = 0 Or d = 1 Then
        Odmiana = f5
    ElseIf j >= 2 And j <= 4 Then
        Odmiana = f2_4
    Else
        Odmiana = f5
    End If
End Function

Private Function OdmianaTysiac(n As Integer) As String
    Dim j As Integer: j = n Mod 10
    Dim d As Integer: d = (n Mod 100) \ 10
    If n = 1 Then
        OdmianaTysiac = "tysi" & ChrW(261) & "c"
    ElseIf d = 1 Then
        OdmianaTysiac = "tysi" & ChrW(281) & "cy"
    ElseIf j >= 2 And j <= 4 Then
        OdmianaTysiac = "tysi" & ChrW(261) & "ce"
    Else
        OdmianaTysiac = "tysi" & ChrW(281) & "cy"
    End If
End Function

Private Function OdmianaMilion(n As Integer) As String
    Dim j As Integer: j = n Mod 10
    Dim d As Integer: d = (n Mod 100) \ 10
    If n = 1 Then
        OdmianaMilion = "milion"
    ElseIf d = 1 Then
        OdmianaMilion = "milion" & ChrW(243) & "w"
    ElseIf j >= 2 And j <= 4 Then
        OdmianaMilion = "miliony"
    Else
        OdmianaMilion = "milion" & ChrW(243) & "w"
    End If
End Function


' =============================================================================
'  SEKCJA 4: SLOWNIK WALUT
'  Aby dodac nowa walute — wystarczy dodac Case w funkcji PobierzWalute
'
'  Format tablicy nazw: nazwy(0)=f1, nazwy(1)=f2_4, nazwy(2)=f5
'  Przyklad PLN: nazwy(0)="zloty", nazwy(1)="zlote", nazwy(2)="zlotych"
' =============================================================================

' Wypelnia tablice 3 form odmiany dla waluty glownej i podrzednej
' Zwraca False jesli waluta nieznana
Private Function PobierzWalute(kod As String, _
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
            PobierzWalute = False
            Exit Function

    End Select

    PobierzWalute = True
End Function


' =============================================================================
'  SEKCJA 5: GLOWNA FUNKCJA PUBLICZNA
' =============================================================================

Function PL_TEKST(kwota As Variant, Optional waluta As String = "PLN") As String
    On Error GoTo Blad

    If IsEmpty(kwota) Or Not IsNumeric(kwota) Then
        PL_TEKST = "Blad: nieprawidlowa wartosc"
        Exit Function
    End If

    Dim liczba As Double
    liczba = CDbl(kwota)

    If liczba < 0 Then
        PL_TEKST = "Blad: ujemna liczba"
        Exit Function
    End If

    If liczba >= 1000000000# Then
        PL_TEKST = "Blad: liczba zbyt duza (max 999 999 999,99)"
        Exit Function
    End If

    ' Pobierz nazwy waluty
    Dim glowna() As String
    Dim podrzedna() As String
    If Not PobierzWalute(waluta, glowna, podrzedna) Then
        PL_TEKST = "Blad: nieznana waluta '" & waluta & "' (dostepne: PLN, EUR, USD, GBP)"
        Exit Function
    End If

    ' Rozdziel na czesc calkowita i grosze
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

    ' Zamien na slowa
    Dim tekstCalkowita As String
    Dim nazwaCalkowita As String
    Dim tekstGrosze As String

    tekstCalkowita = LiczbaSlownie(calkowita)
    nazwaCalkowita = Odmiana(calkowita, glowna(0), glowna(1), glowna(2))
    tekstGrosze = Format(grosze, "00") & "/100 " & podrzedna(2)

    ' Zloz i zwroc wynik
    Dim rezultat As String
    rezultat = tekstCalkowita & " " & nazwaCalkowita & " (" & tekstGrosze & ")"
    rezultat = UCase(Left(rezultat, 1)) & Mid(rezultat, 2)

    PL_TEKST = rezultat
    Exit Function
Blad:
    PL_TEKST = "Blad: " & Err.Description
End Function
