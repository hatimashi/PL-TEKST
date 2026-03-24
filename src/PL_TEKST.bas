Option Explicit

' =============================================================================
'  Modul:  PL_TEKST
'  Opis:   Zamienia liczbe na zapis slowny w jezyku polskim (PLN)
'  Uzycie: =PL_TEKST(A1)
'  UWAGA:  Polskie znaki budowane sa przez Chr() - brak problemow z kodowaniem
' =============================================================================

Private Function Jednosci(i As Integer, zenski As Boolean) As String
    Select Case i
        Case 1: If zenski Then Jednosci = "jedna" Else Jednosci = "jeden"
        Case 2: If zenski Then Jednosci = "dwie" Else Jednosci = "dwa"
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

Private Function Nastki(i As Integer) As String
    Select Case i
        Case 0: Nastki = "dziesi" & ChrW(281) & ChrW(263)
        Case 1: Nastki = "jedena" & ChrW(347) & "cie"
        Case 2: Nastki = "dwana" & ChrW(347) & "cie"
        Case 3: Nastki = "trzyna" & ChrW(347) & "cie"
        Case 4: Nastki = "czterna" & ChrW(347) & "cie"
        Case 5: Nastki = "pi" & ChrW(281) & "tna" & ChrW(347) & "cie"
        Case 6: Nastki = "szesna" & ChrW(347) & "cie"
        Case 7: Nastki = "siedemna" & ChrW(347) & "cie"
        Case 8: Nastki = "osiemna" & ChrW(347) & "cie"
        Case 9: Nastki = "dziewi" & ChrW(281) & "tna" & ChrW(347) & "cie"
        Case Else: Nastki = ""
    End Select
End Function

Private Function Dziesiatki(i As Integer) As String
    Select Case i
        Case 2: Dziesiatki = "dwadzie" & ChrW(347) & "cia"
        Case 3: Dziesiatki = "trzydzie" & ChrW(347) & "ci"
        Case 4: Dziesiatki = "czterdzie" & ChrW(347) & "ci"
        Case 5: Dziesiatki = "pi" & ChrW(281) & ChrW(263) & "dziesi" & ChrW(261) & "t"
        Case 6: Dziesiatki = "sze" & ChrW(347) & ChrW(263) & "dziesi" & ChrW(261) & "t"
        Case 7: Dziesiatki = "siedemdziesi" & ChrW(261) & "t"
        Case 8: Dziesiatki = "osiemdziesi" & ChrW(261) & "t"
        Case 9: Dziesiatki = "dziewi" & ChrW(281) & ChrW(263) & "dziesi" & ChrW(261) & "t"
        Case Else: Dziesiatki = ""
    End Select
End Function

Private Function Setki(i As Integer) As String
    Select Case i
        Case 1: Setki = "sto"
        Case 2: Setki = "dwie" & ChrW(347) & "cie"
        Case 3: Setki = "trzysta"
        Case 4: Setki = "czterysta"
        Case 5: Setki = "pi" & ChrW(281) & ChrW(263) & "set"
        Case 6: Setki = "sze" & ChrW(347) & ChrW(263) & "set"
        Case 7: Setki = "siedemset"
        Case 8: Setki = "osiemset"
        Case 9: Setki = "dziewi" & ChrW(281) & ChrW(263) & "set"
        Case Else: Setki = ""
    End Select
End Function

Private Function TrojkaSlownie(n As Integer, zenski As Boolean) As String
    Dim s As Integer, d As Integer, j As Integer
    Dim wynik As String
    wynik = ""
    s = n \ 100
    d = (n Mod 100) \ 10
    j = n Mod 10
    If s > 0 Then wynik = Setki(s)
    If d = 1 Then
        If Len(wynik) > 0 Then wynik = wynik & " "
        wynik = wynik & Nastki(j)
    Else
        If d > 0 Then
            If Len(wynik) > 0 Then wynik = wynik & " "
            wynik = wynik & Dziesiatki(d)
        End If
        If j > 0 Then
            If Len(wynik) > 0 Then wynik = wynik & " "
            wynik = wynik & Jednosci(j, zenski)
        End If
    End If
    TrojkaSlownie = wynik
End Function

Private Function Odmiana(n As Long, f0 As String, f1 As String, f2 As String) As String
    Dim j As Long: j = n Mod 10
    Dim d As Long: d = (n Mod 100) \ 10
    If n = 1 Then
        Odmiana = f0
    ElseIf d = 1 Then
        Odmiana = f2
    ElseIf j >= 2 And j <= 4 Then
        Odmiana = f1
    Else
        Odmiana = f2
    End If
End Function

Function PL_TEKST(kwota As Variant) As String
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

    If liczba >= 1000000000 Then
        PL_TEKST = "Blad: liczba zbyt duza (max 999 999 999,99)"
        Exit Function
    End If

    Dim calkowita As Long
    Dim grosze As Integer
    calkowita = CLng(Int(liczba))
    grosze = CInt(Round((liczba - Int(liczba)) * 100, 0))
    If grosze >= 100 Then calkowita = calkowita + 1: grosze = 0

    Dim mln As Integer
    Dim tys As Integer
    Dim res As Integer
    mln = CInt(calkowita \ 1000000)
    tys = CInt((calkowita Mod 1000000) \ 1000)
    res = CInt(calkowita Mod 1000)

    Dim wynik As String
    wynik = ""

    If mln > 0 Then
        wynik = TrojkaSlownie(mln, False) & " " & _
                Odmiana(CLng(mln), "milion", "miliony", "milion" & ChrW(243) & "w")
    End If

    If tys > 0 Then
        If Len(wynik) > 0 Then wynik = wynik & " "
        Dim tysSlowo As String
        If tys = 1 Then
            tysSlowo = "tysi" & ChrW(261) & "c"
        ElseIf (tys Mod 100) >= 12 And (tys Mod 100) <= 19 Then
            tysSlowo = TrojkaSlownie(tys, False) & " tysi" & ChrW(281) & "cy"
        ElseIf (tys Mod 10) >= 2 And (tys Mod 10) <= 4 Then
            tysSlowo = TrojkaSlownie(tys, True) & " tysi" & ChrW(261) & "ce"
        Else
            tysSlowo = TrojkaSlownie(tys, False) & " tysi" & ChrW(281) & "cy"
        End If
        wynik = wynik & tysSlowo
    End If

    If res > 0 Then
        If Len(wynik) > 0 Then wynik = wynik & " "
        wynik = wynik & TrojkaSlownie(res, True)
    End If

    If calkowita = 0 Then wynik = "zero"

    Dim nazwaZloty As String
    nazwaZloty = Odmiana(calkowita, _
        "z" & ChrW(322) & "oty", _
        "z" & ChrW(322) & "ote", _
        "z" & ChrW(322) & "otych")

    Dim groszeTxt As String
    groszeTxt = Format(grosze, "00") & "/100 " & _
                Odmiana(CLng(grosze), "grosz", "grosze", "groszy")

    Dim rezultat As String
    rezultat = wynik & " " & nazwaZloty & " (" & groszeTxt & ")"
    If Len(rezultat) > 0 Then rezultat = UCase(Left(rezultat, 1)) & Mid(rezultat, 2)
    PL_TEKST = rezultat
    Exit Function
Blad:
    PL_TEKST = "Blad: " & Err.Description
End Function
