Attribute VB_Name = "Module2"
Option Explicit

' =============================================================================
'  Modul:  PL_TEKST  v1.1
'  Uzycie: =PL_TEKST(A1)
'  Uwaga:  Wszystkie polskie wyrazy sa zwracanymi wartosciami funkcji
'          - zero problemow z kodowaniem przy imporcie
' =============================================================================

' --- Pojedyncze wyrazy jako funkcje (bezpieczne kodowanie) ---

Private Function W_jeden() As String:        W_jeden = "jeden":          End Function
Private Function W_jedna() As String:        W_jedna = "jedna":          End Function
Private Function W_dwa() As String:          W_dwa = "dwa":              End Function
Private Function W_dwie() As String:         W_dwie = "dwie":            End Function
Private Function W_trzy() As String:         W_trzy = "trzy":            End Function
Private Function W_cztery() As String:       W_cztery = "cztery":        End Function
Private Function W_piec() As String:         W_piec = "pi" & ChrW(281) & ChrW(263): End Function
Private Function W_szesc() As String:        W_szesc = "sze" & ChrW(347) & ChrW(263): End Function
Private Function W_siedem() As String:       W_siedem = "siedem":        End Function
Private Function W_osiem() As String:        W_osiem = "osiem":          End Function
Private Function W_dziewiec() As String:     W_dziewiec = "dziewi" & ChrW(281) & ChrW(263): End Function

Private Function W_dziesiec() As String:     W_dziesiec = "dziesi" & ChrW(281) & ChrW(263): End Function
Private Function W_jedenascie() As String:   W_jedenascie = "jedena" & ChrW(347) & "cie": End Function
Private Function W_dwanascie() As String:    W_dwanascie = "dwana" & ChrW(347) & "cie": End Function
Private Function W_trzynascie() As String:   W_trzynascie = "trzyna" & ChrW(347) & "cie": End Function
Private Function W_czternascie() As String:  W_czternascie = "czterna" & ChrW(347) & "cie": End Function
Private Function W_pietnascie() As String:   W_pietnascie = "pi" & ChrW(281) & "tna" & ChrW(347) & "cie": End Function
Private Function W_szesnascie() As String:   W_szesnascie = "szesna" & ChrW(347) & "cie": End Function
Private Function W_siedemnascie() As String: W_siedemnascie = "siedemna" & ChrW(347) & "cie": End Function
Private Function W_osiemnascie() As String:  W_osiemnascie = "osiemna" & ChrW(347) & "cie": End Function
Private Function W_dziewietnascie() As String: W_dziewietnascie = "dziewi" & ChrW(281) & "tna" & ChrW(347) & "cie": End Function

Private Function W_dwadziescia() As String:  W_dwadziescia = "dwadzie" & ChrW(347) & "cia": End Function
Private Function W_trzydziesci() As String:  W_trzydziesci = "trzydzie" & ChrW(347) & "ci": End Function
Private Function W_czterdziesci() As String: W_czterdziesci = "czterdzie" & ChrW(347) & "ci": End Function
Private Function W_piecdziesiat() As String: W_piecdziesiat = W_piec() & "dziesi" & ChrW(261) & "t": End Function
Private Function W_szescdziesiat() As String: W_szescdziesiat = W_szesc() & "dziesi" & ChrW(261) & "t": End Function
Private Function W_siedemdziesiat() As String: W_siedemdziesiat = "siedemdziesi" & ChrW(261) & "t": End Function
Private Function W_osiemdziesiat() As String: W_osiemdziesiat = "osiemdziesi" & ChrW(261) & "t": End Function
Private Function W_dziewiecdziesiat() As String: W_dziewiecdziesiat = W_dziewiec() & "dziesi" & ChrW(261) & "t": End Function

Private Function W_sto() As String:          W_sto = "sto":              End Function
Private Function W_dwiescie() As String:     W_dwiescie = "dwie" & ChrW(347) & "cie": End Function
Private Function W_trzysta() As String:      W_trzysta = "trzysta":      End Function
Private Function W_czterysta() As String:    W_czterysta = "czterysta":  End Function
Private Function W_piecset() As String:      W_piecset = W_piec() & "set": End Function
Private Function W_szescset() As String:     W_szescset = W_szesc() & "set": End Function
Private Function W_siedemset() As String:    W_siedemset = "siedemset":  End Function
Private Function W_osiemset() As String:     W_osiemset = "osiemset":    End Function
Private Function W_dziewiecset() As String:  W_dziewiecset = W_dziewiec() & "set": End Function

' --- Jednosci (1-9) ---
Private Function Jednosci(n As Integer, zenski As Boolean) As String
    Select Case n
        Case 1: If zenski Then Jednosci = W_jedna() Else Jednosci = W_jeden()
        Case 2: If zenski Then Jednosci = W_dwie() Else Jednosci = W_dwa()
        Case 3: Jednosci = W_trzy()
        Case 4: Jednosci = W_cztery()
        Case 5: Jednosci = W_piec()
        Case 6: Jednosci = W_szesc()
        Case 7: Jednosci = W_siedem()
        Case 8: Jednosci = W_osiem()
        Case 9: Jednosci = W_dziewiec()
        Case Else: Jednosci = ""
    End Select
End Function

' --- Nastolatki (10-19) ---
Private Function Nastki(n As Integer) As String
    Select Case n
        Case 10: Nastki = W_dziesiec()
        Case 11: Nastki = W_jedenascie()
        Case 12: Nastki = W_dwanascie()
        Case 13: Nastki = W_trzynascie()
        Case 14: Nastki = W_czternascie()
        Case 15: Nastki = W_pietnascie()
        Case 16: Nastki = W_szesnascie()
        Case 17: Nastki = W_siedemnascie()
        Case 18: Nastki = W_osiemnascie()
        Case 19: Nastki = W_dziewietnascie()
        Case Else: Nastki = ""
    End Select
End Function

' --- Dziesiatki (20-90) ---
Private Function Dziesiatki(n As Integer) As String
    Select Case n
        Case 20: Dziesiatki = W_dwadziescia()
        Case 30: Dziesiatki = W_trzydziesci()
        Case 40: Dziesiatki = W_czterdziesci()
        Case 50: Dziesiatki = W_piecdziesiat()
        Case 60: Dziesiatki = W_szescdziesiat()
        Case 70: Dziesiatki = W_siedemdziesiat()
        Case 80: Dziesiatki = W_osiemdziesiat()
        Case 90: Dziesiatki = W_dziewiecdziesiat()
        Case Else: Dziesiatki = ""
    End Select
End Function

' --- Setki (100-900) ---
Private Function Setki(n As Integer) As String
    Select Case n
        Case 100: Setki = W_sto()
        Case 200: Setki = W_dwiescie()
        Case 300: Setki = W_trzysta()
        Case 400: Setki = W_czterysta()
        Case 500: Setki = W_piecset()
        Case 600: Setki = W_szescset()
        Case 700: Setki = W_siedemset()
        Case 800: Setki = W_osiemset()
        Case 900: Setki = W_dziewiecset()
        Case Else: Setki = ""
    End Select
End Function

' --- Trojka 0-999 na slowa ---
Private Function TrojkaSlownie(n As Integer, zenski As Boolean) As String
    Dim s As Integer, reszta As Integer
    Dim wynik As String
    wynik = ""

    s = (n \ 100) * 100
    reszta = n Mod 100

    If s > 0 Then wynik = Setki(s)

    If reszta >= 10 And reszta <= 19 Then
        If Len(wynik) > 0 Then wynik = wynik & " "
        wynik = wynik & Nastki(reszta)
    Else
        Dim dz As Integer
        Dim j As Integer
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

' --- Odmiana rzeczownika ---
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

' --- Odmiana "tysiac" ---
Private Function OdmianaTysiac(tys As Integer) As String
    Dim j As Integer: j = tys Mod 10
    Dim d As Integer: d = (tys Mod 100) \ 10
    If tys = 1 Then
        OdmianaTysiac = "tysi" & ChrW(261) & "c"
    ElseIf d = 1 Then
        OdmianaTysiac = "tysi" & ChrW(281) & "cy"
    ElseIf j >= 2 And j <= 4 Then
        OdmianaTysiac = "tysi" & ChrW(261) & "ce"
    Else
        OdmianaTysiac = "tysi" & ChrW(281) & "cy"
    End If
End Function

' --- Odmiana "milion" ---
Private Function OdmianaMilion(mln As Integer) As String
    Dim j As Integer: j = mln Mod 10
    Dim d As Integer: d = (mln Mod 100) \ 10
    If mln = 1 Then
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
'  GLOWNA FUNKCJA  =PL_TEKST(kwota)
' =============================================================================
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

    If liczba >= 1000000000# Then
        PL_TEKST = "Blad: liczba zbyt duza (max 999 999 999,99)"
        Exit Function
    End If

    ' Rozdziel na zlote i grosze
    Dim calkowita As Long
    Dim grosze As Integer
    calkowita = CLng(Int(liczba))
    Dim liczbaCur As Currency
    liczbaCur = CCur(liczba)
    grosze = CInt(liczbaCur * 100 - CCur(Int(liczbaCur)) * 100)
    If grosze >= 100 Then
        calkowita = calkowita + 1
        grosze = 0
    End If

    ' Rozloz na miliony / tysiace / reszta
    Dim mln As Integer
    Dim tys As Integer
    Dim res As Integer
    mln = CInt(calkowita \ 1000000)
    tys = CInt((calkowita Mod 1000000) \ 1000)
    res = CInt(calkowita Mod 1000)

    Dim wynik As String
    wynik = ""

    ' Miliony
    If mln > 0 Then
        wynik = TrojkaSlownie(mln, False) & " " & OdmianaMilion(mln)
    End If

    ' Tysiace
    If tys > 0 Then
        If Len(wynik) > 0 Then wynik = wynik & " "
        If tys = 1 Then
            wynik = wynik & OdmianaTysiac(1)
        Else
            wynik = wynik & TrojkaSlownie(tys, False) & " " & OdmianaTysiac(tys)
        End If
    End If

    ' Reszta
    If res > 0 Then
        If Len(wynik) > 0 Then wynik = wynik & " "
        wynik = wynik & TrojkaSlownie(res, False)
    End If

    ' Zero
    If calkowita = 0 Then wynik = "zero"

    ' Odmiana zloty
    Dim zloty As String
    zloty = Odmiana(calkowita, _
        "z" & ChrW(322) & "oty", _
        "z" & ChrW(322) & "ote", _
        "z" & ChrW(322) & "otych")

    ' Grosze
    Dim groszeTxt As String
    groszeTxt = Format(grosze, "00") & "/100 groszy"

    ' Zloz wynik
    Dim rezultat As String
    rezultat = wynik & " " & zloty & " (" & groszeTxt & ")"
    rezultat = UCase(Left(rezultat, 1)) & Mid(rezultat, 2)

    PL_TEKST = rezultat
    Exit Function
Blad:
    PL_TEKST = "Blad: " & Err.Description
End Function

