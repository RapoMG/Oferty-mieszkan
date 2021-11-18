Dim LR As Integer, LC As Integer
Sub Main()
'Pobiera dane ze strony
'Odlicza linie
Call Licznik
'usówa kolor i nanosi na nowo
'Call Pasy <----
'dodaje warunki
Call FormatWarunk
'testowae podsumowanie
Call wiadomosc
End Sub
Sub Kolor()
Call Licznik
Call Pasy
End Sub

Sub Licznik()
'odlicza ostani wiersz
LR = Range("A2").End(xlDown).Row
' i kolumnę
LC = Cells(2, Columns.Count).End(xlToLeft).Column
End Sub

Sub Pasy()
' Obszar informacji - usunięcie tła
Range(Cells(3, 1), Cells(LR, Columns.Count)).Select

    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
' Tło w co 2 wierszu
For i = 3 To LR
    If i Mod 2 = 0 Then
        Range(Cells(i, 1), Cells(i, LC)).Select
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
        End With
    End If
Next i

Range("A1").Select
End Sub

Sub FormatWarunk()
Dim LWar As Integer
Dim War As String
Dim War2 As String ' tylko test
Dim i As Integer 'tylko test
i = 10 ' tylko test
'For i = 3 To LR
    'Rows(i & ":" & i).Select
    Rows("10:10").Select
    LWar = Selection.FormatConditions.Count
    If LWar = 0 Then
        MsgBox ("działa")
        War = "=IF(ISBLANK(R" & i & "C" & LC & "),FALSE,R" & i & "C" & LC & "+7<TODAY())"
        Range("a1").FormulaR1C1 = War
        War2 = Range("a1").FormulaR1C1
        MsgBox ("pobrało")
        MsgBox War2
        Rows("10:10").Select
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:=War2
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .Pattern = xlLightUp
            .PatternColorIndex = xlAutomatic
            .ColorIndex = xlAutomatic
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
            MsgBox ("działa i armaty")
    End If
    MsgBox ("po if")
'Next i
End Sub
