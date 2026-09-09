Attribute VB_Name = "dd_Shat_518"
Dim wbStuff As Object
Dim wsStuff As Object
Dim wb518 As Object
Dim ws518 As Object
Sub Ceels()
    On Error GoTo filesNotOpened
    Dim wsRpu As Object
    Dim wsTabel As Object
    Set wsStuff = Workbooks("штатка.xlsx").Worksheets("штатка")
    Set ws518 = Workbooks("р–≈Ѕ ‘5.18 - .xlsx").Worksheets("рота –≈Ѕ")
    On Error GoTo 0

    ws518.Range("K5") = Date
    For Each c1 In wsStuff.Range("AC142:AC162")
        c1.Select
        For Each c2 In ws518.Range("E36:E56")
            If c1.Value = c2.Value Then
                Set Cel = ws518.Cells(c2.Row, 13)
                Set CelSt = wsStuff.Cells(c1.Row, 24)
                If ws518.Cells(c2.Row, 13).Value = wsStuff.Cells(c1.Row, 24).Value Then
                   Cel.Interior.Color = CelSt.Interior.Color 'шо було
                Else
                   ws518.Cells(c2.Row, 13).Value = wsStuff.Cells(c1.Row, 24).Value
                   Cel.Interior.Color = 65535 'жовтий
                End If
            End If
    '        If Abs(Rng.Value) < 0.01 Then Rng.Value = 0
        Next
    Next
    If ws518.Cells(58, 13).Value <> wsStuff.Cells(163, 24).Value Then
        MsgBox "Ќ≈ сп≥вподаЇ «ј√јЋ№Ќј к≥льк≥сть!!!"
    End If
    
''''    wsStuff.Activate
''''    Range("C6:AE13").Copy
''''
''''    Range("C26").Select
''''    Selection.PasteSpecial Paste:=xlPasteValues
''''    Range("A25") = Date
    
    
Exit Sub

filesNotOpened:
    MsgBox "‘айли: р–≈Ѕ ‘5.18 - .xlsx та штатка.xlsx  повинн≥ бути в≥крит≥!"

End Sub

