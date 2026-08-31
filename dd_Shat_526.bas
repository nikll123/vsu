Attribute VB_Name = "dd_Shat_526"
Dim wbStuff As Object
Dim wsStuff As Object
Dim wb526 As Object
Dim ws526 As Object



Sub Ceels()
On Error GoTo filesNotOpened
Dim wsRpu As Object
Dim wsTabel As Object
Set wsStuff = Workbooks("штатка.xlsx").Worksheets("526-518")
Set ws526 = Workbooks("рРЕБ Ф5.26 - .xlsx").Worksheets("БЧС")
On Error GoTo 0


    For Each c1 In wsStuff.Range("C6:AE13")
        c1.Select
        Set Cel = ws526.Cells(c1.row, c1.Column)
        Set CelSt = wsStuff.Cells(5, c1.Column)
        If Left(Cel.Value, 1) <> "=" Then
            If ws526.Cells(c1.row, c1.Column).Value = c1.Value Then
               Cel.Interior.Color = CelSt.Interior.Color 'шо було
            Else
               ws526.Cells(c1.row, c1.Column).Value = c1.Value
               Cel.Interior.Color = 65535 'жовтий
            End If
        End If
'        If Abs(Rng.Value) < 0.01 Then Rng.Value = 0
    Next
    
    
    wsStuff.Activate
    Range("C6:AE13").copy

    Range("C26").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("A25") = Date
    
    
Exit Sub



filesNotOpened:
    MsgBox "Файли: рРЕБ Ф5.26 - .xlsx та штатка.xlsx  повинні бути вікриті!"

End Sub
