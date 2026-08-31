Attribute VB_Name = "ubd"
Public currentFioRow As Long
Public Const fnUbd = "UBD.xlsx" '  "C:\RR\табель\UBD.xlsx"

Sub getUpdPivotData()
Dim wbTabel As Object
Dim wbUBD As Object
If ActiveWorkbook.Name <> fnUbd Then
    MsgBox fnUbd & " must be active!"
    Exit Sub
End If
Set wbUBD = ActiveWorkbook
Dim rowUbdPivData As Long
rowUbdPivData = 2
For i = 1 To 100
    fn1 = wbUBD.Sheets("files").Range("A" & i).Value
    If Trim(fn1) = "" Then
        Exit For
    End If
    Debug.Print fn1
    Set wbTabel = Workbooks.Open(fn1)
    rowUbdPivData = processFile(wbTabel, wbUBD, rowUbdPivData)
    wbTabel.Close False
Next
End Sub


Function processFile(wbTabel As Object, wbUBD As Object, rowUbdPivData As Long) As Long
    Dim fio As String
    Dim wsUbdPivData As Object
    Set wsUbdPivData = wbUBD.Sheets("pivotdata")

    Dim wsTabel As Object
    Dim r As Long
    Set wsTabel = wbTabel.Worksheets("Табель")
    
    For r = 2 To 70
        If wbUBD.Worksheets("fio").Cells(r, 2) Then
            fio = getFio(wbUBD, r)
            'Debug.Print fio
            dateRow = 7
            colDate1 = 9
            
            fioRow = findFioRow(wsTabel, fio, True)
            If fioRow > 0 Then
                For c = colDate1 To colDate1 + 31
                    x = wsTabel.Cells(fioRow, c).Value
                    If x = 100 Or x = 30 Then
                        'Debug.Print x
                        wsUbdPivData.Cells(rowUbdPivData, 1).Value = fio
                        wsUbdPivData.Cells(rowUbdPivData, 2).Value = CDate(wsTabel.Cells(dateRow, c).Value)
                        wsUbdPivData.Cells(rowUbdPivData, 3).Value = x
                        wsUbdPivData.Cells(rowUbdPivData, 4).Value = 1
                        rowUbdPivData = rowUbdPivData + 1
                    End If
                Next
            Else
                Debug.Print fio & " is not found in: " & wbTabel.Name
            End If
        Else
        
        End If
    Next
    processFile = rowUbdPivData
End Function
Sub getTabelData(wsTabel As Object, fio As String)
    dateRow = 7
    colDate1 = 9
    fioRow = findFioRow(wsTabel, fio)
    For c = colDate1 To colDate1 + 31
        Debug.Print wsTabel.Cells(fioRow, c).Value
    Next
    
End Sub

Sub resetFioRow()
    currentFioRow = 0
End Sub
Function getFio(wbUBD As Object, r As Long) As String
    getFio = wbUBD.Worksheets("fio").Cells(r, 1).Value
End Function

Sub getPeriods()
    Set wbUBD = ActiveWorkbook
    If ActiveWorkbook.Name <> fnUbd Then
        MsgBox fnUbd & " must be active!"
        Exit Sub
    End If
    Set wsReport = wbUBD.Worksheets("report")
    Set wsPivotData = wbUBD.Worksheets("pivotdata")
    Dim dates(200) As Date
    ixDate = 0
    For r = 2 To 67
        If wsReport.Cells(r, 1).Value <> "" Then
            fio = wsReport.Cells(r, 1).Value
            periods = ""
            date1 = 0
            darw2 = 0
            ixDate = 0
            For r2 = 2 To 100000
                
                If wsPivotData.Cells(r2, 1).Value = fio Then
                    If wsPivotData.Cells(r2, 3).Value = 100 Then
                        If ixDate = 0 Then
                            dates(ixDate) = wsPivotData.Cells(r2, 2).Value
                            ixDate = ixDate + 1
                        Else
                            If dates(ixDate - 1) <> wsPivotData.Cells(r2, 2).Value - 1 Then
                                If ixDate = 1 Then
                                    periods = periods & dates(0)
                                ElseIf ixDate = 2 Then
                                    periods = periods & dates(0) & "; " & dates(1)
                                Else
                                    periods = periods & " з " & dates(0) & " по " & dates(ixDate - 1)
                                End If
                                periods = periods & "; " & vbCrLf
                                ixDate = 0
                            Else
                                
                            End If
                        
                            dates(ixDate) = wsPivotData.Cells(r2, 2).Value
                            ixDate = ixDate + 1
                        
                        End If
                        
                    Else

                    End If
                End If
            Next
            
            If ixDate = 1 Then
                periods = periods & dates(0)
            ElseIf ixDate = 2 Then
                periods = periods & dates(0) & ";" & dates(1)
            Else
                periods = periods & " з " & dates(0) & " по " & dates(ixDate - 1)
            End If
            periods = periods & ";"
                                
        End If
        If fio <> "" Then
            wsReport.Cells(r, 2) = periods
            Debug.Print fio
            Debug.Print periods
        End If
        periods = ""
        fio = ""
    Next
    
End Sub
