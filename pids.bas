Attribute VB_Name = "pids"
Sub GetDataActual()
    If ActiveWorkbook.Name <> constPidsWbName Then ' пидс
        MsgBox constPidsWbName & " must be active!"
        Exit Sub
    End If
    Dim wsArc As Object
    Dim dateCur As Date
    Dim datePrev As Date
    Set wsArc = ActiveWorkbook.Worksheets("archive")
    dateCur = int_getStaffDate2()
    qq = getDateRow(wsArc, dateCur)
    datePrev = dateCur - 1
    qq = getDateRow(wsArc, datePrev)
    
End Sub


Sub AddCurData()
    dateCur = int_getStaffDate2()
    datePrev = dateCur - 1
End Sub

Function getDateRow(ws As Object, d As Date) As Long
    colData = 1
    getDateRow = -1
    For i = 2 To 10000
        If ws.Cells(i, colData).Value = d Then
            getDateRow = i
            Exit Function
        End If
    Next i
End Function
