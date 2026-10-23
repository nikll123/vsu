Attribute VB_Name = "otpusk"
Const row1 = 2
Const row2 = 50
Const dateRow = 1
Const col1 = 7
Const col2 = 289
Const colFio = 1
Const colStatusPidr = 37


Sub test1()
     Debug.Print int_checkOtpusk(#4/3/2026#)
End Sub

Sub checkOtpusk()
    If int_checkOtpuskCurrent(0) = 0 Then
        d1 = int_getStaffDate2()
        MsgBox d1 & " - OK"
    End If
End Sub

Sub checkOtpusk2()
    checkOtpusk
    sink = int_checkOtpuskCurrent(2)
End Sub

Function int_checkOtpuskCurrent(days As Integer) As Integer
    int_checkOtpuskCurrent = -1
    Dim d1 As Date
    d1 = int_getStaffDate2()
    If days = 0 Then
        int_checkOtpuskCurrent = int_checkOtpusk(d1, " отпуск не в порядке ")
    Else
        Dim d2 As Date
        For i = 1 To days
            d2 = d1 + i
            int_checkOtpuskCurrent = int_checkOtpusk(d2, " будет изменение ")
        Next
    End If
End Function


' проверка соответствия данных на листе Штатка и Отпуска
Function int_checkOtpusk(date1 As Date, msg As String) As Integer
    int_checkOtpusk = 0
    If ActiveWorkbook.Name <> constStuffWbName Then ' штатка
        MsgBox constStuffWbName & " must be active!"
        int_checkOtpusk = -1
        Exit Function
    End If
    nameOtpusk = "отпуска"
    colDate1 = 0
    Set wsOtpusk = ActiveWorkbook.Worksheets(nameOtpusk)
    Dim d1 As String
    For colDate = col1 To col2
        d1 = wsOtpusk.Cells(dateRow, colDate).Value
        If Not testDate2(d1) Then ' дата
            MsgBox "the value is not a date: " & nameOtpusk & ".Cells(" & dateRow & "," & colDate & ")"
            Exit Function
        ElseIf d1 = date1 Then
            colDate1 = colDate
            Exit For
        End If
    Next colDate
    
    Debug.Assert colDate1 <> 0
    Dim fio As String
    Set wsState = int_getWsStaff()
    For r1 = row1 To row2
        'Debug.Print c.Value
        fio = wsOtpusk.Cells(r1, colFio)
        rFioState = getStateFioRow(fio)
        If rFioState = 0 Then
            MsgBox fio & " is not found on the " & constStuffWsName

            GoTo nextLoop
        End If
        
        If Val(wsOtpusk.Cells(r1, colDate1)) > 0 Then
            If "відпустка" = wsState.Cells(rFioState, colStatusPidr) Then
                ' ok
            Else
                MsgBox fio & msg & date1
                int_checkOtpusk = -1
            End If
        Else
        End If
        
        If "відпустка" = wsState.Cells(rFioState, colStatusPidr) Then
            If Val(wsOtpusk.Cells(r1, colDate1)) > 0 Then
                ' ok
            Else
                MsgBox fio & msg & date1
                int_checkOtpusk = -1
            End If
        Else
        End If
       
        
nextLoop:
    Next
    
End Function




