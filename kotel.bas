Attribute VB_Name = "kotel"
Sub test()
    Dim dd  As Date
    
    dd = Date - 10000

    kotelCreate dd, 0
End Sub


Sub kotelHistCreate()
    If ActiveWorkbook.Name <> constStuffWbName Then ' штатка
        MsgBox constStuffWbName & " must be active!"
        Exit Sub
    End If

    Dim wsKotel As Object
    Set wsKotel = ActiveSheet
    If Not "котел" = wsKotel.Name Then
        MsgBox "Sheet name must be " & "котел"
        Exit Sub
    End If
    dToday = wsKotel.Cells(3, 15).Value
    dTomorrow = dToday + 1
    'dYesterday = dToday - 1
    lastRow = 200
    kotelToday = wsKotel.Cells(2, 3).Value
    For i = 47 To lastRow
        If wsKotel.Cells(i, 1).Value = dToday Then
           wsKotel.Cells(i, 2).Value = kotelToday
           kotelYesterday = wsKotel.Cells(i, 3).Value
           wsKotel.Cells(2, 5).Value = kotelYesterday
        ElseIf wsKotel.Cells(i, 1).Value = dTomorrow Then
           wsKotel.Cells(i, 3).Value = kotelToday
           Exit For
        End If
    Next i
    
    If i > lastRow Then
        MsgBox "Error, not all dates were found"
    End If
    

'    fn = "C:\RR\stroy158\котел\рРЕБ КОТЕЛ.docx"
'    fnPdf = "рРЕБ КОТЕЛ " & txtDate & ".pdf"
'    ffnPdf = "C:\RR\stroy158\котел\" & fnPdf
'    excelPath = "C:\RR\ШТАТКА\штатка.xlsx"
'    If Len(Dir(fn)) > 0 Then   ' check for existance of the file
'        Dim WordApp As Object
'        Dim wordIsOpened As Boolean
'
'        On Error Resume Next 'If Word is already opened
'
''        Set WordApp = GetObject("Word.Aspplication")
''        wordIsOpened = Not WordApp Is Nothing
''        If Not wordIsOpened Then
'            'Word wasnґt running, so open Application
'            Set WordApp = CreateObject("Word.Application")
''        End If
'
'        WordApp.Visible = True
'        On Error GoTo 0
'
'
'        Dim kotelDoc As Object
'        'Check for the existence of the file
'        Set kotelDoc = WordApp.Documents.Open(fn)
'       ' kotelDoc.MailMerge.OpenDataSource Name:=excelPath, SQLStatement:="SELECT * FROM `'котел$'`"
'        kotelDoc.MailMerge.Execute
'        kotelDoc.Save
'
'        wdExportFormatPDF = 17
'        kotelDoc.ExportAsFixedFormat ffnPdf, wdExportFormatPDF
'        kotelDoc.Close SaveChanges:=False
'        MsgBox "File is created: " & ffnPdf
'        '        OutputFileName:=ffnPdf, _
'        '            ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, _
'        '            OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, _
'        '            From:=1, To:=1, Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
'        '            CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
'        '            BitmapMissingFonts:=True, UseISO19005_1:=False
'
'        '      WordApp.Visible = True
'        If wsKotel.Cells(2, 5).Value <> wsKotel.Cells(2, 3).Value Then
'            txtDate = wsKotel.Cells(3, 2).text
'            fn = "C:\RR\stroy158\котел\рРЕБ КОТЕЛ (уточнення).docx"
'            Set kotelDoc = WordApp.Documents.Open(fn)
'            kotelDoc.Save
'            fnPdf = "рРЕБ КОТЕЛ (уточнення) " & txtDate & ".pdf"
'            ffnPdf = "C:\RR\stroy158\котел\" & fnPdf
'            kotelDoc.ExportAsFixedFormat ffnPdf, wdExportFormatPDF
'            MsgBox "File is created: " & ffnPdf
'            wsKotel.Cells(2, 5).Value = wsKotel.Cells(2, 3).Value
'            kotelDoc.Close SaveChanges:=False
'        End If
'
''        If Not wordIsOpened Then
'         '   WordApp.Quit
''        End If
'        Set WordApp = Nothing
'    Else
'        MsgBox "Couldnґt open '" & fn & "'!", vbInformation, "Error loading Document"
'    End If
'

End Sub

Sub kotel2()

    Dim wrdApp As Word.Application
    Dim WrdDoc As Word.Document
    fn = "C:\RR\stroy158\котел\рРЕБ КОТЕЛ.docx"
 
    Set wrdApp = New Word.Application
    'wrdApp.Visible = True
    'wrdApp.Activate
    
    With wrdApp
        
          .Documents.Open (fn)
          
    End With
    
    wrdApp.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
    wrdApp.Quit
    
    Set wrdApp = Nothing

End Sub
