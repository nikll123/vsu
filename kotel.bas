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

Sub kotelLBS()
    r1 = 0
    'rng1 = "A2:J2"
    curDate = Range("M2")
    lastDataRow = Cells(Rows.Count, "A").End(xlUp).row
    For Each c In Range("A2", "A" & lastDataRow)
        If c = curDate Then
            r1 = c.row
            Exit For
        End If
    Next
    If r1 = 0 Then
        r1 = lastDataRow + 1
    End If
    Set rr = Range("A" & r1 & ":K" & r1)
    Range("M2:W2").copy
    rr.PasteSpecial (xlPasteValues)
    For col = 3 To 4  ' C, D
        Set cc = Cells(r1, col)
        If cc.Value < 50 Then
            cc.Interior.Color = RGB(255, 192, 192)
        Else
            cc.ClearFormats
        End If
    Next
    Application.CutCopyMode = False
    
'    On Error Resume Next
'    Set appWord = GetObject(, "Word.Application")
'    If appWord Is Nothing Then
'        Set appWord = CreateObject("Word.Application")
'    End If
'    On Error GoTo 0
'
'    Const lbsFileName = "R:\stroy158\котел\лбс\Перевірка ЛБЗ рота РЕБ.doc"
'    appWord.Visible = True
'    Set WrdDoc = appWord.Documents.Open(lbsFileName, Format:=0)
'    strWorkbookPath = ActiveWorkbook.FullName
'    Set myMerge = WrdDoc.MailMerge
'    If myMerge.State = wdMainAndSourceAndHeader Or myMerge.State = wdMainAndDataSource Then
'         With myMerge.DataSource
'            .FirstRecord = 1
'            .LastRecord = 1
'         End With
'    End If
'
'
'    WrdDoc.MailMerge.Execute
'
'        With WrdDoc.MailMerge
'        ' Connect the active Excel workbook as the data source
'        .OpenDataSource _
'            Name:=strWorkbookPath, _
'            ConfirmConversions:=False, _
'            ReadOnly:=False, _
'            LinkToSource:=True, _
'            AddToRecentFiles:=False, _
'            Connection:="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strWorkbookPath & ";", _
'            SQLStatement:="SELECT * FROM [Sheet1$]" ' Change Sheet1 to your exact worksheet name
'
'        ' Set destination to a new document (wdSendToNewDocument = 0)
'        .Destination = 0
'
'        ' Suppress blank lines if necessary
'        .SuppressBlankLines = True
'
'        ' Run the merge
'        .Execute Pause:=False
'    End With
'
'    ' Close the original template document without saving changes
'    wdDoc.Close SaveChanges:=0 ' wdDoNotSaveChanges = 0
'
'    ' Clean up object references
'    Set wdDoc = Nothing
'    Set wdApp = Nothing
'
'    MsgBox "Mail merge completed successfully!", vbInformation
'
'
'
'    newFn = "R:\stroy158\котел\лбс\Перевірка ЛБЗ рота РЕБ " & curDate & ".doc"
'    WrdDoc.SaveAs2 (newFn)
End Sub
