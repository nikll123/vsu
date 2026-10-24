Attribute VB_Name = "goo"
' Source - https://stackoverflow.com/a/79195996
' https://stackoverflow.com/questions/48484859/vba-excel-upload-transfer-data-to-goole-spreadsheet
' Posted by Ludwig Kuschke
' Retrieved 2026-08-18, License - CC BY-SA 4.0
Const rngFio = "S2:S81"
Const rngTitile = "R2:R81"
Const rngPosition = "I2:I81"
Const posLongRange = "J2:J81"

Sub test()
    getBchsData #1/1/2026#
End Sub

Sub googleSheetsTet()
    Dim httpReq As Object, url As String
    Set httpReq = CreateObject("MSXML2.ServerXMLHTTP")
    url = "https://script.google.com/macros/s/AKfycbyUqYX7Ytli5KJMNLukKxy8jrHXsr76i2240AQOcKP9YAemf8kdrn7rR2JRKAIXDFD0/exec"
    
    On Error GoTo errHandle
    httpReq.Open "GET", url, False
    httpReq.send
    Do Until httpReq.readyState = 4: Loop
    res = httpReq.responseText
    On Error GoTo 0
    
    MsgBox "Test is passed" & vbCr & res
    Exit Sub

errHandle:
    MsgBox "Test is failed" & vbCr & err.Description
End Sub

Sub getGoogleSheetsData(date1 As Date)
    Dim httpReq As Object, url As String
    Set httpReq = CreateObject("MSXML2.ServerXMLHTTP")
    url = "https://script.google.com/macros/s/AKfycbyUqYX7Ytli5KJMNLukKxy8jrHXsr76i2240AQOcKP9YAemf8kdrn7rR2JRKAIXDFD0/exec"
    
    On Error GoTo errHandle
    httpReq.Open "GET", url, False
    httpReq.send
    Do Until httpReq.readyState = 4: Loop
    res = httpReq.responseText
    On Error GoTo 0
    
    MsgBox "Test is passed" & vbCr & res
    Exit Sub

errHandle:
    MsgBox "Test is failed" & vbCr & err.Description
End Sub


Sub UploadStaff()
    Dim http As Object
    Dim url As String
    Dim sheet As Worksheet
    Dim row As Range
    'Dim data As Variant
    Dim jsonData As String
    Dim i As Long, j As Long

    ' URL of the Google Apps Script Web App
    url = "https://script.google.com/macros/s/AKfycbxBcrKfNMZ_CaHWyWLhfxfQhPdzOMlz64QnwmegU59O4MaO4wnpdK6YLcUb7wPUUA771w/exec"
           https://script.google.com/macros/s/AKfycbxBcrKfNMZ_CaHWyWLhfxfQhPdzOMlz64QnwmegU59O4MaO4wnpdK6YLcUb7wPUUA771w/exec
    
    ' Set the worksheet and range to upload
    Set sht = int_getWsStaff()
    ' data = sheet.UsedRange.Value ' Get all the data in the used range of the sheet
    dataFio = sht.Range(rngFio).Value
    dataTit = sht.Range(rngTitile).Value
    dataPos = sht.Range(rngPosition).Value
    dataPosLong = sht.Range(posLongRange).Value
    
    Dim data() As String
    ReDim data(UBound(dataFio), 3)
    j = LBound(data)
    
    For i = LBound(dataFio) To UBound(dataFio)
        If dataFio(i, 1) <> "" Then
            data(j, 0) = dataFio(i, 1)
            data(j, 1) = dataTit(i, 1)
            data(j, 2) = dataPos(i, 1)
            data(j, 3) = dataPosLong(i, 1)
            j = j + 1
        End If
    Next
       Sort2DArray data, SortColumn:=1, Ascending:=True
    ' Prepare the data as JSON
    jsonData = "["
    
    ' Loop through rows and columns to convert data into JSON format
    For i = 0 To UBound(data, 1)
        jsonData = jsonData & "["
        For j = 0 To UBound(data, 2)
            jsonData = jsonData & """" & data(i, j) & """"
            If j < UBound(data, 2) Then
                jsonData = jsonData & ","
            End If
        Next j
        jsonData = jsonData & "]"
        If i < UBound(data, 1) Then
            jsonData = jsonData & ","
        End If
    Next i
    
    jsonData = jsonData & "]"
    
    ' Create the HTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Send the data to the Google Sheets via the Google Apps Script web app
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonData
    
    ' Check the response
    If http.Status = 200 Then
        MsgBox "Data uploaded successfully!", vbInformation
    Else
        MsgBox "Error uploading data: " & http.Status, vbCritical
    End If
End Sub

