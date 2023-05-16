Sub SendEmails()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim wsNameList As Worksheet, wsBodyList As Worksheet, wsSubject As Worksheet, wsErrorList As Worksheet
    Dim rngNameList As Range, rngBodyList As Range
    Dim cellNameList As Range, cellBodyList As Range
    Dim name As String, email As String, subject As String, body As String, tag As String
    Dim flag As Integer, isAutoSend As Integer

    ' Set references to the sheets
    Set wsNameList = ThisWorkbook.Sheets("NameList")
    Set wsBodyList = ThisWorkbook.Sheets("BodyList")
    Set wsSubject = ThisWorkbook.Sheets("Subject")

    ' Check if ErrorList exists, if not then create it
    On Error Resume Next
    Set wsErrorList = ThisWorkbook.Sheets("ErrorList")
    On Error GoTo 0
    If wsErrorList Is Nothing Then
        Set wsErrorList = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsErrorList.name = "ErrorList"
        wsErrorList.Range("A1").Value = "Name"
        wsErrorList.Range("B1").Value = "Email"
        wsErrorList.Range("C1").Value = "Tag"
    End If

    ' Get the subject and the auto send flag
    subject = wsSubject.Range("A2").Value
    isAutoSend = wsSubject.Range("B2").Value

    ' Set reference to the name list range.
    Set rngNameList = wsNameList.Range("A2:C" & wsNameList.Cells(wsNameList.Rows.Count, "A").End(xlUp).Row)

    ' Initialize outlook
    Set OutApp = CreateObject("Outlook.Application")

    ' Loop through each name list cell.
    For Each cellNameList In rngNameList.Cells
        If cellNameList.Column = 1 Then
            name = cellNameList.Value
            email = cellNameList.Offset(0, 1).Value
            tag = cellNameList.Offset(0, 2).Value

            ' Loop through each body list cell to find the matching tag.
            Set rngBodyList = wsBodyList.Range("A2:C" & wsBodyList.Cells(wsBodyList.Rows.Count, "A").End(xlUp).Row)
            flag = 0
            For Each cellBodyList In rngBodyList.Cells
                If cellBodyList.Column = 1 Then
                    If cellBodyList.Value = "END" Then
                        Exit For
                    ElseIf cellBodyList.Value = tag Then
                        body = "Hello " & name & ",<br><br>" & cellBodyList.Offset(0, 1).Value & "<br><br>Best,<br>" & cellBodyList.Offset(0, 2).Value
                        flag = 1
                        Exit For
                    End If
                End If
            Next cellBodyList

            ' If there is no matching tag, copy the row to the ErrorList.
            If flag = 0 Then
               wsErrorList.Cells(wsErrorList.Cells(wsErrorList.Rows.Count, "A").End(xlUp).Row + 1, "A").Value = name
               wsErrorList.Cells(wsErrorList.Cells(wsErrorList.Rows.Count, "A").End(xlUp).Row, "B").Value = email
               wsErrorList.Cells(wsErrorList.Cells(wsErrorList.Rows.Count, "A").End(xlUp).Row, "C").Value = tag
            Else
               ' Create mail item.
               Set OutMail = OutApp.CreateItem(0)
                           With OutMail
                .To = email
                .CC = ""
                .BCC = ""
                .subject = subject
                .HTMLBody = body
                ' Auto send or display the email based on the isAutoSend flag.
                If isAutoSend = 1 Then
                    .Send
                Else
                    .Display
                End If
            End With

            ' Clean up.
            Set OutMail = Nothing
            End If
        End If
        Next cellNameList

    ' Clean up.
    Set OutApp = Nothing
    Set wsNameList = Nothing
    Set wsBodyList = Nothing
    Set wsSubject = Nothing
    Set wsErrorList = Nothing
    Set rngNameList = Nothing
    Set rngBodyList = Nothing
        
End Sub

