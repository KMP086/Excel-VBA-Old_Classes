Public Function emailout(fname As String) As Boolean
    Dim outlookApp As Object
    Dim outmail As Object
    Dim filepath As String
    Dim filename As String
    Dim ucol As Long
    Dim urow As Long
    Dim fileadd As String
    Dim sentto, sentcc, sentsub, sentgreet, sentmsg, sentfare, sentsig As String
    On Error GoTo skipthis
    With ThisWorkbook.Worksheets("Item list")
         fileadd = .Cells(1, 4)
         sentsub = .Cells(2, 4)
         sentto = .Cells(4, 2)
         sentcc = .Cells(4, 4)
         sentgreet = .Cells(5, 2)
         sentmsg = .Cells(10, 2)
         sentfare = .Cells(15, 2)
         sentsig = .Cells(20, 2)
    End With
    Set outlookApp = CreateObject("Outlook.Application")
    Set outmail = outlookApp.CreateItem(olMailItem)
    With outmail
        .to = sentto
        .CC = sentcc
        .Subject = sentsub
        filepath = fileadd & "\" & fname & ".png"
        .Htmlbody = "<br>" & sentgreet & "<br><br>" & sentmsg & "<br><img src= '" & filepath _
        & "' style='border:0'><br>" & sentfare & "<br><br>" & sentsig & "<br>"
        .Display

    End With

    Set outmail = Nothing
    Set outlookApp = Nothing
    emailout = True
    Exit Function
skipthis: emailout = False

End Function
