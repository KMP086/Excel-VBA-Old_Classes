Public Function createimg(wksht As String, filerng As String, fpath As String, fname As String) As Boolean
    On Error GoTo skipthis
    Dim strrng As Range
    Dim rl As Long
    
    With ThisWorkbook.Worksheets(wksht)
        If .AutoFilterMode Then .AutoFilterMode = False
        rl = .Cells(Rows.Count, 1).End(xlUp).row
        .Range("A3:U" & rl).AutoFilter Field:=21, Criteria1:="<>"
        rl = .Cells(Rows.Count, 1).End(xlUp).row
        Set strrng = ThisWorkbook.Worksheets(wksht).Range(filerng & rl)
    End With
    strrng.CopyPicture xlScreen, xlPicture
    lWidth = strrng.Width
    lHeight = strrng.Height
    
    Set Cht = ActiveSheet.ChartObjects.Add(Left:=0, Top:=0, Width:=lWidth, Height:=lHeight)
    Cht.Activate
    
    With Cht.Chart
      .Paste
      .Export filename:=fpath & "\" & fname & ".png", Filtername:="PNG"
    End With
    
    Cht.Delete
    createpng = True
    Exit Function
skipthis: createpng = False

End Function
