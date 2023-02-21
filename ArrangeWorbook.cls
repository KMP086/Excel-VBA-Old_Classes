VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arrwbk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public tbl As Object
Public sht As Object
Public x As Variant
Public y As Long
Public lc As Long
Public uc As Long
Public i As Range
Public fr As Long
Public spr As Long
Public Function coorheader(wbkname As String, shtname As String, fitem As String, r As Long) As Long
    Set sht = Workbooks(wbkname).Worksheets(shtname)
    lc = sht.UsedRange.Columns.Count
    On Error GoTo altvalue
    coorheader = Application.WorksheetFunction.Match(fitem, sht.Range(Cells(r, 1).Address & ":" & _
    Cells(r, lc).Address), 0)
    Exit Function
altvalue:
    coorheader = 0
End Function
Public Function afilter(owrkbook As String, oshtname As String, osr As Long, sc As Long) As Long
    Set sht = Workbooks(owrkbook).Worksheets(oshtname)
    Set tbl = ThisWorkbook.Worksheets("Table List")
        With sht
        lc = .Cells(osr, Columns.Count).End(xlToLeft).Column
        ur = .Cells(Rows.Count, lc).End(xlUp).Row
        If .AutoFilterMode Then .AutoFilterMode = False
        End With
        sht.Range(Cells(osr, 1).Address & ":" & Cells(osr, lc).Address).AutoFilter Field:=(lc), Criteria1:=1
        For Each i In sht.Range(Cells(2, sc).Address & ":" & Cells(ur, lc).Address).SpecialCells(xlCellTypeVisible)
            fr = sht.Cells(Rows.Count, sc).End(xlUp).Row
            If fr = 1 Then
               MsgBox "no filtered item"
               End
            ElseIf fr > 1 Then
                spr = i.Row
                GoTo skipthisscript
            End If
        Next i
    
skipthisscript:
    afilter = spr
End Function

Public Function htrimclean(wbkname As String, shtname As String, rowpos As Long, colpos As Long) As Long
    On Error GoTo endthisscript
    With Workbooks(wbkname).Worksheets(shtname)
        lc = .UsedRange.Columns.Count
        For Each c In .Range(Cells(rowpos, colpos).Address & ":" & Cells(rowpos, lc).Address)
            .Cells(rowpos, c.Column).Value = Application.WorksheetFunction.Clean( _
            Application.WorksheetFunction.Trim(.Cells(rowpos, c.Column).Value))
        Next c
        
    End With
    htrimclean = 1
    Exit Function
endthisscript:
    htrimclean = 0
End Function

Public Function cpshtspecial(fname As String, shtname As String, rptshtname As String, _
sr As Long, sc As Long, lr As Long, lc As Long, pr As Long, pc As Long) As Long
    On Error GoTo endthisscript
    Set sht = Workbooks(fname).Worksheets(shtname)
    With sht
        .Range(Cells(sr, sc).Address & ":" & Cells(lr, lc).Address).SpecialCells(xlCellTypeVisible).Copy
        ThisWorkbook.Worksheets(rptshtname).Range(Cells(pr, pc).Address).PasteSpecial Paste:=xlPasteValues
    End With
    cpshtspecial = 1
    Exit Function
endthisscript:
    cpshtspecial = 0
End Function


