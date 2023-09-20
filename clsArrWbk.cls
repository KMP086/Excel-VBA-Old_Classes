VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArrWbk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function copypaste(wbkname As String, refsht As String, rowpos As Long, colpos As Long, newwrksht As String, newrow As Long, newcol As Long, isvalue As Boolean) As Boolean
    Dim lrow As Long
    On Error GoTo skipthis:
    If colpos = 0 Then GoTo skipthis
    lrow = Workbooks(wbkname).Worksheets(refsht).Cells(Rows.Count, colpos).End(xlUp).Row
    Workbooks(wbkname).Worksheets(refsht).Range(Cells(rowpos, colpos).Address & ":" & Cells(lrow, colpos).Address).Copy
    If isvalue = False Then
        ThisWorkbook.Worksheets(newwrksht).Range(Cells(newrow, newcol).Address).PasteSpecial
    ElseIf isvalue = True Then
        ThisWorkbook.Worksheets(newwrksht).Range(Cells(newrow, newcol).Address).PasteSpecial Paste:=xlPasteValues
    End If
    copypaste = True
    Exit Function
skipthis:
    Debug.Print ("Column error!! ref column no: " & colpos)
    copypaste = False
End Function
Public Function clearsht(wbkname As String, shtname As String) As Boolean
    With Workbooks(wbkname).Worksheets(shtname)
       usedrow = .UsedRange.Rows.Count
       usedcolumn = .UsedRange.Columns.Count
    End With
    On Error GoTo skipthis:
    With Workbooks(wbkname).Worksheets(shtname)
       If unhidden(wbkname, shtname) = True Then
          .Range(Cells(2, 1).Address & ":" & Cells(usedrow, usedcolumn).Address).Clear
          .Range(Cells(2, 1).Address & ":" & Cells(usedrow, usedcolumn).Address).ClearFormats
       End If
    End With
    clearsht = True
    Exit Function
skipthis: clearsht = False
End Function
Public Function unhidden(wbkname As String, shtname As String)
    On Error GoTo skipthis:
       With Workbooks(wbkname).Worksheets(shtname)
           usedrow = .UsedRange.Rows.Count
           usedcolumn = .UsedRange.Columns.Count
           .Columns.EntireColumn.Hidden = False
           .Rows.EntireRow.Hidden = False
           If .AutoFilterMode Then .AutoFilterMode = False
       End With
    unhidden = True
    Exit Function
skipthis: unhidden = False
End Function
Public Function replacethis(wbkname As String, shtname As String, txtfind As String, txtreplace As String) As Boolean
    On Error GoTo skipthis:
    With Workbooks(wbkname).Worksheets(shtname)
         .Cells.Replace What:=txtfind, Replacement:=txtreplace, LookAt:= _
         xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
         ReplaceFormat:=False
    End With
    replacethis = True
    Exit Function
skipthis:
    replacethis = False
End Function
Public Function numberform(wbkname As String, shtname As String, startrow As Long, colpos As Long, formtype As String) As Boolean
    Dim lrow As Long
    On Error GoTo skipthis:
    lrow = lastrow(wbkname, shtname, colpos)
    Workbooks(wbkname).Worksheets(shtname).Range(Cells(startrow, colpos).Address & ":" & Cells(lrow, colpos).Address).NumberFormat = formtype
    numberform = True
    Exit Function
skipthis:
    numberform = False
End Function
Public Function deciformat(nitem As String) As String
    Dim tem, tem2
    tem = Split(nitem, ",")
    tem2 = Split(nitem, ".")
    If UBound(tem) > 0 Then
        If UBound(tem2) = 0 Then
            leftvalue = Left(nitem, Len(nitem) - Len(tem(UBound(tem))) - 1)
            newvalue = Replace(Replace(leftvalue, ",", ""), ".", "") & "." & tem(UBound(tem))
        Else
            If Len(tem(UBound(tem))) < Len(tem2(UBound(tem2))) Then
                leftvalue = Left(nitem, Len(nitem) - Len(tem(UBound(tem))) - 1)
                newvalue = Replace(Replace(leftvalue, ",", ""), ".", "") & "." & tem(UBound(tem))
            End If
            If Len(tem(UBound(tem))) > Len(tem2(UBound(tem2))) Then
                leftvalue = Left(nitem, Len(nitem) - Len(tem2(UBound(tem2))) - 1)
                newvalue = Replace(Replace(leftvalue, ",", ""), ".", "") & "." & tem2(UBound(tem2))
            End If
        End If
    Else
        If UBound(tem) = 0 And UBound(tem2) = 0 Then
            newvalue = nitem
        Else
            leftvalue = Left(nitem, Len(nitem) - Len(tem2(UBound(tem2))) - 1)
            newvalue = Replace(Replace(leftvalue, ",", ""), ".", "") & "." & tem2(UBound(tem2))
        End If
    End If
    deciformat = CStr(newvalue)
End Function

