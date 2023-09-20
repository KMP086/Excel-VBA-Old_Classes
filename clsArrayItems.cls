VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArrayItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public runarr As New clsArrWbk
Public arryitem As New clsArrayItems
Public itemlist As New clsItemList
Public Function headerarr() As Boolean
    Dim i As Range
    Dim x As Long
    Dim rl As Long
    On Error GoTo skipthis
    With ThisWorkbook.Worksheets("Item list")
        rl = .Cells(Rows.Count, 1).End(xlUp).Row
        For Each i In .Range("A2:A" & rl)
            x = x + 1
            ThisWorkbook.Worksheets("Internal Audit Tracker").Range(Cells(1, x).Address).Value = CStr(i)
        Next i
    End With
    headerarr = True
    Exit Function
skipthis: headerarr = False
End Function
Public Function itemarr(shtname As String, headcnt As Long, rowcntr As Long) As String()
      Dim sqlitems(1 To 100) As String
      Dim hcntr As Long
      hcntr = 0
      Erase sqlitems()
      For hcntr = 1 To headcnt
            With ThisWorkbook.Worksheets(shtname)
                 sqlitems(hcntr) = .Cells(rowcntr + 1, hcntr)
                 ''Debug.Print (sqlitems(hcntr))
                 itemarr = sqlitems
            End With
      Next hcntr
End Function
Public Function rowarr(shtname As String, headcnt As Long, rowpos As Long) As String()
      Dim sqlitems(1 To 100) As String
      Dim hcntr As Long
      hcntr = 0
      Erase sqlitems()
      For hcntr = 1 To headcnt
            With ThisWorkbook.Worksheets(shtname)
                 sqlitems(hcntr) = .Cells(rowpos - 1, hcntr)
                 Debug.Print (sqlitems(hcntr))
                 rowarr = sqlitems
            End With
      Next hcntr

End Function
Public Function insertarray(shtname As String, tblname As String) As Boolean
    Dim arr(1 To 500) As String
    Dim item As Range
    Dim nitem As String
    Dim usedrng As Long
    Dim cnt As Long
    Dim c As Long
    Dim maxcol As Long
    Dim sqlarr(1 To 5000000) As String
    On Error GoTo skipthis
    With ThisWorkbook.Worksheets("RAM")
         If .Visible = False Then .Visible = True
         .Cells.Clear
    End With
    With ThisWorkbook.Worksheets(shtname)
         maxcol = .Cells(1, Columns.Count).End(xlToLeft).Column
         usedrng = .Cells(Rows.Count, 5).End(xlUp).Row
         cnt = 0
         For cnt = 1 To (usedrng - 1)
            c = 0
            Erase arr()
            For Each item In .Range(Cells(cnt + 1, 2).Address & ":" & Cells(cnt + 1, maxcol).Address)
                c = item.Column
                nitem = WorksheetFunction.IfError(item, "")
                
                nitem = Replace(nitem, "'", "")
                nitem = Replace(nitem, "*", "")
                
                If c = maxcol Then
                   arr(c) = "'" & CStr(nitem) & "'"
                ElseIf c = 3 Or c = 9 Or c = 11 Or c = 21 Then
                   arr(c) = "'" & Format(nitem, "MM/DD/YYYY HH:MM:SS") & "', "
                Else
                   arr(c) = "'" & CStr(nitem) & "', "
                End If
            
            Next item
                ''Debug.Print (cnt & "," & Join(arr))
            
            If cnt = (usedrng - r) Then
                ThisWorkbook.Worksheets("RAM").Range("A" & cnt).Value = "(" & CStr(Trim(Join(arr))) & ");"
            Else
                ThisWorkbook.Worksheets("RAM").Range("A" & cnt).Value = "(" & CStr(Trim(Join(arr))) & "),"
            End If
         Next cnt
         With ThisWorkbook.Worksheets("RAM")
            If .Visible = True Then .Visible = False
         End With

    End With
    insertarray = True
    Exit Function
skipthis:
    insertarray = False
End Function
Public Function updatearray(shtname As String, tblname As String) As Boolean
    Dim arr(1 To 500) As String
    Dim item As Range
    Dim nitem As String, id As String
    Dim usedrng As Long, cnt As Long, c As Long, maxcol As Long
    Dim sqlarr(1 To 5000000) As String
    On Error GoTo skipthis
    With ThisWorkbook.Worksheets("RAM")
         If .Visible = False Then .Visible = True
         .Cells.Clear
    End With
    With ThisWorkbook.Worksheets(shtname)
         maxcol = .Cells(1, Columns.Count).End(xlToLeft).Column
         If runarr.unhidden(ThisWorkbook.Name, shtname) = False Then Debug.Print ("Error Unhide")
         usedrng = .Cells(Rows.Count, 4).End(xlUp).Row
         If usedrng = 0 Then
            MsgBox "No 'update' word not found in column D"
            End
         ElseIf usedrng > 1 Then
            .Range("A:AH").Sort Key1:=.Columns(4), Order1:=xlAscending, Header:=xlYes
         End If
         
         cnt = 0
         For cnt = 1 To (usedrng - 1)
            c = 0
            Erase arr()
            For Each item In .Range(Cells(cnt + 1, 1).Address & ":" & Cells(cnt + 1, maxcol).Address)
                c = item.Column
                nitem = WorksheetFunction.IfError(item, "")
                nitem = Replace(nitem, "'", "")
                nitem = Replace(nitem, "*", "")
                
                
                If c = 1 Then
                   id = "'" & CStr(nitem) & "'"
                   If id = "'0'" Then End
                ElseIf c = maxcol Then
                   arr(c) = itemlist.param(c) & "='" & CStr(nitem) & "'"
                ElseIf c = 3 Or c = 10 Or c = 12 Then
                   arr(c) = itemlist.param(c) & "='" & Format(nitem, "MM/DD/YYYY HH:MM:SS") & "', "
                Else
                   arr(c) = itemlist.param(c) & "='" & CStr(nitem) & "', "
                End If
            
            Next item
            Debug.Print (cnt & "," & Join(arr))
            With ThisWorkbook.Worksheets("RAM")
                If cnt = (usedrng - r) Then
                   .Range("A" & cnt).Value = "Update tblname Set(" & CStr(Trim(Join(arr))) & ") Where id = " & id & ";"
                Else
                   ThisWorkbook.Worksheets("RAM").Range("A" & cnt).Value = "Update " & tblname & " Set(" & CStr(Trim(Join(arr))) & ")  Where id = " & id & ","
                End If
            End With
         Next cnt
         With ThisWorkbook.Worksheets("RAM")
            If .Visible = True Then .Visible = False
         End With

    End With
    updatearray = True
    Exit Function
skipthis:
    updatearray = False
End Function

