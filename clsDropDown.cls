VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public sht As Object
Function monthlist(shtn As String, rng As String) As String
    On Error GoTo exitscript
    Set sht = ThisWorkbook.Worksheets(shtn)
    With sht
        .Range(rng).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="January,February,March,April,May,June,July,August,September,October, November, December"
    End With
    monthlist = "passed"
    Exit Function
exitscript:
    monthlist = "failed"
End Function
