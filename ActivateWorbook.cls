VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActWbk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public sht As Object
Public Function chkwbk(fileloc As String, wbkname As String) As String
    If Dir(fileloc & "\" & wbkname) = "" Then
       MsgBox "File" & wbkname & " wasn't found!"
       chkwbk = ""
       End
    ElseIf Dir(fileloc & "\" & wbkname) <> "" Then
       Workbooks.Open (fileloc & "\" & wbkname)
       chkwbk = wbkname
    End If
End Function

