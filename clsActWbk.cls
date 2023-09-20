VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsActWbk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function actwks(wrksht As String) As Object
  If wrksht <> "" Then
    Set actwks = ThisWorkbook.Worksheets(wrksht)
    Exit Function
  Else
    Debug.Print ("No declared worksheet!")
    Exit Function
  End If

End Function
Public Function wbkopen(filepath As String) As String
    If Dir(filepath) <> "" Then
       Workbooks.Open (filepath)
       wbkopen = Dir(filepath)
    Else
       Debug.Print ("No file found!")
       End
    End If
End Function
Public Function wbkclose(filename)
    actwbk.appsettings (True)
    Workbooks(filename).Close SaveChanges:=False
End Function
Public Function appsettings(appwbk As Boolean) As Boolean
    On Error GoTo skipthis
    If appwbk = False Then
       Application.DisplayAlerts = False
       Application.ScreenUpdating = False
       Application.DisplayStatusBar = False
       Application.EnableEvents = False
       Application.Calculation = xlManual
       If ThisWorkbook.AutoSaveOn = True Then ThisWorkbook.AutoSaveOn = False
    ElseIf appwbk = True Then
       Application.DisplayAlerts = True
       Application.ScreenUpdating = True
       Application.DisplayStatusBar = True
       Application.EnableEvents = True
       Application.Calculation = xlAutomatic
       If ThisWorkbook.AutoSaveOn = False Then ThisWorkbook.AutoSaveOn = True
    End If
    appsettings = True
Exit Function
skipthis: appsettings = False
End Function
Public Function confirmactive(filename As String)
    Dim ans As Integer
    ans = MsgBox("Are you sure to activate this", vbYesNo + vbQuestion, "Activating Option")
    If ans = vbNo Then End
End Function
