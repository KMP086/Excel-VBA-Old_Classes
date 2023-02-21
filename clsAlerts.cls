VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAlerts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function alerts(tf As String)
    If tf = "t" Then
       Application.DisplayAlerts = True
       Application.DisplayStatusBar = True
       Application.EnableEvents = True
       Application.ScreenUpdating = True
    ElseIf tf = "f" Then
       Application.DisplayAlerts = False
       Application.DisplayStatusBar = False
       Application.EnableEvents = False
       Application.ScreenUpdating = False
    End If
End Function
