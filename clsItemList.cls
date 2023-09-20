VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function monthlistint(monthname As String) As Integer
    Select Case monthname
        Case "January": monthlistint = 1
        Case "February": monthlistint = 2
        Case "March": monthlistint = 3
        Case "April": monthlistint = 4
        Case "May": monthlistint = 5
        Case "June": monthlistint = 6
        Case "July": monthlistint = 7
        Case "August": monthlistint = 8
        Case "September": monthlistint = 9
        Case "October": monthlistint = 10
        Case "November": monthlistint = 11
        Case "December": monthlistint = 12
     End Select
End Function
Public Function sqllogic(yesno As String) As Integer
    Select Case yesno
        Case "Yes": sqllogic = 1
        Case "No": sqllogic = 0
    End Select
End Function
Public Function itemdroplist(shtname As String, cellpos As String, newlist As String) As Boolean
        On Error GoTo skipthis
        Dim monthlist As String
        With ThisWorkbook.Worksheets(shtname)
            .Range(cellpos).Validation.Delete
            .Range(cellpos).Validation.Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Formula1:=newlist
        End With
        itemdroplist = True
        Exit Function
skipthis: itemdroplist = False
End Function
Public Function paracol(tblname) As String
       If tblname = "SQL_MNLDB.dbo.tblIsscError" Then
           paracol = "registered_user, register_date, register_status, week_audited, month_audited, " _
                     & "reference, processor_name, audited_by, processing_date, country, " _
                     & "audit_date, item_reference, process_type, vendor_name, " _
                     & "audit_status, remarks, error_field, incorrect_value, " _
                     & "correct_value, corrected_by, correction_date, rca1, " _
                     & "rca2, rca3, rca4, rca5, preventive_solution, " _
                     & "responsible, appearance, root_cause, error, clean, total"
       ElseIf tblname = "DISPLAY:SQL_MNLDB.dbo.tblIsscError" Then
           paracol = "id, registered_user, register_date, register_status, week_audited, month_audited, " _
                     & "reference, processor_name, audited_by, processing_date, country, " _
                     & "audit_date, item_reference, process_type, vendor_name, " _
                     & "audit_status, remarks, error_field, incorrect_value, " _
                     & "correct_value, corrected_by, correction_date, rca1, " _
                     & "rca2, rca3, rca4, rca5, preventive_solution, " _
                     & "responsible, appearance, root_cause, error, clean, total"
       End If
End Function
Public Function param(colnum As Long) As String
    Select Case colnum
        Case 1: param = "id"
        Case 2: param = "registered_user"
        Case 3: param = "register_date"
        Case 4: param = "register_status"
        Case 5: param = "week_audited"
        Case 6: param = "month_audited"
        Case 7: param = "reference"
        Case 8: param = "processor_name"
        Case 9: param = "audited_by"
        Case 10: param = "processing_date"
        Case 11: param = "country"
        Case 12: param = "audit_date"
        Case 13: param = "item_reference"
        Case 14: param = "process_type"
        Case 15: param = "vendor_name"
        Case 16: param = "audit_status"
        Case 17: param = "remarks"
        Case 18: param = "error_field"
        Case 19: param = "incorrect_value"
        Case 20: param = "correct_value"
        Case 21: param = "corrected_by"
        Case 22: param = "correction_date"
        Case 23: param = "rca1"
        Case 24: param = "rca2"
        Case 25: param = "rca3"
        Case 26: param = "rca4"
        Case 27: param = "rca5"
        Case 28: param = "preventive_solution"
        Case 29: param = "responsible"
        Case 30: param = "appearance"
        Case 31: param = "root_cause"
        Case 32: param = "error"
        Case 33: param = "clean"
        Case 34: param = "total"
    End Select
End Function



