VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQLDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public access As New clsServerConnect
Public disquery As New clsDisplayQuery
Public arrwbk As New clsArrWbk
Public Function sqldisplay(wksht As String, rsrow As Long, rowpos As Long, colpos As Long) As Boolean
    Dim viewitem As Integer, delview As Integer
    Dim monthfrom As String, monthto As String, yearfrom As String, yearto As String, daysel As String, sqlquery As String
    Dim usedrow As Long, col As Long
    If Onerror Then GoTo skipthis
        If ThisWorkbook.Worksheets("Internal Audit Tracker").Cells(2, 5) <> "" Then
            If arrwbk.clearsht(ThisWorkbook.Name, "Internal Audit Tracker") = False Then Debug.Print ("Error Clear")
        End If
   
        
        Set sqlConn = New ADODB.Connection
        sqlConn.Open access.passcode
        Debug.Print (sqlConn.State)
        
        sqlquery = disquery.displayquery
        Debug.Print (sqlquery)
        Set rs = sqlConn.Execute(sqlquery)
        Debug.Print rs.RecordCount
            With ThisWorkbook.Worksheets(wksht)
             ' COLUMNS
                col = 0
                For Each x In rs.Fields
                    .Range(Cells(rsrow, colpos).Address).Offset(0, col).Value = x.Name
                    col = col + 1
                Next x
             ' DATA
                .Range(Cells(rowpos, colpos).Address).CopyFromRecordset rs
            End With
        rs.Close
        sqlConn.Close
        rfpsqldata = True
        Exit Function
skipthis:
        rfpsqldata = False
End Function



