VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQLInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public access As New clsServerConnect
Public arry As New clsArrayItems
Public arritems As New clsArrayItems
Public tblpara As New clsItemList
Public sqlConn As Connection
Public sqlCmd As Command
Public rs As ADODB.Recordset
Public Function insertall(shtname As String, tblname As String) As Boolean
   Dim sqlseq As String
   Dim rowrng As Long
   ''Dim arritem() As Variant
   Dim RowsCollection As New Collection
   Dim item As Range
   Dim intStartIndex As Long
   Const intBatchSize As Long = 1000 ' Number of items to store/process in each batch
   Dim intBatchIndex As Long ' Index for the batch variables array
   Dim intEndIndex As Long ' Number of items to store/process in each batch
   Dim strValues As String
   

   On Error GoTo skipthis
   
   If arry.insertarray(shtname, tblname) = False Then Debug.Print ("Error Sql Sequence!!")
   
   With ThisWorkbook.Worksheets("RAM")
        If .Visible = False Then .Visible = True
        rowrng = .Cells(Rows.Count, 1).End(xlUp).Row
                
        For Each item In .Range("A1:A" & rowrng)
               RowsCollection.Add CStr(item)
               ''arritem(item.row) = CStr(item)
        Next item
                   
   End With
   ''///////////////////////////////////////////////////////////////////////////////////////////////
    intTotalItems = RowsCollection.Count
    intStartIndex = 1 ' Starting index of the batch
    Dim batchVariables(1 To intBatchSize) As Variant
    
    Do While intStartIndex <= intTotalItems
        For intBatchIndex = 1 To intBatchSize
            batchVariables(intBatchIndex) = Null
        Next intBatchIndex
        strValues = ""

        intEndIndex = intStartIndex + intBatchSize - 1 ' Ending index of the batch
        If intEndIndex > intTotalItems Then
            intEndIndex = intTotalItems ' Adjust the intEndIndex if it exceeds the total items
        End If

        For intBatchIndex = 1 To intEndIndex - intStartIndex + 1
            batchVariables(intBatchIndex) = RowsCollection(intStartIndex + intBatchIndex - 1)
        Next intBatchIndex
        
        For Each CurrentRow In batchVariables
            strRowData = CurrentRow
            strValues = strValues & strRowData
        Next CurrentRow
        
        strValues = Left(strValues, Len(strValues) - 1)
        strValues = strValues & ";"
        Debug.Print ("insert into " & tblname & "(" & CStr(tblpara.paracol(tblname)) & ") Values" & strValues)
        ''sql////////////////////////////////////////////////////////////////////////////////////////////
        Set sqlConn = New ADODB.Connection
        sqlConn.Open access.passcode
        sqlConn.Execute "insert into " & tblname & "(" & CStr(tblpara.paracol(tblname)) & ") Values" & strValues
        sqlConn.Close
                
        intStartIndex = intStartIndex + intBatchSize
    Loop
   
   
   MsgBox ("Data is Uploaded")
   ''/////////////////////////////////////////////////////
   If ThisWorkbook.Worksheets("RAM").Visible = True Then ThisWorkbook.Worksheets("RAM").Visible = False
   insertall = True
   Exit Function
skipthis:
   If ThisWorkbook.Worksheets("RAM").Visible = True Then ThisWorkbook.Worksheets("RAM").Visible = False
   insertall = False
End Function

