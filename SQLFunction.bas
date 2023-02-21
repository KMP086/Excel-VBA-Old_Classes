Attribute VB_Name = "SQLFunction"
Option Explicit
Function displayworkfile(serial As String, userposition As String, usercountry As String, outsheet As String, outcol As Integer, storedquery As String)
    
    
    Dim mobjConn As Connection
    Dim mobjCmd As Command
    Set mobjConn = New ADODB.Connection
    mobjConn.Open serial
    Set mobjCmd = New ADODB.Command
    With mobjCmd
        .ActiveConnection = mobjConn
        .CommandText = storedquery
        .CommandType = adCmdStoredProc
        .CommandTimeout = 300
        ' repeat as many times as you have parameters
        .Parameters.Append .CreateParameter("@UserPos", adVarChar, adParamInput, 100, userposition)
        .Parameters.Append .CreateParameter("@UserCountry", adVarChar, adParamInput, 100, usercountry)
        
    End With
    
    Dim rs As ADODB.Recordset
    Set rs = mobjCmd.Execute
    Dim x As Variant
    Dim col As Integer
    Dim filerow As Long
    If rs.EOF Then
        MsgBox "No records found!!" & rs.EOF
    Else
       ''Headers
       With ThisWorkbook.Worksheets(outsheet)
       filerow = .UsedRange.Rows.Count
       For Each x In rs.Fields
           .Range(Cells(filerow, outcol).Address).Offset(0, col).Value = x.Name
           col = col + 1
       Next x
       ''content/body
         .Range(Cells(2, outcol).Address).CopyFromRecordset rs
       End With
    End If
End Function

Function userdata(username As String, serial As String, dbofield As String) As String
    Dim mobjConn As Connection
    Dim mobjCmd As Command
    Set mobjConn = New ADODB.Connection
    mobjConn.Open serial
    Set mobjCmd = New ADODB.Command
    With mobjCmd
        .ActiveConnection = mobjConn
        .CommandText = "dbo.SOAUserIdentify"
        .CommandType = adCmdStoredProc
        ''.CommandTimeout = 300
        ' repeat as many times as you have parameters
        .Parameters.Append .CreateParameter("@SOAUser", adVarChar, adParamInput, 100, username)
    End With
    Dim rs As ADODB.Recordset
    Set rs = mobjCmd.Execute
    userdata = rs.Fields(dbofield)
End Function

Function alterdata(storedprocname As String, serial As String, recordid As Long, soarcvd As String, soadate As String, datedue As String, invdate As String, _
soavendorcodeinvoice As String, vendorcode As String, soavendorname As String, _
invoicenum As String, cwnum As String, account As String, reference As String, soaamount As String, _
cur As String, commento As String, disputereason As String, soaoperators As String, _
soavendorcode As String, branch As String, invoicedatecomparison As String, soavalidationamount As String, soacrr As String, _
soaremarks As String, soacleared As String, soaterms As String, soacwduedate As String, _
soacategories As String, soauserlead As String, soauserpos As String, soauser As String, _
dborecordstatus As String, soausercountry As String, _
 entity As String, addcategory As String, remarksnotify As String)
    Dim deci As New ADODB.Parameter
    Dim mobjConn As Connection
    Dim mobjCmd As Command
    Set mobjConn = New ADODB.Connection
    mobjConn.Open serial
    Set mobjCmd = New ADODB.Command
    With mobjCmd
        .ActiveConnection = mobjConn
        .CommandText = storedprocname
        .CommandType = adCmdStoredProc
        .CommandTimeout = 300
                
        ' repeat as many times as you have parameters
        
        If IsEmpty(recordid) = False And storedprocname = "dbo.SOAUpdateWorkFile" Then
            .Parameters.Append .CreateParameter("@ID", adVarChar, adParamInput, 100, recordid)
        End If
                
        .Parameters.Append .CreateParameter("@SOARcvd", adVarChar, adParamInput, 50, soarcvd)
        .Parameters.Append .CreateParameter("@SOADate", adVarChar, adParamInput, 50, soadate)
        .Parameters.Append .CreateParameter("@DueDate", adVarChar, adParamInput, 50, datedue)
        .Parameters.Append .CreateParameter("@InvDate", adVarChar, adParamInput, 50, invdate)
        .Parameters.Append .CreateParameter("@SOAVendorCodeInvoice ", adVarChar, adParamInput, 200, soavendorcodeinvoice)
        .Parameters.Append .CreateParameter("@VendorCode", adVarChar, adParamInput, 100, vendorcode)
        .Parameters.Append .CreateParameter("@SOAVendorName", adVarChar, adParamInput, 100, soavendorname)
        .Parameters.Append .CreateParameter("@InvoiceNum ", adVarChar, adParamInput, 100, invoicenum)
        .Parameters.Append .CreateParameter("@CWNum ", adVarChar, adParamInput, 100, cwnum)
        .Parameters.Append .CreateParameter("@Account", adVarChar, adParamInput, 100, account)
        .Parameters.Append .CreateParameter("@Reference", adVarChar, adParamInput, 100, reference)
        ''decimal
        .Parameters.Append .CreateParameter("@SOAAmount", adVarChar, adParamInput, 100, soaamount)
        
        .Parameters.Append .CreateParameter("@Currency", adChar, adParamInput, 10, cur)
        .Parameters.Append .CreateParameter("@Comments", adVarChar, adParamInput, 800, commento)
        .Parameters.Append .CreateParameter("@DisputeReason", adVarChar, adParamInput, 800, disputereason)
        .Parameters.Append .CreateParameter("@SOAOperators", adVarChar, adParamInput, 100, soaoperators)
        .Parameters.Append .CreateParameter("@SOAVendorCode", adVarChar, adParamInput, 100, soavendorcode)
        ''decimal
        .Parameters.Append .CreateParameter("@SOAValidationAmount", adVarChar, adParamInput, 100, soavalidationamount)
        
        .Parameters.Append .CreateParameter("@SOACrr", adChar, adParamInput, 10, soacrr)
        .Parameters.Append .CreateParameter("@SOARemarks", adVarChar, adParamInput, 800, soaremarks)
        .Parameters.Append .CreateParameter("@SOACleared", adVarChar, adParamInput, 100, soacleared)
        .Parameters.Append .CreateParameter("@SOATerms", adVarChar, adParamInput, 100, soaterms)
        .Parameters.Append .CreateParameter("@SOACWDueDate", adVarChar, adParamInput, 50, soacwduedate)
        .Parameters.Append .CreateParameter("@SOACategories", adVarChar, adParamInput, 100, soacategories)
        
        .Parameters.Append .CreateParameter("@SOAUserLead", adVarChar, adParamInput, 100, soauserlead)
        .Parameters.Append .CreateParameter("@SOAUserPos", adVarChar, adParamInput, 100, soauserpos)
        .Parameters.Append .CreateParameter("@SOAUser", adVarChar, adParamInput, 100, soauser)
        .Parameters.Append .CreateParameter("@DBORecordStatus", adVarChar, adParamInput, 200, dborecordstatus)
        .Parameters.Append .CreateParameter("@InvoiceDateComparison", adVarChar, adParamInput, 100, invoicedatecomparison)
        .Parameters.Append .CreateParameter("@SOAUserCountry", adVarChar, adParamInput, 150, soausercountry)
        .Parameters.Append .CreateParameter("@Branch", adVarChar, adParamInput, 50, branch)
        .Parameters.Append .CreateParameter("@Entity", adVarChar, adParamInput, 50, entity)
        .Parameters.Append .CreateParameter("@AddCategory", adVarChar, adParamInput, 500, addcategory)
        .Parameters.Append .CreateParameter("@RemarksNotify", adVarChar, adParamInput, 8000, remarksnotify)
        .Execute
    End With
        
End Function


