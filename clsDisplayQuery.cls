VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDisplayQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public itemlist As New clsItemList
Public Function displayquery() As String

        With ThisWorkbook.Worksheets("Item list")
            monthfrom = CStr(itemlist.monthlistint(.Range("D11").Value))
            monthto = CStr(itemlist.monthlistint(.Range("D12").Value))
            ''//////////////////////////////////////////////////////////
            yearfrom = CStr(.Range("D9").Value)
            yearto = CStr(.Range("D10").Value)
            ''///////////////////////////////////////////////////////////
            delview = itemlist.sqllogic(.Range("D15").Value)
            viewitem = itemlist.sqllogic(.Range("D14").Value)
            daysel = CStr(.Range("D13").Value)
        End With
                
        
        If monthfrom = "" Or monthto = "" Then
           MsgBox "Empty month selection!!"
           Exit Function
        ElseIf CInt(monthfrom) > CInt(monthto) Then
           MsgBox "Incorrect month selection!!"
           End
        End If
                
        If yearfrom = "" Or yearto = "" Then
           MsgBox "Empty year selection!!"
           End
        ElseIf CInt(yearfrom) > CInt(yearto) Then
           MsgBox "Incorrect year selection!!"
           End
        End If
        
        
        If daysel <> "All" And delview = 1 Then
           If viewitem = 1 Then
                displayquery = "SELECT " & CStr(itemlist.paracol("DISPLAY:SQL_MNLDB.dbo.tblIsscError")) & " From SQL_MNLDB.dbo.tblIsscError Where " _
                           & "(Month(register_date) between " & monthfrom & " and " & monthto & ") AND " _
                           & "(Year(register_date) between " & yearfrom & " AND " & yearto & ") AND " _
                           & "(Day(register_date) = " & daysel & ") AND registered_user = '" & Environ("USERNAME") & "') AND register_status = 'delete' " _
                           & "Order by id ASC;"
           ElseIf viewitem = 0 Then
                displayquery = "SELECT " & CStr(itemlist.paracol("DISPLAY:SQL_MNLDB.dbo.tblIsscError")) & " From SQL_MNLDB.dbo.tblIsscError Where " _
                           & "(Month(register_date) between " & monthfrom & " and " & monthto & ") AND " _
                           & "(Year(register_date) between " & yearfrom & " AND " & yearto & ") AND " _
                           & "(Day(register_date) = " & daysel & ")AND register_status = 'delete' " _
                           & "Order by id ASC;"
           End If
        
        ElseIf daysel = "All" And delview = 1 Then
           If viewitem = 1 Then
                displayquery = "SELECT " & CStr(itemlist.paracol("DISPLAY:SQL_MNLDB.dbo.tblIsscError")) & " From SQL_MNLDB.dbo.tblIsscError Where " _
                           & "(Month(register_date) between " & monthfrom & " and " & monthto & ") AND " _
                           & "(Year(register_date) between " & yearfrom & " AND " & yearto & ") AND " _
                           & "(registered_user = '" & Environ("USERNAME") & "') AND register_status = 'delete' " _
                           & "Order by id ASC;"
           ElseIf viewitem = 0 Then
                displayquery = "SELECT " & CStr(itemlist.paracol("DISPLAY:SQL_MNLDB.dbo.tblIsscError")) & " From SQL_MNLDB.dbo.tblIsscError Where " _
                           & "(Month(register_date) between " & monthfrom & " and " & monthto & ") AND " _
                           & "(Year(register_date) between " & yearfrom & " AND " & yearto & ")  AND register_status = 'delete' " _
                           & "Order by id ASC;"
           End If
       
       ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                
        ElseIf daysel = "All" And delview = 0 Then
           If viewitem = 1 Then
                displayquery = "SELECT " & CStr(itemlist.paracol("DISPLAY:SQL_MNLDB.dbo.tblIsscError")) & " From SQL_MNLDB.dbo.tblIsscError Where " _
                           & "(Month(register_date) between " & monthfrom & " and " & monthto & ") AND " _
                           & "(Year(register_date) between " & yearfrom & " AND " & yearto & ") AND " _
                           & "(Day(register_date) = " & daysel & ") AND (registered_user = '" & Environ("USERNAME") & "') AND register_status <> 'delete' " _
                           & "Order by id ASC;"
           ElseIf viewitem = 0 Then
                displayquery = "SELECT " & CStr(itemlist.paracol("DISPLAY:SQL_MNLDB.dbo.tblIsscError")) & " From SQL_MNLDB.dbo.tblIsscError Where " _
                           & "(Month(register_date) between " & monthfrom & " and " & monthto & ") AND " _
                           & "(Year(register_date) between " & yearfrom & " AND " & yearto & ") AND (Day(register_date) <> " & daysel & ") AND register_status = 'delete' " _
                           & "Order by id ASC;"
           End If
        
        ElseIf daysel <> "All" And delview = 0 Then
           If viewitem = 1 Then
                displayquery = "SELECT " & CStr(itemlist.paracol("DISPLAY:SQL_MNLDB.dbo.tblIsscError")) & " From SQL_MNLDB.dbo.tblIsscError Where " _
                           & "(Month(register_date) between " & monthfrom & " and " & monthto & ") AND " _
                           & "(Year(register_date) between " & yearfrom & " AND " & yearto & ") AND " _
                           & "(registered_user = '" & Environ("USERNAME") & "') AND register_status <> 'delete' " _
                           & "Order by id ASC;"
           ElseIf viewitem = 0 Then
                displayquery = "SELECT " & CStr(itemlist.paracol("DISPLAY:SQL_MNLDB.dbo.tblIsscError")) & " From SQL_MNLDB.dbo.tblIsscError Where " _
                           & "(Month(register_date) between " & monthfrom & " and " & monthto & ") AND " _
                           & "(Year(register_date) between " & yearfrom & " AND " & yearto & ") AND register_status <> 'delete' " _
                           & "Order by id ASC;"
           End If
        
        End If
        
End Function


