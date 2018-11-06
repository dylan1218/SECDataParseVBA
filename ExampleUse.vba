Sub testingDownloadClass()
'This is an example of how to utilize the class. This sub downloads Q3'2018 SEC XBRL tab deliminated data to the C:\SECVba directory
'note that we've added conditions herein for user notification purposes. But they are not required as the class already handles
'consideration of already existing data

  Dim downloadData As New SECDataClass
    
    If downloadData.fileExists("2018q3") = True Then
        MsgBox ("Data already downloaded ... checking if unzipped")
        If downloadData.dirExists("2018q3") = False Then
            GoTo downLoadandUnzip
        Else
            MsgBox ("Data also unzipped ... exiting sub")
            Exit Sub
        End If
    End If
    
downLoadandUnzip:
    downloadData.httpDownLoad ("2018q3")
    
    If downloadData.fileExists("2018q3") = True Then
        MsgBox ("File sucessfully downloaded")
    Else
        MsgBox ("File not sucessfully downloaded")
    End If
    
End Sub

'The below sub loads all of the SEC data tables to SQLserver 
Sub LoadNumTable()
    Dim TestString As New SQLServerLoad
    Dim TableName As String
    For Each TableName in TestString.vSECfileArray
        Call TestString.createSECTables(TestString.vSECfileNameHeadersArray, TableName)
        'note the loading of these tables will only need to occur once
    Next TableName
End Sub

'the four subs below load data into the respective tables. Method arguments are (tablename, filepath_where_tableData_is_located)
Sub LoadNUM()
    Dim testString As New SQLServerLoad
    Call testString.addDataToTables("num", "C:\SECVba\2018q3\num.txt")
End Sub

Sub LoadPRE()
    Dim testString As New SQLServerLoad
    Call testString.addDataToTables("pre", "C:\SECVba\2018q3\pre.txt")
End Sub

Sub LoadSUB()
    Dim testString As New SQLServerLoad
    Call testString.addDataToTables("sub", "C:\SECVba\2018q3\sub.txt")
End Sub

Sub LoadTAG()
    Dim testString As New SQLServerLoad
    Call testString.addDataToTables("tag", "C:\SECVba\2018q3\tag.txt")
End Sub



Public Function SQLQueryData(SQLQuery As String) As String
  'This function calls an array of queried data to a range
  'will look to turn query functions/methods into a separate class
    Dim GetConnStringMethod As New SQLServerLoad
    
    Dim conn As Variant
    Dim rst As Variant
    Dim Rng As Range
    
    Set Rng = Range(ActiveCell, ActiveCell)
    
    Set conn = CreateObject("ADODB.Connection")
    Set rst = CreateObject("ADODB.Recordset")
    
    conn.ConnectionString = GetConnStringMethod.vSQLServerConnectionString
    
    conn.Open
    With rst
        .ActiveConnection = conn
        .Open SQLQuery
        ActiveSheet.Range("A5").CopyFromRecordset rst
        MsgBox (Rng.Address)
        .Close
    End With
        
    
    conn.Close
    
    
End Function

Sub TestSub()
    'example SQLQueryData function call
    Call SQLQueryData("SELECT adsh, ddate, value FROM num WHERE tag='currentassets' AND ddate='20180331'")

End Sub

