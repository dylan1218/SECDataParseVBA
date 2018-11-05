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
