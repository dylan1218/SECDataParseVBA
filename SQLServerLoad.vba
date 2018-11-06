Option Explicit
Public SECfileNames As Variant
Public SECfileNameHeaders As Variant
Public ConnectionString As String
Public Function vSQLServerConnectionString() As String
    ConnectionString = "Provider=MSOLEDBSQL;" _
             & "Server=(local);" _
             & "Database=master;" _
             & "Integrated Security=SSPI;" _
             & "DataTypeCompatibility=80;"
    vSQLServerConnectionString = ConnectionString
End Function

Public Function vSECfileArray() As Variant
    SECfileNames = Array("num", "pre", "sub", "tag")
    vSECfileArray = SECfileNames
End Function

Public Function vSECfileNameHeadersArray() As Variant
    SECfileNameHeaders = Array(Array("adsh", "tag", "version", "coreg", "ddate", "qtrs", "uom", "value") _
    , Array("adsh", "report", "line", "stmt", "inpth", "rfile", "tag", "version", "plabel") _
    , Array("adsh", "cik", "name", "sic", "countryba", "stprba", "cityba", "zipba", "bas1", "bas2", "baph", "countryma", "stprma", "cityma", "zipma", "mas1", "mas2", "countryinc", "stprinc", "ein", "former", "changed", "afs", "wksi", "fye", "form", "period", "fy", "fp", "filed", "accepted", "prevrpt", "detail", "instance", "nciks", "aciks") _
    , Array("tag", "version", "custom", "abstract", "datatype", "iord", "crdr", "tlabel", "foc"))
    vSECfileNameHeadersArray = SECfileNameHeaders
End Function

Public Function vTableCreateString(vSECfileNameHeadersArray As Variant, vSECfileArrayName As Variant)
    
    Dim tableHeader As Variant
    Dim stringAppend As String
    Dim Counter As Integer
    Dim stringStart As String
    Dim indexInArray As Integer
    
    indexInArray = Application.Match(vSECfileArrayName, Me.vSECfileArray, False) - 1
        
    stringStart = "CREATE TABLE " & Me.vSECfileArray(indexInArray) & "("
    Counter = 0
    For Each tableHeader In vSECfileNameHeadersArray(indexInArray)
        stringAppend = stringAppend & vSECfileNameHeadersArray(indexInArray)(Counter) & " " & "VARCHAR(40)" & " " & "NOT NULL" & ", "
        Counter = Counter + 1
    Next tableHeader
    
    vTableCreateString = stringStart & stringAppend & ")"

End Function

Public Function dropSECTables(SECfileArray As Variant)
    Dim conn As Variant
    Dim rst As Variant
    Dim x As Integer
    
    Set conn = CreateObject("ADODB.Connection")
    Set rst = CreateObject("ADODB.Recordset")
    
    conn.ConnectionString = vSQLServerConnectionString
    
    conn.Open
    
    For x = 0 To 3
        On Error GoTo loopContinue
        conn.Execute "DROP TABLE " & Me.vSECfileArray(x)
loopContinue:
Resume loopContinue2
loopContinue2:
        x = x + 1
    Next x
    conn.Close
End Function

Public Function createSECTables(vSECfileNameHeadersArray As Variant, vSECfileArrayName As Variant)
    'Note this only needs to be ran once
    
    Dim conn As Variant
    Dim rst As Variant
    
    Set conn = CreateObject("ADODB.Connection")
    Set rst = CreateObject("ADODB.Recordset")
    
    conn.ConnectionString = vSQLServerConnectionString
    
    conn.Open
        conn.Execute Me.vTableCreateString(Me.vSECfileNameHeadersArray, vSECfileArrayName)
    conn.Close
    
End Function

Public Function addDataToTables(vSECfileName As Variant, filePath As String)
    
    Dim conn As Variant
    Dim rst As Variant
    Dim strSQL As String
    Dim indexInArray As Integer
    
    'filePath = "C:\SECVba\2018q3\num.txt"
    
    Set conn = CreateObject("ADODB.Connection")
    Set rst = CreateObject("ADODB.Recordset")
    
    indexInArray = Application.Match(vSECfileName, Me.vSECfileArray, False) - 1
    
    conn.ConnectionString = Me.vSQLServerConnectionString
    
    conn.Open
        conn.CommandTimeout = 0
        strSQL = "BULK " & _
        "INSERT " & Me.vSECfileArray(indexInArray) & _
        " FROM " & Chr(39) & filePath & Chr(39) & _
        " WITH ( FIELDTERMINATOR = '\t', ROWTERMINATOR = '0x0a')"
    
    conn.Execute strSQL
    conn.Close
    
End Function

