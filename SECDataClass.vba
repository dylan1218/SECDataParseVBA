Option Explicit
Public baseUrlString As String
Public downloadString As String

'Note this class can be utilized to download tab deliminated SEC data. There is one common input for methods herein which is the EndPoint variable.
'This endpoint variable represents you quarter and year you want to download in the format of Yearq#. Example: 2018q3, note this is case sensitivite.
'The other class is uitilized for querying the data downloaded here
Private Function httpGet(myUrl As String)
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myUrl, False
    WinHttpReq.send
    httpGet = WinHttpReq.Status
End Function

Private Function vDownloadString(EndPoint As String)
    baseUrlString = "https://www.sec.gov/files/dera/data/financial-statement-data-sets/"
    downloadString = baseUrlString & EndPoint & ".zip"
    vDownloadString = downloadString
End Function
Private Function fileStatus(EndPoint As String)
    fileStatus = (httpGet(vDownloadString(EndPoint)))
End Function
Public Function dirExists(EndPoint As String) As Boolean
    Dim s_directory As Variant
    
    s_directory = "C:\SECVba\" & EndPoint
    
    Dim OFSO As Object
    Set OFSO = CreateObject("Scripting.FileSystemObject")
    dirExists = OFSO.FolderExists(s_directory)
End Function

Public Function unzip(EndPoint As String)
    Dim FSO As Object
    Dim oApp As Object
    Dim Fname As Variant
    Dim FileNameFolder As Variant
    Dim unzipPath As String

    Fname = "C:\SECVba\2018q3.zip"

    unzipPath = "C:\SECVba\"
    'Create the folder name
    FileNameFolder = unzipPath & EndPoint
    
    If dirExists(FileNameFolder & "\") = False Then
        'Make the normal folder in DefPath
        MkDir FileNameFolder
    
        'Extract the files into the newly created folder
        Set oApp = CreateObject("Shell.Application")
        oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname).items
    
        On Error Resume Next
        Set FSO = CreateObject("scripting.filesystemobject")
        FSO.deletefolder Environ("Temp") & "\Temporary Directory*", True
    End If
End Function

Public Function httpDownLoad(EndPoint As String)
    If Me.fileExists(EndPoint) = False Then
        If fileStatus(EndPoint) = 200 Then
            Dim WinHttpReq As Object
            Dim oStream As Object
            Dim myUrl As String
            
            Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
            Set oStream = CreateObject("ADODB.Stream")
                
            WinHttpReq.Open "GET", vDownloadString(EndPoint), False
            WinHttpReq.send
            
            oStream.Open
                oStream.Type = 1
                oStream.Write WinHttpReq.responseBody
                oStream.SaveToFile "C:\SECVba\" & EndPoint & ".zip", 1 ' 1 = no overwrite, 2 = overwrite
            oStream.Close
            If Me.dirExists(EndPoint) Then
                Call unzip(EndPoint)
            End If
        End If
    Else
        If Me.dirExists(EndPoint) Then
            Call unzip(EndPoint)
        End If
    End If
End Function

Public Function fileExists(EndPoint)
    
    Dim FilePath As String
    FilePath = Dir("C:\SECVba\" & EndPoint & ".zip")
    
    If FilePath = "" Then
        fileExists = False
    Else
        fileExists = True
    End If
    
End Function
