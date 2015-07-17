Attribute VB_Name = "modFileSearch"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub GetFolders(path As String, rootID As Long)
    'Dim folders() As String
   
    Dim fso As New FileSystemObject
    Dim fld As folder
    Dim folderID As Long
    Dim folderName As String
    
    On Error GoTo getFolderError
    Set fld = fso.GetFolder(path)
    
    'ReDim folders(fld.SubFolders.Count - 1)
    
    For Each fld In fld.SubFolders
        'folders(i) = fld.Name
        folderName = fld.Name
        
        folderName = Replace(folderName, "'", "''")
        
        Set rs = conn.Execute("SELECT folderID FROM tblFolders WHERE folderName = '" & folderName & "' AND rootID=" & rootID)
        'Check if folder is in database
        If rs.EOF Or rs.BOF Then
            'Create Share
            Set rs = conn.Execute("INSERT INTO tblFolders (rootID, folderName, lastUpdate) VALUES (" & rootID & ",'" & folderName & "', NOW())")
            Set rs = conn.Execute("SELECT folderID FROM tblFolders WHERE folderName = '" & folderName & "' AND rootID=" & rootID)
            folderID = rs("folderID")
            rs.Close
        Else
            folderID = rs("folderID")
            rs.Close
            'Update Share
            Set rs = conn.Execute("UPDATE tblFolders SET lastUpdate=NOW() WHERE folderID=" & folderID)
        End If
        
        GetFiles fld.path, folderID
        GetFolders fld.path, folderID
        frmMain.StatusChange "Scanning '" & fld.path & "' for new files"
        frmMain.status.Refresh
        
        Set rs = Nothing
    Next
    On Error GoTo 0
getFolderError:

End Sub

Public Sub GetFiles(ByVal path As String, rootID As Long)
   
    Dim fso As New FileSystemObject
    Dim fld As folder
    Dim fil As File
    Dim filename As String
    Dim fileID As Long
    
    'Path = Replace(Path, "''", "'")
    Set fld = fso.GetFolder(path)
    
    On Error GoTo GetFilesError
    For Each fil In fld.Files
        filename = fil.Name
        filename = Replace(filename, "'", "''")
        
        'Checks if the file is of a banned type
        If FileTypeToImage(filename) <> 10 Then
        
            Set rs = conn.Execute("SELECT fileID FROM tblFiles WHERE fileName = '" & filename & "' AND folderID=" & rootID)
            'Check if folder is in database
            If rs.EOF Or rs.BOF Then
                'Create Share
                Set rs = conn.Execute("INSERT INTO tblFiles (folderID, fileName, fileSize, lastUpdate) VALUES (" & rootID & ",'" & filename & "'," & fil.size & ", NOW())")
                Set rs = conn.Execute("SELECT fileID FROM tblFiles WHERE fileName = '" & filename & "' AND folderID=" & rootID)
                fileID = rs("fileID")
                rs.Close
            Else
                fileID = rs("fileID")
                rs.Close
                'Update Share
                Set rs = conn.Execute("UPDATE tblFiles SET fileSize=" & fil.size & ", lastUpdate=NOW() WHERE fileID=" & fileID)
            End If
            'frmMain.lblStatus.Caption = Path & "\" & fileName
            'frmMain.lblStatus.Refresh
           DoEvents
           Set rs = Nothing
        End If
    Next
    
GetFilesError:
    Exit Sub
    
End Sub

Public Function GetPath(ByVal folderID As Long) As String
Dim rs2 As ADODB.Recordset
Dim folder As String

    Set rs2 = conn.Execute("SELECT rootID, folderName FROM tblFolders WHERE folderID=" & folderID)
    Do While rs2("rootID") <> -1
        folder = rs2("folderName") & "\" & folder
        folderID = rs2("rootID")
        rs2.Close
        Set rs2 = conn.Execute("SELECT rootID, folderName FROM tblFolders WHERE folderID=" & folderID)
    Loop
    folder = rs2("folderName") & "\" & folder
    rs2.Close
    Set rs2 = Nothing
GetPath = folder
End Function

Function FileExists(filename As String) As Boolean
    Dim fileNumber As Long
    On Error GoTo FileExistError
    fileNumber = FreeFile(1)
    Open filename For Input Shared As fileNumber
    Close fileNumber
    
    FileExists = True
    Exit Function
    
FileExistError:
    If Err.number = 55 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close fileNumber
End Function

Function FolderExists(folder As String) As Boolean
    'Add code later
    Dim fso As New FileSystemObject
    Dim fld As folder
    On Error GoTo FolderExistsError
    Set fld = fso.GetFolder(folder)
    
    FolderExists = True
    Exit Function
FolderExistsError:
    FolderExists = False
End Function

Function getSpeed(rank As Long) As String
    Select Case rank
        Case Is = 1, 2: getSpeed = "Fastest"
        Case Is = 3, 4, 5: getSpeed = "Fast"
        Case Is = 6, 7, 8, 9: getSpeed = "Normal"
        Case Is = 10, 11, 12: getSpeed = "Slow"
        Case Else: getSpeed = "Slowest"
    End Select
End Function
