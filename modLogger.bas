Attribute VB_Name = "modLogger"
Private LogFileNumber As Long
Public LogOn As Boolean

Public Sub AddToLog(message As String, msgIn As Boolean)

Dim Data As String

If LogOn Then
    
    If LogFileNumber = 0 And frmMain.bClosing <> 2 Then
        LogFileNumber = FreeFile
        Open App.path & "\log.txt" For Append As LogFileNumber
    End If
    
    If msgIn Then
        Data = Date & " " & time & " < " & message & vbCrLf
    Else
        Data = Date & " " & time & " > " & message & vbCrLf
    End If
    'txtLog.text = txtLog.text & Data
    'txtLog.SelStart = Len(txtLog.text)
    Print #LogFileNumber, Data;
    Close #LogFileNumber
    Open App.path & "\log.txt" For Append As LogFileNumber
    
End If

End Sub

Public Sub CloseLogFile()
    Close LogFileNumber
End Sub

'Records a record of the last IP to send data to me
Public Sub setLastIP(IP As String)
    Dim IpFile As Long
    Dim IPNumber As String * 16
    IPNumber = String(16, " ")
    IPNumber = IP
    IpFile = FreeFile
    
    Open App.path & "\lastIP.txt" For Output As IpFile
    Print #IpFile, IP;
    Close #IpFile
    
End Sub

'Gets the last IP sent to me
Public Function getLastIP() As String
    Dim IpFile As Long
    Dim IPNumber As String * 16
    
    IPNumber = String(16, " ")
    IpFile = FreeFile
    On Error Resume Next
    
    Open App.path & "\lastIP.txt" For Input As IpFile
    
    Input #IpFile, IPNumber
    On Error GoTo 0
    
    Close #IpFile
    getLastIP = Trim(IPNumber)
End Function

Public Function isBadIP(IP As String) As Boolean
    Set rs = conn.Execute("SELECT COUNT(*) AS bans FROM tblBanned WHERE bIP = '" & IP & "'")
    isBadIP = (rs("bans") > 0)
    Set rs = Nothing
End Function
