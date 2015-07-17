Attribute VB_Name = "modUDP"
Option Explicit

Public Sub SendUDP(Data As String, Optional toIP As String = "")
 
    'For i = LBound(UDPSend) To UBound(UDPSend)
    If toIP = "" Then toIP = SubNetMask
        
    frmMain.UDPSend.remoteHost = toIP

    AddToLog toIP & " " & Data, False
    Hash Data

    On Error GoTo NonBlockingSocketsError

    frmMain.UDPSend.SendData Data

    Exit Sub
NonBlockingSocketsError:
    If Err.Number = 10035 Then DoEvents
    If Err.Number = 10004 Then Exit Sub
    Resume
    
    'Next i
End Sub

Public Sub SayHello()
    frmMain.lblTotalUsers.Caption = 1
    frmMain.lblTotalFiles.Caption = frmMain.lblShareFiles.Tag
    frmMain.lblTotalFolders.Caption = frmMain.lblShareFolders.Tag
    frmMain.lblTotalSize.Caption = ChangeByte(frmMain.lblShareSize.Tag)
    frmMain.lblTotalSize.Tag = frmMain.lblShareSize.Tag
    'SendUDP "HELLO|" & GetIP(), SubNetMask
    SendUDP "HELLO|", SubNetMask
End Sub
