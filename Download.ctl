VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl Download 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   600
      Top             =   840
   End
   Begin VB.Timer tmrSpeed 
      Interval        =   2000
      Left            =   120
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock TCPReceive 
      Left            =   120
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Download"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "Download"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Const SavePath = "c:\" 'Not needed anymore

Private mvarFilename As String 'local copy
Private mvarFilePath As String 'local copy
Private mvarSize As Long 'local copy
Private mvarPosition As Long 'local copy
Private mvarFileNumber As Long 'local copy
Private mvarRemoteHost As String 'local copy
Private mvarTransferID As Long 'local copy
Private mvarStatus As Long 'local copy
Private mvarRemotePort As Long 'local copy
Private mvarlastTime As Date 'local copy
Private mvarlastPosition As Long 'local copy
Private mvarSpeed As Long 'local copy

Public Event Finished()
Public Event Update()

Private Property Let lastPosition(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lastPosition = 5
    mvarlastPosition = vData
End Property

Private Property Get lastPosition() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lastPosition
    lastPosition = mvarlastPosition
End Property

Private Property Let lastTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lastTime = 5
    mvarlastTime = vData
End Property

Private Property Get lastTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lastTime
    lastTime = mvarlastTime
End Property

Public Property Let status(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.status = 5
    mvarStatus = vData
    If mvarTransferID <> -1 Then
        Select Case vData
            Case Is = 1: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "Pending"
            Case Is = 2: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "Timed Out"
            Case Is = 3: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "Refused"
            Case Is = 4: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "Complete"
            Case Is = 5: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "Sending"
            Case Is = 6: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "Receiving"
            Case Is = 7: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "Remote Error"
            Case Is = 8: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "File Not Found"
            Case Is = 9: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "Aborted"
            Case Is = 10: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "Too Busy"
            Case Is = 11: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "Disk Full"
            Case Is = 12: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "File Has Been Deleted"
            Case Else: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "Unknown"
        End Select
        Select Case vData
            Case Is = 6
                'Do Nothing much
            Case Else
                frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(3).Text = ChangeByte(0, True) & "/S"
                frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(3).Tag = 0
        End Select
    End If
End Property


Public Property Get status() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.status
    status = mvarStatus
End Property

Public Function getProgress() As Single
If mvarSize = 0 Then
    getProgress = 0
Else
    getProgress = mvarPosition / mvarSize * 100
End If
End Function

Public Sub Start(remoteHost As String, remotePort As Long, filename As String, position As Long, size As Long, Optional bad As Boolean = False)

    mvarFilename = filename
    
    Dim item As ListItem
    Dim index As Long
    index = frmMain.lstTransfers.Tag
    frmMain.lstTransfers.Tag = frmMain.lstTransfers.Tag + 1
    
    Set item = frmMain.lstTransfers.ListItems.Add(, "I" & index, mvarFilename, 2, 2)
    item.ListSubItems.Add , , ChangeByte(CLng(size), True)
    item.ListSubItems.Add , , "Pending"
    item.ListSubItems.Add , , "0 b/S"
    mvarTransferID = index
    Set item = Nothing
    
    If Not bad Then
        Let status = 1
        
        mvarSize = size
        
        'Convert Position to a multiply of BufferSize
        mvarPosition = (position \ bufferSize) * bufferSize
        
        mvarRemoteHost = remoteHost
        mvarRemotePort = remotePort
        
        TCPReceive.remotePort = remotePort
        TCPReceive.remoteHost = remoteHost
        TCPReceive.Connect
    End If
    
End Sub

Public Property Let transferID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.transferID = 5
    mvarTransferID = vData
End Property

Public Property Get transferID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.transferID
    transferID = mvarTransferID
End Property

Public Property Let remoteHost(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.remoteHost = 5
    mvarRemoteHost = vData
End Property

Public Property Get remoteHost() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.remoteHost
    remoteHost = mvarRemoteHost
End Property

Private Property Let fileNumber(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fileNumber = 5
    mvarFileNumber = vData
End Property

Private Property Get fileNumber() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fileNumber
    fileNumber = mvarFileNumber
End Property

Public Property Let position(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.position = 5
    mvarPosition = vData
End Property

Public Property Get position() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.position
    position = mvarPosition
End Property

Public Property Let size(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.size = 5
    mvarSize = vData
End Property

Public Property Get size() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.size
    size = mvarSize
End Property

Public Property Let filename(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.filename = 5
    mvarFilename = vData
End Property

Public Property Get filename() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.filename
    filename = mvarFilename
End Property

Private Sub TCPReceive_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal Helpfile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Select Case number
        Case 10060: Let status = 2 'Connection Timeout
        Case 10061: Let status = 3 'Connection Refused
        Case Else: msg "Unknown Error: Download " & number & " " & Description: status = 99
    End Select
End Sub

Private Sub tmrSpeed_Timer()
    
    If TCPReceive.State = sckConnected Then

        mvarSpeed = getSpeed
        
        mvarlastPosition = mvarPosition
        mvarlastTime = Now()
        frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(3).Tag = mvarSpeed
        frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(3).Text = ChangeByte(CSng(mvarSpeed), True, 0) & "/S"
        'Diplays Position
        If frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text <> "Complete" Then
            frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = Round(Me.getProgress, 0) & "%"
        End If
    Else
        mvarSpeed = 0
        tmrSpeed.Enabled = False
    End If
    RaiseEvent Update
End Sub

Public Function getSpeed() As Long
    'Displays Speed
    Dim difference As Long
    Dim speed As Long
    difference = seconds(CSng(Now() - mvarlastTime))
    If difference = 0 Then
        speed = 0
    Else
        speed = (mvarPosition - mvarlastPosition) / difference
    End If
    
    getSpeed = speed
End Function

Public Function TimeLeft() As String

    Dim bytesLeft As Long
    bytesLeft = ((100 - Me.getProgress) / 100) * Me.size
    If mvarSpeed <> 0 Then
        TimeLeft = ChangeSecond(bytesLeft / mvarSpeed)
    Else
        TimeLeft = "Never"
    End If
End Function

Private Sub UserControl_Initialize()
    mvarTransferID = -1
End Sub

Private Sub UserControl_Resize()
    Width = lblName.Width
    Height = lblName.Height
End Sub

Private Sub TCPReceive_Close()
    On Error GoTo TCPReceiveError
    If FileLen(frmMain.txtDownloadLocation.Text & mvarFilename & DownloadEXT) <> mvarSize Then
        EndTransfer 7
    Else
        EndTransfer 4
    End If
    
Exit Sub
TCPReceiveError:
'Select Case Err.Number
EndTransfer 12
End Sub

Private Sub TCPReceive_Connect()

    Dim hFile As Long
    
    'If file without the .LUSerNet exists then rename it to .LUSerNet
    If FileExists(frmMain.txtDownloadLocation.Text & mvarFilename) Then
        If FileExists(frmMain.txtDownloadLocation.Text & mvarFilename & DownloadEXT) Then Kill frmMain.txtDownloadLocation.Text & mvarFilename & DownloadEXT
        On Error Resume Next
        Name frmMain.txtDownloadLocation.Text & mvarFilename As frmMain.txtDownloadLocation.Text & mvarFilename & DownloadEXT
        On Error GoTo 0
    End If
    
    mvarFilePath = frmMain.txtDownloadLocation.Text & mvarFilename & DownloadEXT

    'Truncate file to resume position
    hFile = CreateFile(mvarFilePath, GENERIC_WRITE, 0, ByVal CLng(0), OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    SetFilePointer hFile, mvarPosition, ByVal 0, FILE_BEGIN
    SetEndOfFile hFile
    CloseHandle hFile
    
    mvarlastTime = Now()
    
    mvarFileNumber = FreeFile(1)
       
    Open mvarFilePath For Binary As mvarFileNumber
    Seek mvarFileNumber, mvarPosition + 1
    Let status = 6
End Sub

Private Sub TCPReceive_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String
    TCPReceive.GetData Data
    If Len(Data) + mvarPosition > mvarSize Then
        'Done transfer
        Data = Left(Data, mvarSize - mvarPosition)
        Put mvarFileNumber, , Data
        EndTransfer 4
    Else
        'Write to file
        On Error GoTo TCPReceive_DataArrivalError
        Put mvarFileNumber, , Data
        On Error GoTo 0
    End If
    mvarPosition = mvarPosition + Len(Data)
    Exit Sub
    
TCPReceive_DataArrivalError:
Select Case Err.number
    Case Is = 61: EndTransfer (11)
    Case Else: msg "Unknown Error:" & Err.number & " " & Err.Description
End Select
End Sub

Public Sub EndTransfer(Optional code As Long = 9)
    mvarSpeed = 0
    TCPReceive.Close
    Close mvarFileNumber
    Let status = code
    If code = 4 Then
        On Error GoTo DownloadEndTransferError
        Name mvarFilePath As Left(mvarFilePath, Len(mvarFilePath) - Len(DownloadEXT))
        On Error GoTo 0
        RaiseEvent Finished
    End If

Exit Sub
DownloadEndTransferError:
If Err.number = 55 Then msg "There was a error renaming:" & vbCrLf & mvarFilePath & vbCrLf & "to:" & vbCrLf & Left(mvarFilePath, Len(mvarFilePath) - Len(DownloadEXT) & vbCrLf & "You will have to do this manually to complete the download")
End Sub

Public Function Receiving() As Boolean
    Select Case mvarStatus
        Case Is = 1, 5, 6: Receiving = True
        Case Else: Receiving = False
    End Select
End Function

Private Sub UserControl_Terminate()
    If Receiving Then EndTransfer
    
    If mvarTransferID <> -1 Then
        On Error Resume Next
        frmMain.lstTransfers.ListItems.Remove "I" & mvarTransferID
    End If
    
End Sub
