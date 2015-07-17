VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl Upload 
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer tmrSpeed 
      Interval        =   2000
      Left            =   120
      Top             =   1440
   End
   Begin VB.Timer tmrTCPSend 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   120
      Top             =   960
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   960
   End
   Begin MSWinsockLib.Winsock TCPListen 
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock TCPSend 
      Left            =   600
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Upload"
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
      Width           =   870
   End
End
Attribute VB_Name = "Upload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarFilename As String 'local copy
Private mvarSize As Long 'local copy
Private mvarPosition As Long 'local copy
Private mvarFileNumber As Long 'local copy
Private mvarRemoteHost As String 'local copy
Private mvarTransferID As Long 'local copy
Private mvarStatus As Long 'local copy
Private mvarLocalPort As Long 'local copy
Private mvarlastTime As Date 'local copy
Private mvarlastPosition As Long 'local copy
Private mvarSpeed As Long 'local copy
Private mvarEOF As Boolean 'local copy

Public Event Finished()
Public Event Update()

Private Sub TCPListen_ConnectionRequest(ByVal requestID As Long)
    'Start transfer
    TCPSend.Accept requestID
    EndListening
    
    mvarlastTime = Now()
    tmrTCPSend.Enabled = True
       
    mvarFileNumber = FreeFile(1)
    
    Open mvarFilename For Random Shared As mvarFileNumber Len = bufferSize
    Let status = 5 'Sending
    Seek mvarFileNumber, mvarPosition \ bufferSize + 1
End Sub

Private Sub TCPSend_Close()
    mvarEOF = True
    CloseFile
End Sub

Private Sub CloseFile()
If mvarFileNumber <> 0 Then
    Close mvarFileNumber
    mvarFileNumber = 0
End If
End Sub

Private Sub TCPSend_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal Helpfile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    TCPSend.Tag = ""
End Sub

Private Sub TCPSend_SendComplete()
    TCPSend.Tag = ""
End Sub

Private Sub tmrSpeed_Timer()
    
    If TCPSend.State = sckConnected Then

        'Displays Speed
        mvarSpeed = getSpeed
        
        mvarlastPosition = mvarPosition
        mvarlastTime = Now()
        frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(3).Tag = 0 - mvarSpeed
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
    mvarlastTime = Now()
End Sub

Private Sub UserControl_Resize()
    Width = lblName.Width
    Height = lblName.Height
End Sub

Private Sub tmrTCPSend_Timer()
     
    Dim Data As String * bufferSize
       
'    me.Status = 5 'Sending

    'DoEvents
    If TCPSend.Tag <> "SENDING" Then
        'If TCPSend.State <> sckConnected Then
        If TCPSend.State <> sckConnected Or MyEOF(mvarFileNumber) Then
            tmrTCPSend.Enabled = False
            
            If MyEOF(mvarFileNumber) Then
                Let status = 4 'complete
                RaiseEvent Finished
            Else
                Let status = 7 'Remote Error
            End If
        
            TCPSend.Close
        Else
            On Error GoTo TCP_ConnectError
            TCPSend.Tag = "SENDING"
            Get mvarFileNumber, (Me.position \ bufferSize) + 1, Data
            TCPSend.SendData Data 'sends in bufferSize chunks
            Me.position = Me.position + bufferSize
            On Error GoTo 0
        End If
    End If
Exit Sub
TCP_ConnectError:
    Select Case Err.number
        Case 2: DoEvents: Resume 'the wierd error 2 bug
        Case Else: msg Err.number & " " & Err.Description & " - in tmrTimeOut_Timer"
    End Select
Exit Sub
End Sub

Private Function MyEOF(number As Long) As Boolean
    If Not mvarEOF Then
        MyEOF = EOF(number)
    Else
        MyEOF = True
    End If
End Function

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
            Case Else: frmMain.lstTransfers.ListItems("I" & mvarTransferID).ListSubItems(2).Text = "Unknown"
        End Select
        Select Case vData
            Case Is = 1, 2, 3, 4, 7, 8, 9
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

Public Property Let LocalPort(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.status = 5
    mvarLocalPort = vData
End Property

Public Property Get LocalPort() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.status
    LocalPort = mvarLocalPort
End Property

Public Sub Listen(remoteHost As String, LocalPort As Long, ByVal filename As String, position As Long)
    mvarRemoteHost = remoteHost
    
    mvarFilename = filename
    
    Dim Exists As Boolean
    Exists = FileExists(mvarFilename)
    If Exists Then
        mvarSize = FileLen(filename)
        
        'Convert Position to a multiply of BufferSize
        mvarPosition = (position \ bufferSize) * bufferSize
        
        TCPListen.LocalPort = LocalPort
        On Error GoTo listenError
        TCPListen.Listen
        mvarLocalPort = LocalPort
        
        tmrTimeOut.Interval = 10
        tmrTimeOut.Enabled = True
        
        Dim tmpFile() As String
        tmpFile = Split(filename, "\")
        filename = tmpFile(UBound(tmpFile))
    End If
    
    Dim item As ListItem
    Dim index As Long
    index = frmMain.lstTransfers.Tag
    frmMain.lstTransfers.Tag = frmMain.lstTransfers.Tag + 1
    Set item = frmMain.lstTransfers.ListItems.Add(, "I" & index, filename, 1, 1)
    If Exists Then
        item.ListSubItems.Add , , ChangeByte(FileLen(mvarFilename), True)
    Else
        item.ListSubItems.Add , , ChangeByte(0, True)
    End If
    item.ListSubItems.Add , , "Pending"
    item.ListSubItems.Add , , "0 b/S"
    mvarTransferID = index
    Set item = Nothing
    
    If Exists Then
        Let status = 1
    Else
        Let status = 8
    End If
    Exit Sub
listenError:
    Select Case Err.number
        Case Is = 10048:
            LocalPort = LocalPort + 1
            TCPListen.LocalPort = LocalPort
            Resume
        Case Else: msg "Unknown Error:" & Err.number & " " & Err.Description
    End Select
End Sub

Public Function getProgress() As Single
If mvarSize = 0 Then
    getProgress = 0
Else
    getProgress = mvarPosition / mvarSize * 100
End If
End Function

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

Public Sub EndListening(Optional code As Long = 99)
    If code <> 99 Then Let status = code
    TCPListen.Close
End Sub

Private Sub tmrTimeOut_Timer()

If tmrTimeOut.Tag <> "WAITING" Then
    tmrTimeOut.Enabled = False

    'Uploads(Index).listen remoteHost, LocalPort, filename, position
    
    'They got 10 seconds to connection, otherwise bye bye
    tmrTimeOut.Interval = 1000 * 10
    tmrTimeOut.Tag = "WAITING"
    tmrTimeOut.Enabled = True
Else
    tmrTimeOut.Enabled = False
    tmrTimeOut.Tag = ""
    If mvarStatus = 1 Then  'pending
        'unload tcp stuff
        Me.EndListening 2
        'Uploads.Remove Index
    End If
End If
End Sub

Public Sub EndTransfer()
    EndListening
    tmrTCPSend.Enabled = False
    TCPSend.Close
    Let status = 9
End Sub

Private Sub UserControl_Terminate()
    EndListening
    If mvarTransferID <> -1 Then
        On Error Resume Next
        frmMain.lstTransfers.ListItems.Remove "I" & mvarTransferID
    End If
    CloseFile
End Sub

Public Function Sending() As Boolean
    Select Case mvarStatus
        Case Is = 1, 5, 6: Sending = True
        Case Else: Sending = False
    End Select
End Function
