VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarFilename As String 'local copy
Private mvarPosition As Long 'local copy
Private mvarSize As Long 'local copy
Private mvarStatus As Long 'local copy
Private mvarPort As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarTransferID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvartmrTCPID As Long 'local copy
Private mvarTCPListenID As Long 'local copy
Private mvarTCPID As Long 'local copy
Private mvarlastPosition As Long 'local copy
Private mvarFileNumber As Long 'local copy
Private mvarTCPTransferID As Long 'local copy
Private mvarlastTime As Date 'local copy
Public Property Let lastTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lastTime = 5
    mvarlastTime = vData
End Property


Public Property Get lastTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lastTime
    lastTime = mvarlastTime
End Property



Public Property Let TCPTransferID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TCPTransferID = 5
    mvarTCPTransferID = vData
End Property


Public Property Get TCPTransferID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TCPTransferID
    TCPTransferID = mvarTCPTransferID
End Property



Public Property Let fileNumber(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fileNumber = 5
    mvarFileNumber = vData
End Property


Public Property Get fileNumber() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fileNumber
    fileNumber = mvarFileNumber
End Property


Public Property Let lastPosition(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lastPosition = 5
    mvarlastPosition = vData
End Property


Public Property Get lastPosition() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lastPosition
    lastPosition = mvarlastPosition
End Property



Public Property Let TCPID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TCPID = 5
    mvarTCPID = vData
End Property


Public Property Get TCPID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TCPID
    TCPID = mvarTCPID
End Property



Public Property Let TCPListenID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TCPListenID = 5
    mvarTCPListenID = vData
End Property


Public Property Get TCPListenID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TCPListenID
    TCPListenID = mvarTCPListenID
End Property



Public Property Let tmrTCPID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.tmrTCPID = 5
    mvartmrTCPID = vData
End Property


Public Property Get tmrTCPID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.tmrTCPID
    tmrTCPID = mvartmrTCPID
End Property



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


Public Property Let Port(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Port = 5
    mvarPort = vData
End Property

Public Property Get Port() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Port
    Port = mvarPort
End Property

Public Property Let Status(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Status = 5
    mvarStatus = vData
    If mvarTransferID <> -1 Then
        Select Case vData
            Case Is = 1: frmMain.lstTransfers.ListItems(mvarTransferID).ListSubItems(2).Text = "Pending"
            Case Is = 2: frmMain.lstTransfers.ListItems(mvarTransferID).ListSubItems(2).Text = "Timed Out"
            Case Is = 3: frmMain.lstTransfers.ListItems(mvarTransferID).ListSubItems(2).Text = "Refused"
            Case Is = 4: frmMain.lstTransfers.ListItems(mvarTransferID).ListSubItems(2).Text = "Complete"
            Case Is = 5: frmMain.lstTransfers.ListItems(mvarTransferID).ListSubItems(2).Text = "Sending"
            Case Is = 6: frmMain.lstTransfers.ListItems(mvarTransferID).ListSubItems(2).Text = "Receiving"
            Case Is = 7: frmMain.lstTransfers.ListItems(mvarTransferID).ListSubItems(2).Text = "Remote Error"
            Case Is = 8: frmMain.lstTransfers.ListItems(mvarTransferID).ListSubItems(2).Text = "File Not Found"
            Case Else: frmMain.lstTransfers.ListItems(mvarTransferID).ListSubItems(2).Text = "Unknown"
        End Select
    End If
End Property

Public Property Get Status() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
End Property

Public Property Let size(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Size = 5
    mvarSize = vData
End Property


Public Property Get size() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Size
    size = mvarSize
End Property

Public Property Let position(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Position = 5
    mvarPosition = vData
End Property

Public Property Get position() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Position
    position = mvarPosition
End Property

Public Property Let filename(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Filename = 5
    mvarFilename = vData
    On Error Resume Next
    mvarSize = FileLen(vData)
End Property

Public Property Get filename() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Filename
    filename = mvarFilename
End Property

Private Sub Class_Initialize()
    mvarTransferID = -1
    mvarlastTime = Now()
End Sub
