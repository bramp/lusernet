VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmNetwork 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   334
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock UDP 
      Left            =   2040
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2580
      TabIndex        =   13
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3780
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdDetect 
      Caption         =   "Auto-Detect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Network"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4755
      Begin VB.OptionButton optNetwork 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Manually specify"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox cmbNetwork 
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   4215
      End
      Begin VB.OptionButton optNetwork 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   260
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame frmIPRange 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "IP Range"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   4755
      Begin AutoUpdate.MyIPRange IPRange 
         Height          =   285
         Left            =   3000
         TabIndex        =   14
         Top             =   210
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Your IP Range:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblIP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "127.0.0.1"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Your IP:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox cmdClose 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4740
      Picture         =   "frmNetwork.frx":0000
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   1
      Top             =   60
      Width           =   210
   End
   Begin InetCtlsObjects.Inet inet 
      Left            =   1320
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNetwork.frx":0512
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LUSerNet Network Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   30
      Width           =   4215
   End
End
Attribute VB_Name = "frmNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Network
    Name As String
    IPRange As String
    BannedRange As String
End Type

Private Sub cmbNetwork_Click()
    optNetwork(0).Value = True
End Sub

Private Sub cmdCancel_Click()
    Unload frmNetwork
End Sub

Private Sub cmdClose_Click()
    Unload frmNetwork
End Sub

Private Sub cmdDetect_Click()
    'Determines the correct IP address
    Dim ipAddys() As String
    ipAddys = LocalIPAddresses

    'Dim ipAddysStr As String
    Dim i As Long
    
    'For i = LBound(ipAddys) To UBound(ipAddys)
    '    ipAddysStr = ipAddysStr & ipAddys(i) & vbCrLf
    'Next i
    
    i = LBound(ipAddys)
    
    If i = 0 Then ipAddys(0) = GetIP 'win98 hack

    'ipAddys is a array of IPs on this machine, lets try and pick the best one

    Dim found As Boolean
    found = False

    Do While i < UBound(ipAddys) And Not found
        Select Case ipAddys(i)
            Case Is = "127.0.0.1": 'Do Nothing
            Case Else: found = True
        End Select
        i = i + 1
    Loop

    lblIP.Caption = ipAddys(i - 1)
    IPRange.IpAddress = Left(ipAddys(i - 1), InStr(InStr(1, ipAddys(i - 1), ".") + 1, ipAddys(i - 1), ".")) & "255.255"

End Sub

Private Sub Form_Load()
    frmUpdate.MySkinner.Repaint frmNetwork
    
    Dim ipAddys() As String
    ipAddys = LocalIPAddresses

    If LBound(ipAddys) = 0 Then ipAddys(0) = GetIP 'win98 hack
   
    Dim NetworkData As String
    Dim Networks() As Network
    Dim tmp() As String
    Dim tmp2() As String
   
    On Error GoTo errorTimeOut
    NetworkData = inet.OpenURL(UpdateTXTURL & "network.txt")
    On Error GoTo 0
    
    On Error GoTo errorHandle
    
    cmbNetwork.Clear
    cmbNetwork.Text = "Select Your Network"
    
    If Left(Trim(NetworkData), 1) = "<" Then
        cmbNetwork.AddItem "Unable to download Network Lists"
        Exit Sub
    End If
    
    tmp = Split(NetworkData, vbCrLf)
    
    ReDim Networks(UBound(tmp))
    
    Dim i As Long
    
    For i = LBound(tmp) To UBound(tmp)
        tmp2 = Split(tmp(i), "|")
        Networks(i).Name = tmp2(0)
        cmbNetwork.AddItem tmp2(0)
        Networks(i).IPRange = tmp2(1)
        Networks(i).BannedRange = tmp2(2)
    Next i
    
    On Error GoTo 0
    
Exit Sub
errorHandle:
Select Case Err.Number
    Case Is = 13: '"Error occured downloading update list"
    Case Else: MsgBox "Unknown Error: " & Err.Number & " " & Err.Description
End Select

Exit Sub
errorTimeOut:
Select Case Err.Number
    Case Is = 35761: 'Error occured connecting to download server, try again later"
    Case Else: MsgBox "Unknown Error: " & Err.Number & " " & Err.Description
End Select
End Sub

Private Sub Form_Resize()
    frmUpdate.MySkinner.Repaint frmNetwork
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub frmIPRange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub lblIP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub optNetwork_Click(Index As Integer)
    If Index = 1 Then
        cmdDetect.Enabled = True
        frmIPRange.Enabled = True
        Label1.Enabled = True
        Label2.Enabled = True
        lblIP.Enabled = True
        IPRange.Enabled = True
    Else
        cmdDetect.Enabled = False
        frmIPRange.Enabled = False
        Label1.Enabled = False
        Label2.Enabled = False
        lblIP.Enabled = False
        IPRange.Enabled = False
    End If
End Sub

Function GetIP() As String
    GetIP = UDP.LocalIP
End Function
