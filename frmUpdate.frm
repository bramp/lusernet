VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmUpdate 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "LUSerNet Update"
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   StartUpPosition =   3  'Windows Default
   Begin AutoUpdate.MySkinner MySkinner 
      Left            =   3480
      Top             =   360
      _extentx        =   1693
      _extenty        =   635
      imgne           =   "frmUpdate.frx":1442
      imgn            =   "frmUpdate.frx":15A8
      imgnw           =   "frmUpdate.frx":1658
      imge            =   "frmUpdate.frx":1C0E
      imgw            =   "frmUpdate.frx":1C6E
      imgse           =   "frmUpdate.frx":1CCE
      imgs            =   "frmUpdate.frx":1D50
      imgsw           =   "frmUpdate.frx":1DB4
      backcolor       =   16777215
      forecolor1      =   0
      forecolor2      =   0
   End
   Begin VB.Timer tmrStart 
      Interval        =   100
      Left            =   120
      Top             =   1200
   End
   Begin VB.PictureBox cmdClose 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4290
      Picture         =   "frmUpdate.frx":1EF6
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   3
      Top             =   45
      Width           =   210
   End
   Begin InetCtlsObjects.Inet inet 
      Left            =   3840
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LUSerNet Update"
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
      TabIndex        =   4
      Top             =   30
      Width           =   2055
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Checking For Update...."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload frmUpdate
    End
End Sub

Private Sub cmdAbort_Click()
    If cmdAbort.Caption <> "Abort" Then
        If OpenLUSerNet Then
            Unload frmUpdate
            End
        End If
    Else
        Unload frmUpdate
        End
    End If
End Sub

Private Sub Form_Load()

MySkinner.Repaint frmUpdate

frmUpdate.Show
DoEvents

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub lblStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Progress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub tmrStart_Timer()

tmrStart.Enabled = False
lblStatus.Caption = "Checking For Update..."

Dim UpdateData As String
On Error GoTo errorTimeOut
UpdateData = inet.OpenURL(UpdateTXTURL & "update.txt")
On Error GoTo 0

'MsgBox UpdateData

Dim tmp() As String
Dim tmp2() As String
Dim VersionNumber As Long
Dim downloads() As String
Dim files() As String
Dim i As Long

tmp = Split(UpdateData, vbCrLf)
On Error GoTo errorHandle
VersionNumber = CLng(tmp(0))
On Error GoTo 0

If VersionNumber > GetSetting("LUSerNet", "Main", "Version", 0) Then
    
    For i = 1 To UBound(tmp)
        ReDim Preserve downloads(i - 1)
        ReDim Preserve files(i - 1)
        tmp2 = Split(tmp(i), "|")
        downloads(i - 1) = tmp2(0)
        files(i - 1) = tmp2(1)
    Next i
    
    Dim b() As Byte
    Progress.Max = UBound(downloads) + 1
    For i = 0 To UBound(downloads)
        
        lblStatus.Caption = "Updating " & downloads(i) & "..."
        DoEvents
        On Error GoTo errorHandle
        Open App.Path & "\" & files(i) For Binary Access Write As #1
        On Error GoTo 0
        
        b() = inet.OpenURL(UpdateTXTURL & downloads(i), icByteArray)
        
        Put #1, , b()
        Close #1
        
        Progress.Value = i + 1
    
    Next i
    
    SaveSetting "LUSerNet", "Main", "Version", VersionNumber
    lblStatus.Caption = "Done All Updates"
    lblStatus.Refresh
Else
    lblStatus.Caption = "No Updates Found"
    lblStatus.Refresh
End If

If OpenLUSerNet Then
    Unload frmUpdate
    End
End If

Exit Sub
errorHandle:
Select Case Err.Number
    Case Is = 13: lblStatus.Caption = "Error occured downloading update list": cmdAbort.Caption = "Continue"
    Case Is = 70: If MsgBox("A Error Occured Trying To Patch " & files(i) & " - If this file is open, close it!", vbRetryCancel) = vbRetry Then Resume
    Case Else: MsgBox "Unknown Error: " & Err.Number & " " & Err.Description
End Select

Exit Sub
errorTimeOut:
Select Case Err.Number
    Case Is = 35761: lblStatus.Caption = "Error occured connecting to download server, try again later": cmdAbort.Caption = "Continue"
    Case Else: MsgBox "Unknown Error: " & Err.Number & " " & Err.Description
End Select

End Sub

Function OpenLUSerNet() As Boolean
    'Opens LUSerNet returns true if done so
    
    OpenLUSerNet = False
    
    'Firstly we check network settings
    frmNetwork.Show vbModal, Me
    
    On Error GoTo OpenLUSerNetError
    'Shell App.Path & "\LUSerNet.exe /done"
    OpenLUSerNet = True
    
    Exit Function
OpenLUSerNetError:
    Select Case Err.Number
        Case Is = 53: lblStatus.Caption = "There was a error opening LUSerNet": cmdAbort.Caption = "Abort"
        Case Else: MsgBox "Unknown Error: " & Err.Number & " " & Err.Description
    End Select
End Function
