VERSION 5.00
Begin VB.Form frmRUSure 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Are you sure?"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   258
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.OptionButton optQuit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Stop sharing and quit once finished transfering"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Value           =   -1  'True
      Width           =   3615
   End
   Begin VB.OptionButton optQuit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Stop the transfers and quit"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin LUSerNet.MySkinner Skinner 
      Left            =   120
      Top             =   1560
      _ExtentX        =   1693
      _ExtentY        =   635
      imgNE           =   "frmRUSure.frx":0000
      imgN            =   "frmRUSure.frx":0166
      imgNW           =   "frmRUSure.frx":02CC
      imgE            =   "frmRUSure.frx":0882
      imgW            =   "frmRUSure.frx":0904
      imgSE           =   "frmRUSure.frx":0986
      imgS            =   "frmRUSure.frx":0A08
      imgSW           =   "frmRUSure.frx":0A8A
      BackColor       =   16777215
      ForeColor1      =   13545141
      ForeColor2      =   9402231
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure?"
      BeginProperty Font 
         Name            =   "Verdana"
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
      TabIndex        =   5
      Top             =   30
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You are currently download or uploading files. Would you like to:"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "frmRUSure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmMain.bClosing = 0
    Unload Me
End Sub

Private Sub cmdOk_Click()

    If optQuit(0).Value Then
        frmMain.bClosing = 2
    ElseIf optQuit(1).Value Then
        frmMain.bClosing = 1
        frmMain.lblConnected = "Not Connected"
    Else
        frmMain.bClosing = 0
    End If
    
    Unload Me
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormDrag Me
End Sub

Private Sub Form_Resize()
    Skinner.Repaint Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormDrag Me
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormDrag Me
End Sub
