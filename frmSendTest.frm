VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Download"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock udp 
      Left            =   2280
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   9876
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    udp.SendData "GET|ac3-install.exe|53248|0|127.0.0.1"
End Sub
