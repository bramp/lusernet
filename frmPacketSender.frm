VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPacketSender 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "frmLUSerNet Packet Sender"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock udp 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   9876
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmPacketSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Hash(ByRef data As String)
    Dim outData As String
    Dim i As Long
    outData = ""
    For i = 1 To Len(data)
        outData = outData & Chr(Asc(Mid(data, i, 1)) Xor 128)
    Next i
    data = outData
End Sub

Private Sub cmdSend_Click()
    Dim data As String
    data = txtSend.Text
    Hash data
    udp.SendData data
End Sub
