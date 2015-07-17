VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'    PlayWAV "H:\My Projects\My VB\LUSerNet\graphics\heil.wav"
'    Dim byteArray() As Byte
'    byteArray = LoadResData("110", 10)
'    Dim fileNumber As Long
'    fileNumber = FreeFile
'    Open "c:\windows\temp\hitler.wav" For Output As #1
'    Dim i As Long
'    For i = LBound(byteArray) To UBound(byteArray)
'        Put fileNumber, , byteArray(i)
'    Next i
'    PlayWAV "c:\windows\temp\hitler.wav"
'    Kill "c:\windows\temp\hitler.wav".
PlayWaveRes "HITLER"
End Sub

