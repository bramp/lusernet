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
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()



Static bStarted As Boolean

If bStarted = True Then Exit Sub
bStarted = True

Dim folderPath As String
'Cycle Through All Folders in Folders
folderPath = "c:\tmp\"
GetFolders folderPath, -1

bStarted = False



End Sub

Private Sub Form_Load()
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\LUSerNet.mdb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
conn.Close
Set conn = Nothing
Set rs = Nothing
End Sub
