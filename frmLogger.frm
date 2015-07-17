VERSION 5.00
Begin VB.Form frmLogger 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "UDP Logger"
   ClientHeight    =   4680
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   7560
   Icon            =   "frmLogger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   4455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LogFileNumber As Long
Public LogOn As Boolean

Public Sub AddToLog(message As String, msgIn As Boolean)

Dim data As String

If LogOn Then
    If msgIn Then
        data = Date & " " & Time & " < " & message & vbCrLf
    Else
        data = Date & " " & Time & " > " & message & vbCrLf
    End If
    txtLog.Text = txtLog.Text & data
    txtLog.SelStart = Len(txtLog.Text)
    Print #LogFileNumber, data;
End If

End Sub

Private Sub Form_Load()
If LogOn Then
    LogFileNumber = FreeFile
    Open App.Path & "\log.txt" For Append As LogFileNumber
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close LogFileNumber
End Sub
