VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuDownload 
         Caption         =   "&Download"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTransferPopup 
      Caption         =   "mnuTransferPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenDownload 
         Caption         =   "&Open"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuStat 
         Caption         =   "-Stats-"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Progress: "
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTimeLeft 
         Caption         =   "TimeLeft:"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGap 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove Finished"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRemoveAll 
         Caption         =   "Remove &All Finished"
      End
      Begin VB.Menu mnuStopDownload 
         Caption         =   "&Stop"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuDownload_Click()
    Call frmMain.mnuDownload
End Sub

Private Sub mnuOpenDownload_Click()
    Call frmMain.mnuOpenDownload
End Sub

Private Sub mnuRemove_Click()
    Call frmMain.mnuRemove
End Sub

Private Sub mnuRemoveAll_Click()
    Call frmMain.mnuRemoveAll
End Sub

Private Sub mnuStopDownload_Click()
    Call frmMain.mnuStopDownload
End Sub
