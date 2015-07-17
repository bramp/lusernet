VERSION 5.00
Begin VB.Form frmLeecher 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close This Message"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   1695
   End
   Begin VB.PictureBox picAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6375
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.Label lblLeecher 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You are a l33cher, please share your files!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   9495
      End
      Begin VB.Label lblLeech 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   9495
      End
      Begin VB.Label lblLuser 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   0
         TabIndex        =   1
         Top             =   2640
         Width           =   9495
      End
   End
   Begin VB.Timer tmrFlash 
      Interval        =   500
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmLeecher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Left = 0
    Me.Top = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    
    picAll.Left = (Me.Width - picAll.Width) / 2
    picAll.Top = (Me.Height - picAll.Height) / 2
    
    cmdClose.Left = Me.Width - 120 - cmdClose.Width
    cmdClose.Top = Me.Height - 120 - cmdClose.Height
    
    lblLeech.Caption = "l33ch n. (Also `l33cher'.)" & vbCrLf & "    Among BBS types, crackers and warez d00dz, one who consumes knowledge without generating new software, cracks, or techniques." & vbCrLf & "    BBS culture specifically defines a leech as someone who downloads files with few or no uploads in return, and who does not contribute to the message section." & vbCrLf & "    Cracker culture extends this definition to someone (a lamer, usually) who constantly presses informed sources for information and/or assistance, but has nothing to contribute."
    lblLuser.Caption = "luser    /loo'zr/ n. [common]" & vbCrLf & "    A user; esp. one who is also a loser. ( luser and loser are pronounced identically.) This word was coined around 1975 at MIT. Under ITS, when you first walked up to a terminal at MIT and typed Control-Z to get the computer's attention, it printed out some status information, including how many people were already using the computer; it might print '14 users', for example." & vbCrLf & "Someone thought it would be a great joke to patch the system to print '14 losers' instead. There ensued a great controversy, as some of the users didn't particularly want to be called losers to their faces every time they used the computer."
    lblLuser.Caption = lblLuser.Caption & vbCrLf & "    For a while several hackers struggled covertly, each changing the message behind the back of the others; any time you logged into the computer it was even money whether it would say 'users' or 'losers'. Finally, someone tried the compromise 'lusers', and it stuck. Later one of the ITS machines supported luser as a request-for-help command. ITS died the death in mid-1990, except as a museum piece; the usage lives on, however, and the term `luser' is often seen in program comments and on Usenet."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim downloads As Single
    Dim uploads As Single
    downloads = GetSetting("LUSerNet", "Main", "TotalDownload", 0)
    uploads = GetSetting("LUSerNet", "Main", "TotalUpload", 0)
    msg "You have downloaded " & ChangeByte(downloads, True, 0) & " and only uploaded " & ChangeByte(uploads, True, 0) & vbCrLf & "PLEASE share more to never see this again"
    frmMain.bLeecherOn = False
End Sub

Private Sub tmrFlash_Timer()
    lblLeecher.Visible = Not lblLeecher.Visible
    PlayWaveRes "SIREN"
End Sub
