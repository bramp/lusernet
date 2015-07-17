VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin prjTransferTest.Download Download 
      Left            =   3120
      Top             =   120
      _ExtentX        =   2117
      _ExtentY        =   529
   End
   Begin prjTransferTest.Upload Upload 
      Left            =   120
      Top             =   120
      _ExtentX        =   1535
      _ExtentY        =   529
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdShare 
      Caption         =   "Share"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin MSComctlLib.ImageList imgTransfer 
      Left            =   360
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferTest.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferTest.frx":041B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstTransfers 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgTransfer"
      SmallIcons      =   "imgTransfer"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FileName"
         Object.Width           =   5027
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   1720
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Progress"
         Text            =   "Progress"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Speed"
         Object.Width           =   1455
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const pos = 0 'Resume from
Const size = 12510273 'File size
 
Private Sub cmdDownload_Click()
    Download.Start "127.0.0.1", "9000", "spock", pos, size
End Sub

Private Sub cmdShare_Click()
    Upload.Listen "127.0.0.1", "80", "c:\spock.mp3", pos
End Sub

'Private Sub Download_Finished()
'    MsgBox "LO"
'End Sub

