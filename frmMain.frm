VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "LUSerNet"
   ClientHeight    =   7725
   ClientLeft      =   150
   ClientTop       =   510
   ClientWidth     =   10800
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   515
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   720
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEAEB5&
      ForeColor       =   &H80000008&
      Height          =   5025
      Index           =   1
      Left            =   240
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   3
      Tag             =   "Search"
      Top             =   960
      Width           =   6090
      Begin VB.PictureBox picCombo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   330
         TabIndex        =   45
         Top             =   4560
         Width           =   4980
         Begin VB.ComboBox cmbSearch 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmMain.frx":1442
            Left            =   -30
            List            =   "frmMain.frx":1444
            TabIndex        =   31
            Tag             =   "1"
            Text            =   "Type Search Word Here"
            ToolTipText     =   "Type search word here"
            Top             =   -30
            Width           =   5025
         End
      End
      Begin VB.CommandButton cmdSearchGo 
         Caption         =   "Search"
         Height          =   270
         Left            =   5205
         TabIndex        =   4
         ToolTipText     =   "Click here to search"
         Top             =   4575
         Width           =   720
      End
      Begin MSComctlLib.ListView lstMain 
         Height          =   4335
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Displays a list of found files"
         Top             =   120
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgSearch"
         SmallIcons      =   "imgSearch"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FileName"
            Object.Width           =   6747
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   1773
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Speed"
            Object.Width           =   1191
         EndProperty
      End
      Begin MSComctlLib.ListView lstTMP 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgSearch"
         SmallIcons      =   "imgSearch"
         ColHdrIcons     =   "imgSearch"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FileName"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2558
         EndProperty
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEAEB5&
      ForeColor       =   &H80000008&
      Height          =   5025
      Index           =   0
      Left            =   0
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   19
      Tag             =   "Welcome"
      Top             =   720
      Width           =   6090
      Begin VB.TextBox txtMOTD 
         Appearance      =   0  'Flat
         BackColor       =   &H00CEAEB5&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   39
         Text            =   "frmMain.frx":1446
         Top             =   4350
         Width           =   4335
      End
      Begin VB.Label lblHamor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   5055
         Left            =   5880
         TabIndex        =   75
         Top             =   0
         Width           =   255
      End
      Begin VB.Image picLogo 
         Appearance      =   0  'Flat
         Height          =   3045
         Left            =   840
         Picture         =   "frmMain.frx":1453
         Top             =   840
         Width           =   4290
      End
      Begin VB.Label lblMoreInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00CEAEB5&
         Caption         =   "Click here to find out more"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1800
         MouseIcon       =   "frmMain.frx":2BE89
         MousePointer    =   99  'Custom
         TabIndex        =   41
         ToolTipText     =   "Click here to find out more about the MOTD"
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label lblMOTD 
         Alignment       =   2  'Center
         BackColor       =   &H00CEAEB5&
         Caption         =   "Message Of The Day"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   40
         ToolTipText     =   "A short message to inform you of changes to LUSerNet"
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Shape shpMOTD 
         BackColor       =   &H00CEAEB5&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H009C8486&
         Height          =   495
         Left            =   720
         Top             =   4200
         Width           =   4575
      End
      Begin VB.Label lblTotalFiles 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TotalFiles"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         ToolTipText     =   "This shows how many files there are on the network"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTotalFolders 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TotalFolders"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblTotalSize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TotalSize"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   25
         ToolTipText     =   "This shows the total size of all files on the network"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblTotalUsers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TotalUsers"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         ToolTipText     =   "This shows how many users are online"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblConnected 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Not Connected - Maybe a firewall in the way?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Tag             =   "1"
         ToolTipText     =   "This shows the status of your connection"
         Top             =   60
         Width           =   5775
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Online Users:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Files Shared:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total File Size:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H008F7777&
         Height          =   480
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   6135
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEAEB5&
      ForeColor       =   &H80000008&
      Height          =   5025
      Index           =   3
      Left            =   720
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   7
      Tag             =   "Sharing"
      Top             =   1440
      Width           =   6090
      Begin VB.CommandButton Command1 
         Caption         =   "Crash"
         Height          =   375
         Left            =   4800
         TabIndex        =   79
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CEAEB5&
         Caption         =   "Transfered"
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   4440
         TabIndex        =   76
         ToolTipText     =   "Shows how much you are sharing"
         Top             =   2280
         Width           =   1455
         Begin VB.Label lblTotalDownloads 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Download 100GB"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblTotalUploads 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Upload 100GB"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdSelectDownloadLocation 
         Caption         =   "Browse..."
         Height          =   285
         Left            =   5040
         TabIndex        =   12
         ToolTipText     =   "Click here to change where your files go"
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox txtDownloadLocation 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "This is where all your download files go"
         Top             =   4560
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CEAEB5&
         Caption         =   "Shared"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   4440
         TabIndex        =   13
         ToolTipText     =   "Shows how much you are sharing"
         Top             =   960
         Width           =   1455
         Begin VB.Label lblShareFiles 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Files: 1000"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblShareFolders 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Folders: 1000"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Size:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblShareSize 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1000 Gigabytes"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdShareAdd 
         Caption         =   "Add Share"
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         ToolTipText     =   "Add a new folder to be shared"
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdShareRemove 
         Caption         =   "Remove Share"
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         ToolTipText     =   "Remove the selected folder from being shared"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdShareRefresh 
         Caption         =   "Refresh Share"
         Height          =   375
         Left            =   4440
         TabIndex        =   10
         ToolTipText     =   "Rescans all your shares for changes"
         Top             =   4080
         Width           =   1455
      End
      Begin MSComctlLib.TreeView treeShares 
         Height          =   4335
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Shows a list of all the folders/files being shared"
         Top             =   120
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   7646
         _Version        =   393217
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "Images"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Downloaded Files Go Here:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   4590
         Width           =   2055
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEAEB5&
      ForeColor       =   &H80000008&
      Height          =   5025
      Index           =   5
      Left            =   1200
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   30
      Tag             =   "About"
      Top             =   1920
      Width           =   6090
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "and last but not least ISS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   91
         Top             =   3960
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "GuyNextDoor, Fiery Badger Pr0n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   90
         Top             =   3600
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Beanie, Carter, Unbeleiver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   89
         Top             =   3240
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dave, Edd, Cheka"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   120
         TabIndex        =   88
         Top             =   2880
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Greets:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   87
         Top             =   2520
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Barny, bramp, Andy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   86
         Top             =   2160
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Emiri"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   85
         Top             =   1440
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "bramp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   84
         Top             =   720
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Webpage:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   83
         Top             =   1800
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Graphics:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   82
         Top             =   1080
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Coder:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   81
         Top             =   480
         Width           =   5775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":2C193
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   80
         Top             =   4440
         Width           =   5895
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008F7777&
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   6060
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEAEB5&
      ForeColor       =   &H80000008&
      Height          =   5025
      Index           =   4
      Left            =   960
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "Chat"
      Top             =   1680
      Width           =   6090
      Begin LUSerNet.HyperLink HyperLink 
         Height          =   255
         Left            =   1800
         TabIndex        =   59
         Top             =   3120
         Width           =   2655
         _ExtentX        =   4895
         _ExtentY        =   450
         Text            =   "www.lusernet.34sp.com/?page=chat"
         URL             =   "http://www.lusernet.34sp.com/?page=chat"
         BackColor       =   13545141
      End
      Begin LUSerNet.HyperLink HyperLink1 
         Height          =   255
         Left            =   1440
         TabIndex        =   64
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         Text            =   "http://lusernet.tk"
         URL             =   "http://lusernet.tk"
         BackColor       =   13545141
      End
      Begin LUSerNet.HyperLink HyperLink2 
         Height          =   255
         Left            =   1680
         TabIndex        =   65
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         Text            =   "http://lusernet.34sp.com/?page=docs"
         URL             =   "http://lusernet.34sp.com/?page=docs"
         BackColor       =   13545141
      End
      Begin LUSerNet.HyperLink HyperLink3 
         Height          =   255
         Left            =   1680
         TabIndex        =   66
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Text            =   "http://lusernet.34sp.com/forum"
         URL             =   "http://lusernet.34sp.com/forum"
         BackColor       =   13545141
      End
      Begin LUSerNet.HyperLink HyperLink4 
         Height          =   255
         Left            =   2520
         TabIndex        =   71
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Text            =   "bramp@lusernet.tk"
         URL             =   "mailto:bramp@lusernet.tk"
         BackColor       =   13545141
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "please do so via the forums, not by contacting me directly. Thanks"
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   4680
         Width           =   6615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "the above links.  Also if you would like to make any suggestions for future versions"
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   4440
         Width           =   6615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please note that I will try to answer all emails, but only email me after checking"
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   4200
         Width           =   6615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can contact bramp via email:"
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   3960
         Width           =   2370
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008F7777&
         Caption         =   "Email Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   69
         Top             =   3480
         Width           =   6060
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Also information about the program is usally released on the forum first."
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   1800
         Width           =   4935
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The forum offers a easy way for other people to quickly help you."
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   1560
         Width           =   4560
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discussion Forum:"
         Height          =   195
         Left            =   360
         TabIndex        =   63
         Top             =   1200
         Width           =   1290
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document Pages:"
         Height          =   195
         Left            =   360
         TabIndex        =   62
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Website: "
         Height          =   195
         Left            =   360
         TabIndex        =   61
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "There are a few ways to get help via our website:"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   480
         Width           =   3480
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QuakeNet at #LUSerNet more information on how to join can be found at"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   2880
         Width           =   5175
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can chat to the developers in real time via IRC. We are currently on "
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   2640
         Width           =   5130
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008F7777&
         Caption         =   "Real Time Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   56
         Top             =   2160
         Width           =   6060
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008F7777&
         Caption         =   "Web Based Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   6060
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEAEB5&
      ForeColor       =   &H80000008&
      Height          =   5025
      Index           =   2
      Left            =   480
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   1
      Tag             =   "Transfers"
      Top             =   1200
      Width           =   6090
      Begin LUSerNet.UpDown UpDown 
         Height          =   270
         Index           =   1
         Left            =   3225
         TabIndex        =   54
         ToolTipText     =   "How many downloads can happen at once"
         Top             =   4575
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   450
         Value           =   5
         Min             =   1
         Max             =   10
      End
      Begin LUSerNet.UpDown UpDown 
         Height          =   270
         Index           =   0
         Left            =   1785
         TabIndex        =   53
         ToolTipText     =   "How many uploads can happen at once"
         Top             =   4575
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   450
         Min             =   1
         Max             =   10
      End
      Begin VB.TextBox txtUploadTotal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "5"
         ToolTipText     =   "How many uploads can happen at once"
         Top             =   4560
         Width           =   495
      End
      Begin MSComctlLib.ListView lstTransfers 
         Height          =   4335
         Left            =   120
         TabIndex        =   2
         Tag             =   "0"
         ToolTipText     =   "Shows all current transfers in and out"
         Top             =   120
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   7646
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
      Begin VB.TextBox txtDownloadTotal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "5"
         ToolTipText     =   "How many downloads can happen at once"
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Downloads at the same time"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3540
         TabIndex        =   51
         Top             =   4605
         Width           =   2175
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Uploads or"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2130
         TabIndex        =   50
         Top             =   4605
         Width           =   855
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No more than"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   450
         TabIndex        =   48
         Top             =   4605
         Width           =   1455
      End
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   10000
      Left            =   9240
      Top             =   6480
   End
   Begin VB.Timer tmrMessage 
      Enabled         =   0   'False
      Left            =   8760
      Top             =   6480
   End
   Begin LUSerNet.MySkinner Skinner 
      Left            =   7680
      Top             =   4080
      _ExtentX        =   1693
      _ExtentY        =   635
      imgNE           =   "frmMain.frx":2C23E
      imgN            =   "frmMain.frx":2C3A4
      imgNW           =   "frmMain.frx":2C50A
      imgE            =   "frmMain.frx":2CAC0
      imgW            =   "frmMain.frx":2CB42
      imgSE           =   "frmMain.frx":2CBC4
      imgS            =   "frmMain.frx":2CC46
      imgSW           =   "frmMain.frx":2CCC8
      BackColor       =   16777215
      ForeColor1      =   13545141
      ForeColor2      =   9402231
   End
   Begin MSComctlLib.ImageList imgSearch 
      Left            =   8880
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CD4A
            Key             =   "unknown"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D144
            Key             =   "text"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D550
            Key             =   "audio"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D953
            Key             =   "movie"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DD72
            Key             =   "html"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E19A
            Key             =   "image"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E599
            Key             =   "zip"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E9AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EDB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F1A7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LUSerNet.miniGraph miniGraph 
      Height          =   255
      Left            =   7560
      Top             =   5640
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      BackColor       =   9402231
   End
   Begin VB.Timer tmrHello 
      Interval        =   60000
      Left            =   8280
      Top             =   6480
   End
   Begin MSComctlLib.ImageList imgTrayIcon 
      Left            =   8880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F59E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LUSerNet.Download Download 
      Index           =   0
      Left            =   7680
      Top             =   3720
      _ExtentX        =   2117
      _ExtentY        =   529
   End
   Begin LUSerNet.Upload Upload 
      Index           =   0
      Left            =   7680
      Top             =   3360
      _ExtentX        =   1535
      _ExtentY        =   529
   End
   Begin VB.Timer tmrMOTD 
      Interval        =   1
      Left            =   8760
      Top             =   6000
   End
   Begin LUSerNet.MyButton cmdClose 
      Height          =   210
      Left            =   5880
      TabIndex        =   37
      Top             =   60
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      picNormal       =   "frmMain.frx":309F0
      ToolTip         =   "Close LUSerNet"
   End
   Begin MSComctlLib.ImageList imgTransfer 
      Left            =   8160
      Top             =   840
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
            Picture         =   "frmMain.frx":30F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3132D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LUSerNet.MyButton cmdTab 
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      picNormal       =   "frmMain.frx":31761
      ToolTip         =   "Welcome Page"
   End
   Begin MSWinsockLib.Winsock UDPSend 
      Left            =   600
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   9876
   End
   Begin MSComctlLib.ImageList Images 
      Left            =   8160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":340F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34445
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34797
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34AE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34E89
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3527F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock UDPListen 
      Index           =   0
      Left            =   120
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   9876
   End
   Begin LUSerNet.MyButton cmdTab 
      Height          =   360
      Index           =   2
      Left            =   2280
      TabIndex        =   32
      Top             =   360
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
      picNormal       =   "frmMain.frx":356C7
      ToolTip         =   "Transfers Page - This shows all downloads and uploads"
   End
   Begin LUSerNet.MyButton cmdTab 
      Height          =   360
      Index           =   3
      Left            =   3480
      TabIndex        =   33
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      picNormal       =   "frmMain.frx":37F99
      ToolTip         =   "Share Page - You can choose what files you want to share on the network here"
   End
   Begin LUSerNet.MyButton cmdTab 
      Height          =   360
      Index           =   5
      Left            =   5280
      TabIndex        =   34
      Top             =   360
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   635
      picNormal       =   "frmMain.frx":3A4AB
      ToolTip         =   "About Page - Look here for some cool graphics and information about how the people behind LUSerNet"
   End
   Begin LUSerNet.MyButton cmdTab 
      Height          =   360
      Index           =   4
      Left            =   4560
      TabIndex        =   35
      Top             =   360
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   635
      picNormal       =   "frmMain.frx":3C47D
      ToolTip         =   "Need Help?"
   End
   Begin LUSerNet.MyButton cmdMinimise 
      Height          =   210
      Left            =   5610
      TabIndex        =   38
      Top             =   60
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      picNormal       =   "frmMain.frx":3DFD1
      ToolTip         =   "Minimise LUSerNet to the system tray"
   End
   Begin InetCtlsObjects.Inet iNetMOTD 
      Left            =   8160
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin LUSerNet.MyButton cmdTab 
      Height          =   360
      Index           =   1
      Left            =   1320
      TabIndex        =   46
      Top             =   360
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   635
      picNormal       =   "frmMain.frx":3E4F3
      ToolTip         =   "Search Page - This is where you would search for files"
   End
   Begin VB.Label lblResizer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Resizer"
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   7920
      MousePointer    =   6  'Size NE SW
      TabIndex        =   47
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape shpGraph 
      Height          =   375
      Left            =   7920
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label status 
      Appearance      =   0  'Flat
      BackColor       =   &H008F7777&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   44
      Top             =   7080
      Width           =   3135
   End
   Begin VB.Shape shpStatus 
      BorderColor     =   &H007B6163&
      FillColor       =   &H008F7777&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7800
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LUSerNet (beta)"
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
      TabIndex        =   36
      Top             =   30
      Width           =   5175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.mvps.org/vbnet/index.html?code/shell/shchangenotify.htm

Option Explicit
Private WithEvents sizeSort As clsQuickSort
Attribute sizeSort.VB_VarHelpID = -1

Const dbName = "share.dat" '"LUSerNet.mdb"
'Const webMOTDURL = "http://www.psicentral.net/Edd/luser/prog_motd.php" '
Const webMOTDURL = "http://www.lusernet.34sp.com/motd"
'Const UploadLimit = 5

Private Type MyMessage
    msg As String
    time As Long
End Type

Public WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1

'Sets if the app can be resized
Const isResizable = False

'Min dimensions of the screen, measured in pixels
Const minX = 412
Const minY = 408

Private UserCount As Long 'a count of how many users found something

Public bShareRefreshing As Boolean 'indicates if the shares are being re-freshed
Public bClosing As Long 'varible indicating the programs state, 0-open, 1-closing but waiting, 2-closing now
Public bLeecherOn As Boolean
Private MsgArray() As MyMessage

Private Sub cmbSearch_Click()
    'On drop down click, this function grabs whats in the brackets, and removes rest
    Dim sStart As Long
    Dim sEnd As Long
    sStart = InStr(1, cmbSearch.Text, "(")
    sEnd = InStr(sStart + 1, cmbSearch.Text, ")")
    If sEnd <> 0 Then
        cmbSearch.Text = Mid(cmbSearch.Text, sStart + 1, sEnd - sStart - 1)
    End If
End Sub

Private Sub cmbSearch_GotFocus()
    If cmbSearch.Tag = 1 Then
        cmbSearch.Tag = 0
        cmbSearch.Text = ""
    End If
End Sub

Private Sub cmbSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearchGo_Click
        KeyAscii = 0
    End If
End Sub

Private Sub cmdClose_Click()
    Unload frmMain
End Sub

Private Sub cmdMinimise_Click()
    frmMain.WindowState = vbMinimized
End Sub

Private Sub cmdSearchGo_Click()

    'Easter Egg Code
    If cmdSearchGo.Tag <> "No" Then
        Select Case UCase(cmbSearch.Text)
            Case Is = "XXX", "PORN": msg "You Dirty Perv!": cmdSearchGo.Tag = "No"
            Case Is = "ANIMAL PORN": msg "You SICK SAD FUCK!": cmdSearchGo.Tag = "No"
        End Select
    End If
    Select Case UCase(cmbSearch.Text)
        Case Is = "BRAMP", "CHEKA", "GUYNEXTDOOR", "EDD", "EMIRI", "DAVE": msg cmbSearch.Text & " is " & RandomPlace()
    End Select

    'Normal Code
    Dim SearchText
    SearchText = Replace(cmbSearch.Text, "|", "")

    If Len(SearchText) <> 0 Then

        Dim i As Long
        Dim AllreadyAdded As Boolean

        AllreadyAdded = False
        For i = 5 To cmbSearch.ListCount
            If cmbSearch.Text = cmbSearch.List(i) Then AllreadyAdded = True
        Next i

        If Not AllreadyAdded Then
            'Add entry to combo
            If cmbSearch.ListCount >= 8 Then
                cmbSearch.RemoveItem 5
            End If
            cmbSearch.AddItem cmbSearch.Text
        End If

        lstMain.ListItems.Clear
        lstMain.ToolTipText = ""
        StatusChange "No Results"
        UserCount = 0
        'SendUDP "FIND|" & SearchText & "|" & GetIP, SubNetMask
        SendUDP "FIND|" & SearchText & "|", SubNetMask
    End If
End Sub

Private Sub cmdSelectDownloadLocation_Click()
    Dim Location As String
    Location = ShowBrowseFolders("Please Select The Folder You Wish To Download Files To")
    If Location <> "" Then
        txtDownloadLocation.Text = Location
        addShare Location
        SaveSettings
    End If
End Sub

Private Sub cmdShareAdd_Click()
    Dim newShare As String

    newShare = ShowBrowseFolders("Please Select The Folder You Wish To Share" & vbCrLf & "Selecting a large folder may appear to hang the program")
    If newShare <> "" Then
        addShare newShare
    End If
    
End Sub

Private Sub cmdShareRefresh_Click()

    If bShareRefreshing Then Exit Sub
    bShareRefreshing = True
    'Re-add everything
    'Loop all root folders
    Dim folders() As String
    Dim i As Long
    Set rs = conn.Execute("SELECT folderName FROM tblFolders WHERE rootID=-1 ORDER BY folderName ASC")
    i = 0
    Do While Not rs.EOF
        ReDim Preserve folders(i)
        folders(i) = rs("folderName")
        rs.MoveNext
        i = i + 1
    Loop
    rs.Close
    Set rs = conn.Execute("DELETE FROM tblFolders")
    Set rs = conn.Execute("DELETE FROM tblFiles")
    If i > 0 Then
        For i = 0 To UBound(folders)
            addShare folders(i)
        Next i
    End If

    Set rs = Nothing
    bShareRefreshing = False
    
    TryToClose
    
    'Redisplay everything
    ShareRefreshTree
End Sub

Private Sub cmdShareRemove_Click()
    If treeShares.SelectedItem.Image <> 2 Then
        msg "Please Select A Share"
    Else
        If treeShares.SelectedItem.Text = Left(txtDownloadLocation.Text, Len(txtDownloadLocation.Text) - 1) Then
            msg "Sorry can't remove your download location from your shares"
        Else
            'Remove selected Item
            Dim rootID As Long
            rootID = Right(treeShares.SelectedItem.Key, Len(treeShares.SelectedItem.Key) - 2)
            Dim folderID() As Long 'Holds the folderID of ALL sub folders
            Dim i As Long
            Dim ii As Long
    
            i = 0
            ii = 0
            ReDim Preserve folderID(i)
            folderID(i) = rootID
            Do
                'Loop all folders
                Set rs = conn.Execute("SELECT folderID, folderName FROM tblFolders WHERE rootID=" & rootID & " ORDER BY folderName ASC")
    
                Do While Not rs.EOF
                    i = i + 1
                    ReDim Preserve folderID(i)
                    folderID(i) = rs("folderID")
                    rs.MoveNext
                Loop
                rootID = folderID(ii)
                ii = ii + 1
                rs.Close
                Set rs = Nothing
            Loop While ii < i
    
            For i = 0 To UBound(folderID)
                Set rs = conn.Execute("DELETE FROM tblFolders WHERE folderID=" & folderID(i))
                Set rs = conn.Execute("DELETE FROM tblFiles WHERE folderID=" & folderID(i))
            Next i
    
            Set rs = Nothing
            ShareRefreshTree
        End If
    End If

    Set rs = Nothing

End Sub

Private Sub cmdTab_Click(index As Integer)
    Dim i
    For i = 0 To picFrame.UBound
        'picFrame(i).BackColor = RGB(181, 174, 206)
        picFrame(i).Left = 3
        picFrame(i).Top = 24 + 6 + 22
        picFrame(i).Enabled = False
        picFrame(i).Visible = False
        cmdTab(i).Top = 3 + 23
        cmdTab(i).picY = 0
        If i > 0 Then
            cmdTab(i).Left = cmdTab(i - 1).Left + cmdTab(i - 1).Width + 2
        Else
            cmdTab(i).Left = 11 + 2
        End If
    Next i
    picFrame(index).Enabled = True
    picFrame(index).Visible = True
    cmdTab(index).picY = 24

    Select Case index
        Case Is = 1: cmbSearch.SetFocus
        'Case Is = 5: About3D.StartPlaying
        'Case Else: About3D.StopPlaying
    End Select

End Sub

Private Sub Command1_Click()
    SendUDP "HI|abc|", SubNetMask
    MsgBox "Crash"
End Sub

Private Sub Download_Finished(index As Integer)
   
    UpdateTotalDownloads (Download(index).size)
   
    TryToClose

End Sub

Private Sub Download_Update(index As Integer)
    UpdateSpeeds
End Sub

Private Sub Form_Load()

    If App.PrevInstance Then
        ActivatePrevInstance
        End 'Add end here to make 100% sure
    End If

    'Checks if the files exists
    'If Not (FileExists(App.path & "\" & dbName) And FileExists(App.path & "\AutoUpdate.exe")) Then
    If Not FileExists(App.path & "\" & dbName) Then
        MsgBox "Some critical files are missing please re-install LUSerNet" & vbCrLf & "Perhaps you didn't update correct?" & vbCrLf & "Email: bramp@lusernet.tk for help", vbCritical
        End
    End If
       
    'Removed auto-update for a while
    'If Command = "" Then
    '    Shell App.path & "\AutoUpdate.exe"
    '    End
    'End If
    
    CompactTheDatabase (App.path & "\" & dbName)
    
    bClosing = 0
    ReDim MsgArray(0)

    Set sizeSort = New clsQuickSort

    Set gSysTray = New clsSysTray
    Set gSysTray.SourceWindow = Me
    gSysTray.ChangeIcon imgTrayIcon.ListImages(1).Picture

    'Loads all the DB stuff
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\" & dbName
    
    'Checks to see if the program crashed while using a IP
    Dim lastIP As String
    lastIP = getLastIP
    If isValidIP(lastIP) And lastIP <> "" Then
        conn.Execute "INSERT INTO tblBanned (bIP, bDate) VALUES ('" & lastIP & "', NOW())"
        msg "It appears that LUSerNet crashed while receiving data from:" & vbCrLf & lastIP & vbCrLf & "This addressed has now been blocked" & vbCrLf & "We are sorry for the inconvenience"
        setLastIP ""
    End If
    
    'Remove any banned IP that are older than ~2 days
    conn.Execute "DELETE FROM tblBanned WHERE bDate < (NOW()-0.005);"
    
    'Used to allow pings
    'SocketsInitialize

    'Determines the correct IP address
    Dim ipAddys() As String
    ipAddys = LocalIPAddresses

    'Dim ipAddysStr As String
    Dim i As Long
    
    'For i = LBound(ipAddys) To UBound(ipAddys)
    '    ipAddysStr = ipAddysStr & ipAddys(i) & vbCrLf
    'Next i

    'Old Connect Code

    'Dim IPRange As String

    'IPRange = Left(SubNetMask, InStr(SubNetMask, "255") - 1)

    'i = LBound(ipAddys)
    'Do While (Left(ipAddys(i), Len(IPRange)) <> IPRange And i < UBound(ipAddys)) Or (ipAddys(i) = "0.0.0.0")
    '    i = i + 1
    'Loop

    'If i = UBound(ipAddys) And i <> 0 Then
        'They are not in the correct range coz all IPs on their machine don't match
    '    msg "You don't appear to be in the correct IP range, your IPs are: " & vbCrLf & ipAddysStr & "You should be in the range " & SubNetMask
    'Else
        'If i = 0 Then ipAddys(0) = GetIP 'win98 hack
        
        'IPRange = Left(BlockedSubNetMask, InStr(BlockedSubNetMask, "255") - 1)
        'If IPRange = Left(ipAddys(i), Len(IPRange)) Then
            'They are blocked
        '    msg "Your Administrator has set a banned range of IPs and you appear to be in it." & vbCrLf & "Please contact your administrator to find out why."
        'Else
        '    UDPListen.Bind 9876, ipAddys(i)
        'End If
    
    'End If
    
    If UBound(ipAddys) = 0 Then ipAddys(0) = GetIP 'win98 hack
    
    For i = LBound(ipAddys) To UBound(ipAddys) - 1
        If i <> 0 Then
            If isValidIP(ipAddys(i)) Then
                Load UDPListen(i)
            End If
            'Load UDPSend(i)
            'UDPSend(i).Tag = ipAddys(i)
        End If
        If isValidIP(ipAddys(i)) Then
            'MsgBox UDPListen.UBound & " " & ipAddys(i)
            On Error GoTo Form_LoadBindError
            UDPListen(i).Close
            UDPListen(i).Bind 9876, ipAddys(i)
            On Error GoTo 0
            'MsgBox "done"
        End If
    Next i

    cmdTab_Click (0)

    If Command = "/log" Then
        LogOn = True
    End If

    LoadSettings
    
    'Set registry entrys
    SetRegValue HKEY_CLASSES_ROOT, ".LUSerNet", "", "Unfinished LUSerNet Download"
    
    'Create resizer handlers
    lblResizer(0).Caption = ""
    For i = 1 To 7
        Load lblResizer(i)
    Next i
    
    frmMain.Width = 6180 + 15 * 0
    frmMain.Height = 6000 + 15 * 8

    DoGUIStuff

    ShareRefreshTree

    frmMain.lstTransfers.Tag = 0
    tmrHello.Tag = 0
    SayHello

'    'Show Stats
'    Dim totalUpload As Single
'    totalUpload = GetSetting("LUSerNet", "Main", "TotalUpload", 0)
'
'    Dim fso As New FileSystemObject
'    Dim drv As Drive
'    Dim freeHD As Single
'    Dim totalHD As Single
'
'    freeHD = 0
'    totalHD = 0
'    On Error Resume Next
'    For i = Asc("A") To Asc("Z")
'        Set drv = Nothing
'        Set drv = fso.GetDrive(Chr(i))
'        If drv.DriveType = 2 Then
'            freeHD = freeHD + drv.FreeSpace
'            totalHD = totalHD + drv.TotalSize
'        End If
'    Next i
'    On Error GoTo 0
'    'STAT|TotalUpload|freeHD|totalHD|
'    SendUDP "STAT|" & totalUpload & "|" & freeHD & "|" & totalHD & "|"

    miniGraph.Start
    'miniGraph.Visible = False
    
    UpdateTotalDownloads 0
    UpdateTotalUploads 0
    
    'Sets up about page stuff
    'About3D.AddText lblTitle.Caption & vbCrLf & vbCrLf & "Coded For" & vbCrLf & "Lancs Uni Students" & vbCrLf & "By" & vbCrLf & "Lancs Uni Students"
    'About3D.AddText "Authors"
    'About3D.AddText "Coding & Design" & vbCrLf & "Bramp"
    'About3D.AddText "Graphics" & vbCrLf & "Emiri"
    'About3D.AddText "Website" & vbCrLf & "Barny, Bramp, Andy"
    'About3D.AddText "IRC Support" & vbCrLf & "Dave"
    'About3D.AddText "Other Helpers" & vbCrLf & "Edd"
    'About3D.AddText "Greets To" & vbCrLf & "Cheka" & vbCrLf & "GuyNextDoor" & vbCrLf & "Unbeleiver" & vbCrLf & "Carter"
    'About3D.AddText "Also Greets To" & vbCrLf & "3:30am Wonders" & vbCrLf & "&" & vbCrLf & "Fiery Badger Pr0n"
    'About3D.AddText "And Don't Forget"
    'About3D.AddText "ISS"
    
'    lblAbout.Caption = ""
'    lblAbout.Caption = lblAbout.Caption & lblTitle.Caption & vbCrLf & "Coded By Lancs Uni Students" & vbCrLf
'    lblAbout.Caption = lblAbout.Caption & "Authors" & vbCrLf
'    lblAbout.Caption = lblAbout.Caption & "Coding & Design: Bramp" & vbCrLf
'    lblAbout.Caption = lblAbout.Caption & "Graphics: Emiri" & vbCrLf
'    lblAbout.Caption = lblAbout.Caption & "Website: Barny, Bramp, Andy" & vbCrLf
'    lblAbout.Caption = lblAbout.Caption & "IRC Support: Dave" & vbCrLf
'    lblAbout.Caption = lblAbout.Caption & "Other Helpers Edd"

Exit Sub

Form_LoadBindError:
    MsgBox "Error " & Err.number & " Binding '" & ipAddys(i) & "'" & vbCrLf & "LUSerNet may or maynot continue to work" & vbCrLf & "Report this error to bramp@lusernet.tk"
Resume Next

End Sub

Public Sub ShareRefreshTree()

    Dim count As Long
    Dim Node As Node

    treeShares.Nodes.Clear

    Set Node = treeShares.Nodes.Add(, , "root", "Your Shares", 1)
    Node.Selected = True

    'Loop all root folders
    Set rs = conn.Execute("SELECT folderID, folderName FROM tblFolders WHERE rootID=-1 ORDER BY folderName ASC")

    count = 0
    Do While Not rs.EOF
        Set Node = treeShares.Nodes.Add("root", tvwChild, "FO" & rs("folderID"), rs("folderName"), 2)
        treeShares.Nodes.Add "FO" & rs("folderID"), tvwChild, "BFO" & rs("folderID"), "browsing...", 0
        Node.EnsureVisible
        rs.MoveNext
        count = count + 1
    Loop

    If count <> 0 Then rs.Close
    Set Node = Nothing

    Set rs = conn.Execute("SELECT COUNT(*) as Folders FROM tblFolders")
    lblShareFolders.Caption = "Folders: " & rs("Folders")
    lblShareFolders.Tag = rs("Folders")
    rs.Close
    
    Set rs = conn.Execute("SELECT COUNT(*) as Files FROM tblFiles")
    lblShareFiles.Caption = "Files: " & rs("Files")
    lblShareFiles.Tag = rs("Files")
    rs.Close
    
    Set rs = conn.Execute("SELECT SUM(FileSize) as Files FROM tblFiles")
    If Not IsNull(rs("Files")) Then
        lblShareSize.Caption = ChangeByte(rs("Files"), True, 2)
        lblShareSize.Tag = rs("Files")
    Else
        lblShareSize.Caption = "Nothing"
        lblShareSize.Tag = 0
    End If

    Set rs = Nothing
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormDrag Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    gSysTray.RemoveFromSysTray
End Sub

Private Sub Form_Resize()
    'Add or remove icon depending on Min or Max
    If Me.WindowState = vbMinimized Then
        gSysTray.MinToSysTray
    Else
        gSysTray.RemoveFromSysTray
    End If

    MakeRounded Me
    Skinner.Repaint Me
    If isResizable Then AddResizers Me
End Sub

Private Sub lblConnected_DblClick()
    'Double click on connected, to say hello
    SayHello
End Sub

Private Sub lblHomePage_Click()
    ShellExecute Me.hwnd, "open", "http://LUSerNet.tk", "", "", 0
End Sub

Private Sub lblHamor_DblClick()
    msg "Crack Heads Are Scary!! They Made Us Do NI!"
End Sub

'Private Sub lblGreets_Click(index As Integer)
'    Select Case index
'        Case Is = 3:
'            'Heil Hitler
'            'PlayWaveRes "Hitler"
'            ShellExecute Me.hwnd, "open", "http://www.psicentral.net/Edd/luser/gorilla.htm", "", "", 0
'    End Select
'End Sub

Private Sub lblMoreInfo_Click()
    ToggleMOTD
End Sub

Private Sub lblMOTD_DblClick()
    'reload MOTD
    txtMOTD.Text = "Loading..."
    tmrMOTD.Enabled = True
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormDrag Me
End Sub

Private Sub lstMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstMain.SortKey = ColumnHeader.index - 1
    If lstMain.SortOrder = lvwAscending Then lstMain.SortOrder = lvwDescending Else lstMain.SortOrder = lvwAscending
    If lstMain.SortKey = 1 Then
        sizeSort.Sort lstMain.ListItems.count, 1
    Else
        lstMain.Sorted = True
        lstMain.Sorted = False
    End If
End Sub

Private Sub lstMain_DblClick()
    'If Not isNothing(lstMain.HitTest(X, Y)) Then
    '    lstMain.HitTest(X, Y).Selected = True
    If Not isNothing(lstMain.SelectedItem) Then
        Call mnuDownload
    End If
End Sub

Private Sub lstMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        frmMenu.mnuDownload.Caption = "Download"
        If isNothing(lstMain.HitTest(x, y)) Then
            frmMenu.mnuDownload.Enabled = False
        Else
            lstMain.HitTest(x, y).Selected = True
            If FileExists(txtDownloadLocation.Text & lstMain.HitTest(x, y).Text & ".LUSerNet") Then
                frmMenu.mnuDownload.Caption = "Resume"
            Else
                frmMenu.mnuDownload.Caption = "Download"
            End If
            frmMenu.mnuDownload.Enabled = True
        End If
        PopupMenu frmMenu.mnuPopup
    End If
End Sub

'Changes status text for "length" seconds min
Sub StatusChange(message As String, Optional Length As Long = 0)
    'Dim tmpMsg As MyMessage
    
    'add message to end
    'tmpMsg.msg = message
    'tmpMsg.time = Length
    
    'ReDim Preserve MsgArray(UBound(MsgArray) + 1)
    
    'MsgArray(UBound(MsgArray)) = tmpMsg
    
    status.Caption = message
    frmMain.Caption = message
    If Not isNothing(gSysTray) Then gSysTray.ChangeToolTip message
End Sub

Private Sub lstTransfers_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If Not isNothing(lstTransfers.SelectedItem) Then
            Dim index As Long
            index = CLng(Right(lstTransfers.SelectedItem.Tag, Len(lstTransfers.SelectedItem.Tag) - 1))
            If Left(lstTransfers.SelectedItem.Tag, 1) = "U" Then
                If Not Upload(index).Sending Then mnuRemove
            Else
                If Not Download(index).Receiving Then mnuRemove
            End If
        End If
    End If
End Sub

Private Sub lstTransfers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If isNothing(lstTransfers.HitTest(x, y)) Then
            frmMenu.mnuStat.Visible = False
            frmMenu.mnuProgress.Visible = False
            frmMenu.mnuTimeLeft.Visible = False
            frmMenu.mnuGap.Visible = False
            frmMenu.mnuRemove.Visible = False
            frmMenu.mnuRemoveAll.Visible = True
            frmMenu.mnuStopDownload.Visible = False
            frmMenu.mnuOpenDownload.Visible = False
        Else
            lstTransfers.HitTest(x, y).Selected = True
            frmMenu.mnuRemove.Visible = True
            frmMenu.mnuRemoveAll.Visible = True
            frmMenu.mnuStopDownload.Visible = True
            frmMenu.mnuOpenDownload.Visible = False
            Dim index As Long
            index = CLng(Right(lstTransfers.HitTest(x, y).Tag, Len(lstTransfers.HitTest(x, y).Tag) - 1))
            If Left(lstTransfers.HitTest(x, y).Tag, 1) = "U" Then
                'Selected Transfer is Upload
                frmMenu.mnuProgress.Caption = "Progres: " & ChangeByte(Upload(index).position, True, 0) & "/" & ChangeByte(Upload(index).size, True, 0)
                frmMenu.mnuTimeLeft.Caption = "Timeleft: " & Upload(index).TimeLeft
                frmMenu.mnuStopDownload.Enabled = False
                If Upload(index).Sending Then
                    frmMenu.mnuRemove.Enabled = False
                    frmMenu.mnuStat.Visible = True
                    frmMenu.mnuProgress.Visible = True
                    frmMenu.mnuTimeLeft.Visible = True
                    frmMenu.mnuGap.Visible = True
                Else
                    frmMenu.mnuRemove.Enabled = True
                    frmMenu.mnuStat.Visible = False
                    frmMenu.mnuProgress.Visible = False
                    frmMenu.mnuTimeLeft.Visible = False
                    frmMenu.mnuGap.Visible = False
                End If

            Else
                'Selected Transfer is Download
                frmMenu.mnuProgress.Caption = "Progres: " & ChangeByte(Download(index).position, True, 0) & "/" & ChangeByte(Download(index).size, True, 0)
                frmMenu.mnuTimeLeft.Caption = "Timeleft: " & Download(index).TimeLeft
                If Download(index).Receiving Then
                    frmMenu.mnuRemove.Enabled = False
                    frmMenu.mnuStopDownload.Enabled = True
                    frmMenu.mnuStopDownload.Visible = True
                    frmMenu.mnuStat.Visible = True
                    frmMenu.mnuProgress.Visible = True
                    frmMenu.mnuTimeLeft.Visible = True
                    frmMenu.mnuGap.Visible = True

                Else
                    frmMenu.mnuRemove.Enabled = True
                    frmMenu.mnuStopDownload.Enabled = False
                    frmMenu.mnuStopDownload.Visible = False
                    frmMenu.mnuStat.Visible = False
                    frmMenu.mnuProgress.Visible = False
                    frmMenu.mnuTimeLeft.Visible = False
                    frmMenu.mnuGap.Visible = False

                    If Download(index).status = 4 Then
                        frmMenu.mnuOpenDownload.Visible = True
                    Else
                        frmMenu.mnuOpenDownload.Visible = False
                    End If
                End If
            End If
        End If
        PopupMenu frmMenu.mnuTransferPopup
    End If
End Sub

Private Sub tmrHello_Timer()
    tmrHello.Tag = tmrHello.Tag + 1

    'Every 20 minutes say hello 'each Hi is 32bytes, 100 Users saying Hi 100 times every 10 minutes = 0.5kb/s constant
    If (tmrHello.Tag Mod 20) = 0 Then
        SayHello
    End If

    'Every 30minutes refresh shares
    If tmrHello.Tag >= 60 Then
        cmdShareRefresh_Click
        tmrHello.Tag = 0
    End If
End Sub

Private Sub tmrMessage_Timer()
    tmrMessage.Enabled = False
    
    Dim message As String
    Dim i As Long
    
    If UBound(MsgArray) > 0 Then
        message = MsgArray(1).msg
        status.Caption = message
        frmMain.Caption = message
        If Not isNothing(gSysTray) Then gSysTray.ChangeToolTip message
        
        tmrMessage.Interval = MsgArray(1).time * 1000 + 1
        
        'Now move all items down one
        For i = 1 To UBound(MsgArray)
            MsgArray(i - 1) = MsgArray(i)
        Next i
        ReDim Preserve MsgArray(UBound(MsgArray) - 1)
    Else
        If status.Caption <> "" Then
            If tmrMessage.Interval = 3000 Then
                message = ""
                status.Caption = message
                frmMain.Caption = message
                If Not isNothing(gSysTray) Then gSysTray.ChangeToolTip message
                tmrMessage.Interval = 100
            Else
                If UBound(MsgArray) = 0 Then tmrMessage.Interval = 3000
            End If
        End If
        
    End If
    
    tmrMessage.Enabled = True
End Sub

Private Sub tmrMOTD_Timer()

    Dim strMOTD As String

    'On Error Resume Next
    On Error GoTo tmrMOTDError
    tmrMOTD.Enabled = False
    
    strMOTD = Trim(iNetMOTD.OpenURL(webMOTDURL))
    If Left(strMOTD, 1) = "<" Then
        txtMOTD.Text = "Unable to download MOTD"
    Else
        txtMOTD.Text = strMOTD
    End If
    
    Exit Sub
tmrMOTDError:
    txtMOTD.Text = "Unable to download MOTD"
End Sub

Public Sub UpdateSpeeds()
    Dim i As Long
    Dim USpeed As Single
    Dim DSpeed As Single

    For i = 1 To lstTransfers.ListItems.count
        If lstTransfers.ListItems(i).ListSubItems(3).Tag < 0 Then USpeed = USpeed + lstTransfers.ListItems(i).ListSubItems(3).Tag
        If lstTransfers.ListItems(i).ListSubItems(3).Tag > 0 Then DSpeed = DSpeed + lstTransfers.ListItems(i).ListSubItems(3).Tag
    Next i

    USpeed = Abs(USpeed)

    miniGraph.AddData CLng(USpeed / 1024), CLng(DSpeed / 1024)

    If (USpeed = 0) And (DSpeed = 0) Then
        If Left(status.Caption, 6) = "Upload" Then
            StatusChange "Upload: " & ChangeByte(USpeed, True, 0) & "/s Download: " & ChangeByte(DSpeed, True, 0) & "/s"
        End If
    Else
        StatusChange "Upload: " & ChangeByte(USpeed, True, 0) & "/s Download: " & ChangeByte(DSpeed, True, 0) & "/s"
    End If
End Sub

Private Sub tmrUpdate_Timer()

    On Error Resume Next

    tmrUpdate.Enabled = False
    
    'Updates Auto-Updater
    If FileExists(App.path & "\NewAutoUpdate.exe") Then
        Kill App.path & "\AutoUpdate.exe"
        Name App.path & "\NewAutoUpdate.exe" As App.path & "\AutoUpdate.exe"
    End If
    
    If lblShareFolders.Tag = 0 Then
        msg "Your are not sharing any folders, therefore your download folder has automatically been added"
        addShare txtDownloadLocation.Text
    End If
End Sub

Private Sub treeShares_Expand(ByVal Node As MSComctlLib.Node)
    If Left(Node.Child.Key, 1) = "B" Then
        'Populate
        treeShares.Nodes.Remove ("B" & Node.Key)

        Dim count As Long
        Dim tmpNode As Node
        Dim rootID As Long

        rootID = Right(Node.Key, Len(Node.Key) - 2)

        'Loop all folders
        Set rs = conn.Execute("SELECT folderID, folderName FROM tblFolders WHERE rootID=" & rootID & " ORDER BY folderName ASC")

        count = 0
        Do While Not rs.EOF
            Set tmpNode = treeShares.Nodes.Add(Node.Key, tvwChild, "FO" & rs("folderID"), rs("folderName"), 3)
            treeShares.Nodes.Add "FO" & rs("folderID"), tvwChild, "BFO" & rs("folderID"), "browsing...", 0
            rs.MoveNext
            count = count + 1
        Loop

        If count <> 0 Then rs.Close

        'Loop all Files
        Set rs = conn.Execute("SELECT folderID, fileName, fileID FROM tblFiles WHERE folderID=" & rootID & " ORDER BY fileName ASC")

        count = 0
        Do While Not rs.EOF
            Set tmpNode = treeShares.Nodes.Add(Node.Key, tvwChild, "G" & rs("fileID"), rs("fileName"), 4)

            rs.MoveNext
            count = count + 1
        Loop

        If count <> 0 Then rs.Close

        'tmpNode.EnsureVisible
        Set tmpNode = Nothing
        Set rs = Nothing
    End If
End Sub

Private Sub txtDownloadLocation_Change()
    If Right(txtDownloadLocation.Text, 1) <> "\" Then
        txtDownloadLocation.Text = txtDownloadLocation.Text & "\"
        SaveSettings
    End If
End Sub

Private Sub txtUploadTotal_Change()
    txtDownloadTotal.Text = txtUploadTotal.Text
End Sub


Private Sub UDPListen_DataArrival(idx As Integer, ByVal bytesTotal As Long)
    Dim Data As String
    Dim parts() As String
    Dim index As Long
    Dim remoteHost As String

    If bytesTotal >= 8192 Then
        GoTo badPacket
    End If
    
    If Not bClosing = 0 Then Exit Sub 'Invisible to network while closing

    UDPListen(idx).GetData Data
    remoteHost = UDPListen(idx).RemoteHostIP
    
    If isBadIP(remoteHost) Then
        GoTo badIP
    End If
    
    'Parse Data
    Hash Data
    'Log Data
    AddToLog remoteHost & " " & Data, True
    
    'Store the IP
    setLastIP remoteHost
    
    Clean Data 'Clean Data (any bad chars)

    If InStr(Data, "||") <> 0 Then
        GoTo badPacket
    End If

    parts = Split(Data, "|")
    
    If UBound(parts) = -1 Then
        GoTo badPacket
    End If

    If Not isMe(remoteHost) Then 'Make sure they ain't local events
    If lblConnected.Tag = "1" Or parts(0) = "HI" Or parts(0) = "HELLO" Then
        
        If parts(0) = "HI" And UBound(parts) >= 4 Then
        
            'Checks they are numbers
            If Not (isNumPositiveLg(parts(2)) And isNumPositiveLg(parts(3)) And isNumPositiveDb(parts(4))) Then
                GoTo badPacket
            End If
        
            On Error Resume Next
            lblTotalUsers.Caption = CLng(lblTotalUsers.Caption) + 1
            lblTotalFiles.Caption = CLng(lblTotalFiles.Caption) + CLng(parts(2))
            lblTotalFolders.Caption = CLng(lblTotalFolders.Caption) + CLng(parts(3))
            lblTotalSize.Tag = CDbl(lblTotalSize.Tag) + CDbl(parts(4))
            lblTotalSize.Caption = ChangeByte(lblTotalSize.Tag)
            lblConnected.Caption = "Connected"
            lblConnected.Tag = "1"
            On Error GoTo 0
        
        ElseIf parts(0) = "HELLO" And UBound(parts) >= 0 Then
        
             'Something dodgy is going on
            'If remoteHost <> parts(UBound(parts)) Then
            '    GoTo badPacket
            'End If
            
            'If Not isValidIP(parts(1)) Then GoTo badPacket
            
            'Someone is saying Hello, send hi back with filesharing stats
            'HI|Version|Files|Folders|Size|
            SendUDP "HI|" & App.Major & "." & App.Minor & "." & App.Revision & "|" & lblShareFiles.Tag & "|" & lblShareFolders.Tag & "|" & lblShareSize.Tag & "|", remoteHost
            'SendUDP "HI|" & App.Major & "." & App.Minor & "." & App.Revision & "|" & lblShareFiles.Tag & "|" & lblShareFolders.Tag & "|" & lblShareSize.Tag & "|", SubNetMask
       
        
        ElseIf parts(0) = "FIND" And UBound(parts) >= 1 Then
            If CountUploads < txtUploadTotal.Text Then
                'Only does this if the program isn't closing
    
                 'Something dodgy is going on
                'If remoteHost <> parts(UBound(parts)) Then
                '    GoTo badPacket
                'End If
                
                'If Not isValidIP(parts(2)) Then GoTo badPacket
    
                'Do a Search and return results
                Dim resultsParts() As String
                Dim Results() As String
                ReDim Preserve Results(0)
                Results(0) = SearchFile(parts(1))
    
                Dim limit As Long
                limit = (8000 - Len("FOUND|") - Len(GetIP))
                Dim i As Long
                Dim ii As Long
                Dim tmp As String
    
                'If Len(Results(0)) >= limit Then
                'If the results are too big, then chop them up
                resultsParts = Split(Results(0), "|")
    
                ii = 0
                Results(ii) = ""
                For i = 0 To UBound(resultsParts) - 1 Step 2
                    tmp = resultsParts(i) & "|" & resultsParts(i + 1) & "|"
                    If Len(Results(ii)) + Len(tmp) > limit Then
                        ii = ii + 1
                        ReDim Preserve Results(ii)
                    End If
                    Results(ii) = Results(ii) & tmp
                Next i
    
                For i = 0 To UBound(Results)
                    If Results(i) <> "" Then
                        'SendUDP "FOUND|" & Results(i) & GetIP, parts(UBound(parts))
                        SendUDP "FOUND|" & Results(i), remoteHost
                    End If
                Next i

            End If

        ElseIf parts(0) = "FOUND" And UBound(parts) >= 2 And (UBound(parts) Mod 2) = 1 Then
            
             'Something dodgy is going on
            'If remoteHost <> parts(UBound(parts)) Then
            '    GoTo badPacket
            'End If
            
            'Display results
            Dim tmpItem As ListItem

            UserCount = UserCount + 1

            i = 1
            Do While i < UBound(parts)
            
                If parts(i) <> "" And isNumPositiveLg(parts(i + 1)) Then
            
                    Set tmpItem = lstMain.ListItems.Add(, , parts(i), , FileTypeToImage(parts(i)))
                    tmpItem.Tag = remoteHost
                    tmpItem.ListSubItems.Add , , ChangeByte(CLng(parts(i + 1)), True)
                    tmpItem.ListSubItems(1).Tag = parts(i + 1)
                    
                    'Fastest, Fast, Normal, Slow, Slowest
                    tmpItem.ListSubItems.Add , , getSpeed(UserCount)
                    
                    'tmpItem.ListSubItems(2).ForeColor = RGB(0, 255, 0)
                Else
                    GoTo badPacket
                End If
                
                i = i + 2
                StatusChange "Found " & lstMain.ListItems.count & " items"
            Loop

        ElseIf parts(0) = "GET" And UBound(parts) >= 3 Then
            'File Request, open port and respond with what port is open, then send file
            'GET|Filename|size|position|ip

             'Something dodgy is going on
            'If remoteHost <> parts(UBound(parts)) Then
            '    GoTo badPacket
            'End If
            
            If Not (isNumPositiveLg(parts(2)) And isNumPositiveLg(parts(3))) Then GoTo badPacket
            
            If CLng(parts(2)) < CLng(parts(3)) Then GoTo badPacket 'trying to resume from past the end of the file, maybe a better error msg is needed
            
            If CountUploads < txtUploadTotal.Text Then

                Dim Error404 As Boolean
    
                Dim filename As String
                
                'Chops off directory names
                parts(1) = Right(parts(1), Len(parts(1)) - InStrRev(parts(1), "\"))
                
                filename = parts(1)
                parts(1) = Replace(parts(1), "'", "''")
    
                'Search for file in database
                Set rs = conn.Execute("SELECT folderID FROM tblFiles WHERE fileName='" & parts(1) & "' AND fileSize=" & parts(2))
                If Not rs.EOF Then
    
                    parts(1) = GetPath(rs("folderID")) & filename
    
                    index = Upload.UBound + 1
                    Load Upload(index)
    
                    Upload(index).Listen remoteHost, Rnd * 100 + 9876, parts(1), parts(3) + 1
                    lstTransfers.ListItems("I" & Upload(index).transferID).Tag = "U" & index
    
                    If FileExists(parts(1)) Then
                        rs.Close
                        Upload(index).status = 1
    
                        'Send OK reply
                        'SendUDP "OK|" & filename & "|" & FileLen(parts(1)) & "|" & parts(3) & "|" & Upload(Index).LocalPort & "|" & GetIP, parts(UBound(parts))
                        SendUDP "OK|" & filename & "|" & FileLen(parts(1)) & "|" & parts(3) & "|" & Upload(index).LocalPort & "|", remoteHost
                    Else
                        Error404 = True
                    End If
                Else
                    Error404 = True
    
                End If
    
                Set rs = Nothing
    
                If Error404 And filename <> "" Then
                    'Error file not shared
                    SendUDP "ERROR|" & filename & "|404", remoteHost
                    'unload tcp stuff
                    Upload(index).EndListening 8
                End If
            
            Else 'Reached Upload Limit
                SendUDP "ERROR|" & parts(1) & "|502", remoteHost
                'unload tcp stuff
                Upload(index).EndListening 10
            End If
            
        ElseIf parts(0) = "SEND" Then
            'File Request, send file to specific port/ip
            'SEND|Filename|size|position|port|ip

        ElseIf parts(0) = "OK" And UBound(parts) >= 4 Then
            'OK|Filename|size|position|port|ip
             
             'Something dodgy is going on
            'If remoteHost <> parts(UBound(parts)) Then
            '    GoTo badPacket
            'End If
            
            'Checks they are numbers
            If Not (isNumPositiveLg(parts(2)) And isNumPositiveLg(parts(3)) And isNumPositiveLg(parts(4))) Then
                GoTo badPacket
            End If
            
            'Checks port validiatly
            If Not (parts(4) >= 9876 And parts(4) <= 9977) Then
                GoTo badPacket
            End If
            
            'Someone says u can get file from them, go connect to port/ip
            'OK|Filename|size|position|port|ip
            index = Download.UBound + 1
            Load Download(index)
            Download(index).Start remoteHost, CLng(parts(4)), parts(1), CLng(parts(3)), CLng(parts(2))
            lstTransfers.ListItems("I" & Download(index).transferID).Tag = "D" & index

        ElseIf parts(0) = "ERROR" And UBound(parts) >= 1 Then
        
            'Error occured with file
            'ERROR|Filename|Error
            index = Download.UBound + 1
            Load Download(index)
            Download(index).Start remoteHost, 0, parts(1), 0, 0, True
            lstTransfers.ListItems("I" & Download(index).transferID).Tag = "D" & index
            Select Case parts(2)
                Case Is = "404": Download(index).EndTransfer 8
                Case Is = "502": Download(index).EndTransfer 10
                Case Else: Download(index).EndTransfer 99
            End Select
        ElseIf parts(0) = "STAT" Then
            'Do Nothing
        Else
            GoTo badPacket
        End If 'Ends packet selection IF
        
    Else
        GoTo badPacket
    End If
    End If

setLastIP ""
Exit Sub

badPacket:
AddToLog "Bad Packet (" & remoteHost & ")", True
setLastIP ""
Exit Sub

badIP:
AddToLog "Banned IP  (" & remoteHost & ")", True
setLastIP ""

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If CountUploads + CountDownloads > 0 Then
        frmRUSure.Show 1, frmMain
    Else
        bClosing = 2
    End If

    If bClosing = 2 Then
        'Save Settings
        SaveSettings

        Dim i As Long
        On Error Resume Next

        For i = Upload.LBound To Upload.UBound
            Unload Upload(i)
        Next i

        For i = Download.LBound To Download.UBound
            Unload Download(i)
        Next i
        
        For i = 1 To UDPListen.UBound
            Unload UDPListen(i)
            'Unload UDPSend(i)
        Next i
        
        setLastIP ""
        
        If Not bShareRefreshing Then

            'Clean up ping stuff
            'SocketsCleanup
    
            Unload frmMenu
    
            'Close all database
            conn.Close
            Set conn = Nothing
            rs.Close
            Set rs = Nothing
            Set sizeSort = Nothing
            CompactTheDatabase (App.path & "\" & dbName)
            
            CloseLogFile
            
            End
        Else
            Cancel = 1
            lblConnected.Caption = "Not Connected"
            msg "Your shares are refreshing, once finished the program will exit"
            bClosing = 1
        End If
    Else
        Cancel = 1
    End If
End Sub

Private Sub sizeSort_isLess(ByVal ndx1 As Long, ByVal ndx2 As Long, result As Integer)

    'Result = StrComp(Items(ndx1), Items(ndx2))

    Dim item1 As Long
    Dim item2 As Long

    item1 = lstMain.ListItems(ndx1).ListSubItems(1).Tag
    item2 = lstMain.ListItems(ndx2).ListSubItems(1).Tag

    If lstMain.SortOrder = lvwAscending Then
        If item1 = item2 Then
            result = 0
        ElseIf item1 < item2 Then
            result = -1
        Else
            result = 1
        End If
    Else
        If item1 = item2 Then
            result = 0
        ElseIf item1 > item2 Then
            result = -1
        Else
            result = 1
        End If
    End If
End Sub

Private Sub sizeSort_SwapItems(ByVal ndx1 As Long, ByVal ndx2 As Long)
    'All this just to swap 2 items...
    Dim tmp As ListItem
    Dim i As Long

    Set tmp = lstTMP.ListItems.Add(, "TEMP", "")
    tmp.ListSubItems.Add , , ""
    tmp.ListSubItems.Add , , ""
    tmp.ListSubItems.Add , , ""

    'Set tmp = lstMain.ListItems(ndx1)
    tmp.Text = lstMain.ListItems(ndx1).Text
    tmp.Tag = lstMain.ListItems(ndx1).Tag
    tmp.SmallIcon = lstMain.ListItems(ndx1).SmallIcon

    For i = 1 To lstMain.ListItems(ndx1).ListSubItems.count
        tmp.ListSubItems(i).Text = lstMain.ListItems(ndx1).ListSubItems(i).Text
        tmp.ListSubItems(i).Tag = lstMain.ListItems(ndx1).ListSubItems(i).Tag
    Next i

    'Set lstMain.ListItems(ndx1) = lstMain.ListItems(ndx2)
    lstMain.ListItems(ndx1).Text = lstMain.ListItems(ndx2).Text
    lstMain.ListItems(ndx1).Tag = lstMain.ListItems(ndx2).Tag
    lstMain.ListItems(ndx1).SmallIcon = lstMain.ListItems(ndx2).SmallIcon

    For i = 1 To lstMain.ListItems(ndx1).ListSubItems.count
        lstMain.ListItems(ndx1).ListSubItems(i).Text = lstMain.ListItems(ndx2).ListSubItems(i).Text
        lstMain.ListItems(ndx1).ListSubItems(i).Tag = lstMain.ListItems(ndx2).ListSubItems(i).Tag
    Next i

    'Set lstMain.ListItems(ndx2) = tmp
    lstMain.ListItems(ndx2).Text = tmp.Text
    lstMain.ListItems(ndx2).Tag = tmp.Tag
    lstMain.ListItems(ndx2).SmallIcon = tmp.SmallIcon

    For i = 1 To lstMain.ListItems(ndx2).ListSubItems.count
        lstMain.ListItems(ndx2).ListSubItems(i).Text = tmp.ListSubItems(i).Text
        lstMain.ListItems(ndx2).ListSubItems(i).Tag = tmp.ListSubItems(i).Tag
    Next i
    Set tmp = Nothing
    lstTMP.ListItems.Remove "TEMP"
End Sub

Private Sub ToggleMOTD()
    If picLogo.Visible = True Then
        'Make MOTD big
        picLogo.Visible = False
        shpMOTD.Top = 64
        lblMOTD.Top = 56
        shpMOTD.Height = 249
        'txtMOTD.MultiLine = True
        txtMOTD.Top = 74
        txtMOTD.Height = 230
    Else
        'Make it small
        picLogo.Visible = True
        shpMOTD.Top = 280
        lblMOTD.Top = 272
        shpMOTD.Height = 33
        'txtMOTD.MultiLine = False
        txtMOTD.Top = 290
        txtMOTD.Height = 17
    End If
End Sub


Sub mnuDownload()
    'GET|Enterprise Alternate Intro.mpg|14747132|0|127.0.0.1
    If CountDownloads < txtUploadTotal.Text Then
        Dim Start As Long
        On Error Resume Next
        Start = FileLen(frmMain.txtDownloadLocation.Text & lstMain.SelectedItem.Text & DownloadEXT) - 8192
        'Now work it out to the nearest bufferSize
        Start = (Start \ bufferSize) * bufferSize
        On Error GoTo 0
        If Start < 0 Then Start = 0
        'SendUDP "GET|" & lstMain.SelectedItem.Text & "|" & lstMain.SelectedItem.ListSubItems(1).Tag & "|" & Start & "|" & GetIP, lstMain.SelectedItem.Tag
        SendUDP "GET|" & lstMain.SelectedItem.Text & "|" & lstMain.SelectedItem.ListSubItems(1).Tag & "|" & Start & "|", lstMain.SelectedItem.Tag
    Else
        msg "Sorry you are allready downloading " & CountDownloads & " files, either wait" & vbCrLf & "or increase your download limit on the transfer page"
    End If
End Sub

Private Function CountUploads() As Long

    Dim i As Long
    Dim count As Long
    count = 0
    For i = Upload.LBound To Upload.UBound
        If Not isNothing(Upload(i)) Then
            If Upload(i).Sending Then count = count + 1
        End If
    Next i

    CountUploads = count

End Function

Private Function CountDownloads() As Long

    Dim i As Long
    Dim count As Long
    count = 0
    For i = Download.LBound To Download.UBound
        If Not isNothing(Download(i)) Then
            If Download(i).Receiving Then count = count + 1
        End If
    Next i

    CountDownloads = count

End Function

Public Sub mnuRemove()
    Dim index As Long
    index = Right(lstTransfers.SelectedItem.Tag, Len(lstTransfers.SelectedItem.Tag) - 1)
    If Left(lstTransfers.SelectedItem.Tag, 1) = "U" Then
        Unload Upload(index)
    Else
        Unload Download(index)
    End If
End Sub

Public Sub mnuRemoveAll()
    Dim i As Long
    For i = Download.LBound To Download.UBound
        If i <> 0 Then
            If Not isNothing(Download(i)) Then
                If Not Download(i).Receiving Then Unload Download(i)
            End If
        End If
    Next i

    For i = Upload.LBound To Upload.UBound
        If i <> 0 Then
            If Not isNothing(Upload(i)) Then
                If Not Upload(i).Sending Then Unload Upload(i)
            End If
        End If
    Next i
End Sub

Public Sub mnuStopDownload()
    Dim index As Long
    index = Right(lstTransfers.SelectedItem.Tag, Len(lstTransfers.SelectedItem.Tag) - 1)
    Download(index).EndTransfer 9
End Sub

Public Sub mnuOpenDownload()
    Dim index As Long
    index = Right(lstTransfers.SelectedItem.Tag, Len(lstTransfers.SelectedItem.Tag) - 1)
    ShellExecute Me.hwnd, "open", txtDownloadLocation.Text & Download(index).filename, "", "", 1
End Sub


Private Sub UpDown_Change(index As Integer)
    
    If UpDown(index).Value = 0 Then UpDown(index).Value = 5
    
    If CountDownloads > UpDown(index).Value Then UpDown(index).Value = CountDownloads
    
    txtUploadTotal.Text = UpDown(index).Value
    txtDownloadTotal.Text = UpDown(index).Value
    
    UpDown(1 - index).Value = UpDown(index).Value
    
End Sub

Private Sub TryToClose()
    If bClosing = 1 And CountUploads + CountDownloads = 0 And Not bShareRefreshing Then
        Unload Me
    End If
End Sub

Private Sub Upload_Finished(index As Integer)

    UpdateTotalUploads (Upload(index).size)
    
    TryToClose
End Sub

Private Sub Upload_Update(index As Integer)
    UpdateSpeeds
End Sub

Private Sub lblResizer_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If isResizable And Button = 1 Then Resizer Me, index, x, y, minX, minY
End Sub

Function isMe(IP As String) As Boolean
    Dim found As Boolean
    Dim i As Long
    i = 0
    found = False
    Do While Not found And i <= UDPListen.UBound
        found = (IP = UDPListen(i).LocalIP)
        i = i + 1
    Loop
    isMe = found
End Function
