VERSION 5.00
Begin VB.Form frmSkin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   708
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   735
      Left            =   2040
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "Reload Skin"
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin Project1.MySkinner Skinner 
      Left            =   1440
      Top             =   360
      _ExtentX        =   1693
      _ExtentY        =   635
      imgNE           =   "frmSkin.frx":0000
      imgN            =   "frmSkin.frx":0166
      imgNW           =   "frmSkin.frx":02CC
      imgE            =   "frmSkin.frx":0882
      imgW            =   "frmSkin.frx":0904
      imgSE           =   "frmSkin.frx":0986
      imgS            =   "frmSkin.frx":0A08
      imgSW           =   "frmSkin.frx":0A8A
      BackColor       =   13545141
   End
   Begin VB.Label lblResizer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   0
      Left            =   960
      MousePointer    =   6  'Size NE SW
      TabIndex        =   0
      Top             =   3480
      Width           =   4575
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Min dimensions of the screen, measured in pixels
Const minX = 200
Const minY = 200

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdReload_Click()
    Set Skinner.imgN = LoadPicture("n.bmp")
    Set Skinner.imgE = LoadPicture("e.bmp")
    Set Skinner.imgS = LoadPicture("s.bmp")
    Set Skinner.imgW = LoadPicture("w.bmp")
    Set Skinner.imgNE = LoadPicture("ne.bmp")
    Set Skinner.imgNW = LoadPicture("nw.bmp")
    Set Skinner.imgSE = LoadPicture("se.bmp")
    Set Skinner.imgSW = LoadPicture("sw.bmp")
End Sub

Private Sub Form_Load()
    
    Dim i As Long

    For i = 1 To 7
        Load lblResizer(i)
    Next i
    
End Sub

Private Sub AddResizers(frm As Form)
    Dim handleSizeX As Long
    Dim handleSizeY As Long
    
    handleSizeX = 4 '* Screen.TwipsPerPixelX
    handleSizeY = 4 '* Screen.TwipsPerPixelY
    
    '0=N, 1=E, 2=S, 3=W
    '4=NE, 5=SE, 6=SW, 7=NW
    
    'North
    lblResizer(0).Move handleSizeX, 0, frm.ScaleWidth - 2 * handleSizeX, handleSizeY
    lblResizer(0).Visible = True
    lblResizer(0).MousePointer = 7
    
    'East
    lblResizer(1).Move frm.ScaleWidth - handleSizeX, handleSizeY, handleSizeX, frm.ScaleHeight - 2 * handleSizeY
    lblResizer(1).Visible = True
    lblResizer(1).MousePointer = 9
    
    'South
    lblResizer(2).Move handleSizeX, frm.ScaleHeight - handleSizeY, frm.ScaleWidth - 2 * handleSizeX, handleSizeY
    lblResizer(2).Visible = True
    lblResizer(2).MousePointer = 7
    
    'West
    lblResizer(3).Move 0, handleSizeY, handleSizeX, frm.ScaleHeight - 2 * handleSizeY ' .Left = 0
    lblResizer(3).Visible = True
    lblResizer(3).MousePointer = 9
    
    'North East
    lblResizer(4).Move frm.ScaleWidth - handleSizeX, 0, handleSizeX, handleSizeY
    lblResizer(4).Visible = True
    lblResizer(4).MousePointer = 6
    
    'South East
    lblResizer(5).Move frm.ScaleWidth - handleSizeX, frm.ScaleHeight - handleSizeY, handleSizeX, handleSizeY
    lblResizer(5).Visible = True
    lblResizer(5).MousePointer = 8
    
    'South West
    lblResizer(6).Move 0, frm.ScaleHeight - handleSizeY, handleSizeX, handleSizeY
    lblResizer(6).Visible = True
    lblResizer(6).MousePointer = 6
    
    'North West
    lblResizer(7).Move 0, 0, handleSizeX, handleSizeY
    lblResizer(7).Visible = True
    lblResizer(7).MousePointer = 8
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Form_Resize()
    Skinner.Repaint frmSkin
    AddResizers Me
End Sub

Private Sub lblResizer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim newX As Long
    Dim newY As Long
    
    'Decides which directions are being sized
    Dim H1 As Boolean '<-
    Dim V1 As Boolean '^
    Dim H2 As Boolean '->
    Dim V2 As Boolean '!^
    
    H1 = False
    V1 = False
    H2 = False
    V2 = False
    
    Select Case Index
        Case Is = 0: V1 = True
        Case Is = 1: H2 = True
        Case Is = 2: V2 = True
        Case Is = 3: H1 = True
        Case Is = 4: V1 = True: H2 = True
        Case Is = 5: V2 = True: H2 = True
        Case Is = 6: V2 = True: H1 = True
        Case Is = 7: V1 = True: H1 = True
    End Select

    If Button = 1 Then
        If H2 Then
            newX = Me.Width + X
            If (newX < minX * Screen.TwipsPerPixelX) Then newX = minX * Screen.TwipsPerPixelX
            Me.Width = newX
        ElseIf H1 Then
            newX = Me.Width - X
            If (newX < minX * Screen.TwipsPerPixelX) Then
                newX = minX * Screen.TwipsPerPixelX
                X = Me.Width - newX
            End If
            Me.Width = newX
            Me.Left = Me.Left + X
        End If
        
        If V2 Then
            newY = Me.Height + Y
            If (newY < minY * Screen.TwipsPerPixelY) Then newY = minY * Screen.TwipsPerPixelY
            Me.Height = newY
        ElseIf V1 Then
            newY = Me.Height - Y
            If (newY < minY * Screen.TwipsPerPixelY) Then
                newY = minY * Screen.TwipsPerPixelY
                Y = Me.Height - newY
            End If
            Me.Height = newY
            Me.Top = Me.Top + Y
        End If
    End If
End Sub
