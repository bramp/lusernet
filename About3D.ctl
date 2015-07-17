VERSION 5.00
Begin VB.UserControl About3D 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00290000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ForeColor       =   &H00F7D8D8&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer tmrSwapMessage 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   2520
   End
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1680
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer tmrFadeOut 
      Enabled         =   0   'False
      Left            =   600
      Top             =   3000
   End
   Begin VB.Timer tmrFlyIn 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3000
   End
End
Attribute VB_Name = "About3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Event Declarations:
Event TextDone()
Event AllDone()

Private Type Pixel
    x As Single
    y As Single
End Type

'Holds all the messages to display
Private MessageArray() As String
Private CurrentMessage As Long

Private FinalStartTextArray() As Pixel
Private StartTextArray() As Pixel
Private LastArray() As Pixel
Private DeltaArray() As Pixel
Private TextCenter As Pixel

Private Moves As Long
'Default Property Values:
Const m_def_TotalMoves = 0
'Property Variables:
Dim m_TotalMoves As Long

Dim btmrFlyIn As Boolean
Dim btmrSwapMessage As Boolean

Dim bStarted As Boolean
Dim bRunning As Boolean

Public Sub StopPlaying()
If bRunning Then
    btmrFlyIn = tmrFlyIn.Enabled
    btmrSwapMessage = tmrSwapMessage.Enabled
    tmrFlyIn.Enabled = False
    tmrSwapMessage.Enabled = False
    
    bRunning = False
End If
End Sub

Public Sub StartPlaying()
If Not bRunning Then
    If CurrentMessage = 0 Then CurrentMessage = 1
    If bStarted = False Then
        bStarted = True
        tmrSwapMessage.Enabled = True
    Else
        tmrSwapMessage.Enabled = btmrSwapMessage
    End If
    tmrFlyIn.Enabled = btmrFlyIn
    
    bRunning = True
End If
End Sub


Public Sub AddText(Text As String)

tmrSwapMessage.Enabled = True

ReDim Preserve MessageArray(UBound(MessageArray()) + 1)
MessageArray(UBound(MessageArray())) = Text

End Sub

Private Sub tmrFlyIn_Timer()

Dim i As Long

'UserControl.Cls

'Plot them all
For i = 1 To UBound(StartTextArray())
    'Remove old one
    UserControl.PSet (LastArray(i).x, LastArray(i).y), UserControl.BackColor
    'Adds new one
    LastArray(i).x = StartTextArray(i).x + TextCenter.x
    LastArray(i).y = StartTextArray(i).y + TextCenter.y
    UserControl.PSet (LastArray(i).x, LastArray(i).y), UserControl.ForeColor
    
    StartTextArray(i).x = StartTextArray(i).x + DeltaArray(i).x
    StartTextArray(i).y = StartTextArray(i).y + DeltaArray(i).y
Next i

If Moves = m_TotalMoves Then

    'Draws new one again
    For i = 1 To UBound(FinalStartTextArray())
        'Adds new one
        UserControl.PSet (FinalStartTextArray(i).x + TextCenter.x, FinalStartTextArray(i).y + TextCenter.y), UserControl.ForeColor
    Next i

    tmrFlyIn.Enabled = False
    RaiseEvent TextDone
    CurrentMessage = CurrentMessage + 1
    tmrSwapMessage.Enabled = True
    If CurrentMessage > UBound(MessageArray) Then
        CurrentMessage = 1
        RaiseEvent AllDone
    End If
End If

Moves = Moves + 1
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get TotalMoves() As Long
    TotalMoves = m_TotalMoves
End Property

Public Property Let TotalMoves(ByVal New_TotalMoves As Long)
    m_TotalMoves = New_TotalMoves
    PropertyChanged "TotalMoves"
End Property

Private Sub tmrSwapMessage_Timer()

tmrSwapMessage.Enabled = False

Dim i As Long
Dim Text As String

Text = MessageArray(CurrentMessage)

Dim fadeColor As Long
Dim fadeDepth As Long

Dim ii As Long

'UserControl.AutoRedraw = False

If UBound(FinalStartTextArray()) <> 0 Then
    fadeDepth = 50000 / UBound(FinalStartTextArray()) 'Makes sure everything fades at same speed
    For ii = 1 To fadeDepth
        fadeColor = GetFadedColor(UserControl.ForeColor, UserControl.BackColor, ii, fadeDepth)
        For i = 1 To UBound(FinalStartTextArray())
            UserControl.PSet (FinalStartTextArray(i).x + TextCenter.x, FinalStartTextArray(i).y + TextCenter.y), fadeColor
        Next i
        UserControl.Refresh
    Next ii
End If

'Clears to make sure
UserControl.Cls

'UserControl.AutoRedraw = True

ReDim StartTextArray(0)
ReDim FinalStartTextArray(0)
ReDim LastArray(0)

'Get the pixel array for the text
picText.Width = picText.TextWidth(Text)
picText.Height = picText.TextHeight(Text)

TextCenter.x = (UserControl.ScaleWidth - picText.Width) \ 2
TextCenter.y = (UserControl.ScaleHeight - picText.Height) \ 2

picText.Cls
picText.Print Text

Dim x As Single
Dim y As Single

'Makes the actual array 'change to step 2 for funky effects
For x = 0 To picText.Width Step 3
    For y = 0 To picText.Height Step 3
        If picText.Point(x, y) = vbBlack Then
            ReDim Preserve StartTextArray(UBound(StartTextArray()) + 1)
            StartTextArray(UBound(StartTextArray())).x = x
            StartTextArray(UBound(StartTextArray())).y = y
        End If
    Next y
Next x

FinalStartTextArray = StartTextArray
LastArray = StartTextArray
DeltaArray = StartTextArray

'Now scramble the StartTextArray
For i = 1 To UBound(StartTextArray())
    StartTextArray(i).x = Rnd * UserControl.ScaleWidth - TextCenter.x
    StartTextArray(i).y = Rnd * UserControl.ScaleHeight - TextCenter.y
    
    DeltaArray(i).x = (FinalStartTextArray(i).x - StartTextArray(i).x) / m_TotalMoves
    DeltaArray(i).y = (FinalStartTextArray(i).y - StartTextArray(i).y) / m_TotalMoves
Next i

Moves = 0
tmrFlyIn.Enabled = True
End Sub

Private Sub UserControl_Initialize()
    ReDim StartTextArray(0)
    ReDim FinalStartTextArray(0)
    
    ReDim MessageArray(0)
    CurrentMessage = 0
    
    Randomize
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_TotalMoves = m_def_TotalMoves
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_TotalMoves = PropBag.ReadProperty("TotalMoves", m_def_TotalMoves)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H290000)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &HF7D8D8)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TotalMoves", m_TotalMoves, m_def_TotalMoves)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H290000)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &HF7D8D8)
End Sub

Public Function GetFadedColor(c1 As Long, c2 As Long, FN As Long, FS As Long) As Long
    Dim i&, red1%, green1%, blue1%, red2%, green2%, blue2%, pat1!, pat2!, pat3!, cx1!, cx2!, cx3!
    
    ' get the red, green, and blue values out of the different
    ' colors
    red1% = (c1 And 255)
    green1% = (c1 \ 256 And 255)
    blue1% = (c1 \ 65536 And 255)
    red2% = (c2 And 255)
    green2% = (c2 \ 256 And 255)
    blue2% = (c2 \ 65536 And 255)
    
    ' get the step of the color changing
    pat1 = (red2% - red1%) / FS
    pat2 = (green2% - green1%) / FS
    pat3 = (blue2% - blue1%) / FS

    ' set the cx variables at the starting colors
    cx1 = red1%
    cx2 = green1%
    cx3 = blue1%

    ' loop till you reach the faze you are at in the fading
    For i& = 1 To FN
        cx1 = cx1 + pat1
        cx2 = cx2 + pat2
        cx3 = cx3 + pat3
    Next
    GetFadedColor = RGB(cx1, cx2, cx3)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

