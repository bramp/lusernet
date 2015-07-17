VERSION 5.00
Begin VB.UserControl miniGraph 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H008F7777&
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HitBehavior     =   0  'None
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer tmrDraw 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   2880
   End
End
Attribute VB_Name = "miniGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Default Property Values:
Const m_def_BackColor = 0
Const m_def_Color1 = 255 'Red
Const m_def_Color2 = 16711680 'Blue
Const m_def_ScaleY = 800
Const m_def_ScaleX = 60

'Property Variables:
Dim m_BackColor As OLE_COLOR
Dim m_Color1 As OLE_COLOR
Dim m_Color2 As OLE_COLOR
Dim m_ScaleY As Long
Dim m_ScaleX As Long

Dim mLastUpload As Long
Dim mLastDownload As Long


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    UserControl.BackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AddData(Upload As Long, Download As Long) As Variant
    mLastUpload = Upload
    mLastDownload = Download
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Color1() As OLE_COLOR
Attribute Color1.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    Color1 = m_Color1
End Property

Public Property Let ScaleX(ByVal New_ScaleX As Long)
    m_ScaleX = New_ScaleX
    PropertyChanged "ScaleX"
End Property

Public Property Get ScaleX() As Long
    ScaleX = m_ScaleX
End Property

Public Property Let ScaleY(ByVal New_ScaleY As Long)
    m_ScaleY = New_ScaleY
    PropertyChanged "ScaleY"
End Property

Public Property Get ScaleY() As Long
    ScaleY = m_ScaleY
End Property

Public Property Let Color1(ByVal New_Color1 As OLE_COLOR)
    m_Color1 = New_Color1
    PropertyChanged "Color1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Color2() As OLE_COLOR
Attribute Color2.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    Color2 = m_Color2
End Property

Public Property Let Color2(ByVal New_Color2 As OLE_COLOR)
    m_Color2 = New_Color2
    PropertyChanged "Color2"
End Property

Private Sub tmrDraw_Timer()

    Dim x As Long
    Dim y As Long

    'Move Everything 1 pixel left
    For x = 1 To UserControl.ScaleWidth
        For y = 0 To UserControl.ScaleHeight
            UserControl.PSet (x - 1, y), UserControl.Point(x, y)
            'UserControl.PSet (x, y), 0
        Next y
    Next x
    
    UserControl.Line (UserControl.ScaleWidth - 1, UserControl.ScaleHeight)-(UserControl.ScaleWidth - 1, -1), UserControl.BackColor
    
    Dim UY As Long
    Dim DY As Long
    
    UY = UserControl.ScaleHeight - (UserControl.ScaleHeight / m_ScaleY) * mLastUpload
    DY = UserControl.ScaleHeight - (UserControl.ScaleHeight / m_ScaleY) * mLastDownload
    
    'Draw new lines
    If mLastUpload > mLastDownload Then
        UserControl.Line (UserControl.ScaleWidth - 1, UY)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), m_Color1
        UserControl.Line (UserControl.ScaleWidth - 1, DY)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), m_Color2
    Else
        UserControl.Line (UserControl.ScaleWidth - 1, DY)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), m_Color2
        UserControl.Line (UserControl.ScaleWidth - 1, UY)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), m_Color1
    End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_Color1 = m_def_Color1
    m_Color2 = m_def_Color2
    m_ScaleX = m_def_ScaleX
    m_ScaleY = m_def_ScaleY
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_Color1 = PropBag.ReadProperty("Color1", m_def_Color1)
    m_Color2 = PropBag.ReadProperty("Color2", m_def_Color2)
    m_ScaleX = PropBag.ReadProperty("m_ScaleX", m_def_ScaleX)
    m_ScaleY = PropBag.ReadProperty("m_ScaleY", m_def_ScaleY)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Color1", m_Color1, m_def_Color1)
    Call PropBag.WriteProperty("Color2", m_Color2, m_def_Color2)
    Call PropBag.WriteProperty("ScaleX", m_ScaleX, m_def_ScaleX)
    Call PropBag.WriteProperty("ScaleY", m_ScaleY, m_def_ScaleY)
End Sub

Public Sub Start()
    Let ScaleX = m_def_ScaleX
    Let ScaleY = m_def_ScaleY
    tmrDraw.Enabled = True
End Sub

Public Sub Finish()
    tmrDraw.Enabled = False
End Sub
