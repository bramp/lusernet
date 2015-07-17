VERSION 5.00
Begin VB.UserControl MyButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer tmrOff 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   1200
   End
   Begin VB.Image picMain 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "MyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************************************************************
'** This is the MyButton object. My picture buttons        **
'** By Andrew Brampton 16/03/2001                          **
'** Tested 16/03/2001                                      **
'** Known Bugs : NONE                                      **
'************************************************************
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function apiGetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long

Private Declare Function apiWindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'local variable(s) to hold property value(s)
Private mvarpicNormal As StdPicture  'local copy
Private mvarpicX As Long 'local copy
Private mvarpicY As Long 'local copy

Private pntCursor As POINTAPI
Private mbHasCapture As Boolean

Public Event Click()

Public Property Let picX(ByVal vData As String)
    mvarpicX = vData
    Call MovePic(mvarpicX, mvarpicY)
End Property

Public Property Get picX() As String
    picX = mvarpicX
End Property

Public Property Let picY(ByVal vData As String)
    mvarpicY = vData
    Call MovePic(mvarpicX, mvarpicY)
End Property

Public Property Get picY() As String
    picY = mvarpicY
End Property

Public Property Set picNormal(ByVal vData As StdPicture)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.picNormal = 5
    Set mvarpicNormal = vData
    Call ChangePic(mvarpicNormal, mvarpicX, mvarpicY)
End Property

Public Property Get picNormal() As StdPicture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.picNormal
    Set picNormal = mvarpicNormal
End Property

Private Sub picMain_Click()
RaiseEvent Click
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Ambient.UserMode = True Then tmrOff.Enabled = True
End Sub

Private Sub tmrOff_Timer()
        
    apiGetCursorPos pntCursor

    If apiWindowFromPoint(pntCursor.x, pntCursor.y) = UserControl.hwnd Then

        'Mouse Over
        If Not mbHasCapture Then
            Call MovePic(mvarpicX, picMain.Height / 2)
            mbHasCapture = True
        End If
    Else
        
        'Mouse Off
        If mbHasCapture Then
            Call MovePic(mvarpicX, mvarpicY)
            mbHasCapture = False
        End If
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set mvarpicNormal = PropBag.ReadProperty("picNormal", "")
    mvarpicX = PropBag.ReadProperty("picX", 0)
    mvarpicY = PropBag.ReadProperty("picY", 0)
    Call ChangePic(mvarpicNormal, mvarpicX, mvarpicY)
    Call tmrOff_Timer
    picMain.ToolTipText = PropBag.ReadProperty("ToolTip", "")
End Sub

Private Sub ChangePic(pic As StdPicture, Optional x As Long, Optional y As Long)
    On Error Resume Next
    picMain.Picture = pic
    picMain.Left = -x
    picMain.Top = -y
End Sub

Private Sub MovePic(x As Long, y As Long)
    If picMain.Left <> -x Then picMain.Left = -x
    If picMain.Top <> -y Then picMain.Top = -y
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "picNormal", mvarpicNormal, 0
    PropBag.WriteProperty "picX", mvarpicX, "0"
    PropBag.WriteProperty "picY", mvarpicY, "0"
    Call PropBag.WriteProperty("ToolTip", picMain.ToolTipText, "")
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picMain,picMain,-1,ToolTip
Public Property Get ToolTip() As String
Attribute ToolTip.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTip = picMain.ToolTipText
End Property

Public Property Let ToolTip(ByVal New_ToolTip As String)
    picMain.ToolTipText = New_ToolTip
    PropertyChanged "ToolTip"
End Property

