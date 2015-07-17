VERSION 5.00
Begin VB.UserControl HyperLink 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label lblLink 
      BackStyle       =   0  'Transparent
      Caption         =   "www.Whatever.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      MouseIcon       =   "HyperLink.ctx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "HyperLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Default Property Values:
Const m_def_URL = ""
'Property Variables:
Dim m_URL As String
'Event Declarations:
Event Click() 'MappingInfo=lblLink,lblLink,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLink,lblLink,-1,Caption
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Text = lblLink.Caption
End Property

Public Property Let Text(ByVal New_Text As String)
    lblLink.Caption = New_Text
    PropertyChanged "Text"
End Property

Private Sub lblLink_Click()
    ShellExecute UserControl.hwnd, "open", m_URL, "", "", 0
    RaiseEvent Click
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_URL = m_def_URL
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblLink.Caption = PropBag.ReadProperty("Text", "www.Whatever.com")
    m_URL = PropBag.ReadProperty("URL", m_def_URL)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

Private Sub UserControl_Resize()
    lblLink.Left = 0
    lblLink.Top = 0
    lblLink.Width = UserControl.Width
    lblLink.Height = UserControl.Height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", lblLink.Caption, "www.Whatever.com")
    Call PropBag.WriteProperty("URL", m_URL, m_def_URL)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get URL() As String
    URL = m_URL
End Property

Public Property Let URL(ByVal New_URL As String)
    m_URL = New_URL
    PropertyChanged "URL"
End Property

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

