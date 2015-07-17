VERSION 5.00
Begin VB.UserControl UpDown 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdDown 
      Appearance      =   0  'Flat
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   6.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1560
   End
   Begin VB.CommandButton cmdUp 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   6.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   300
      Width           =   1560
   End
End
Attribute VB_Name = "UpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event Change()
'Default Property Values:
Const m_def_Value = 0
Const m_def_Min = 0
Const m_def_Max = 0
'Property Variables:
Dim m_Value As Long
Dim m_Min As Long
Dim m_Max As Long

Private Sub cmdDown_Click()
    If m_Value > m_Min Then m_Value = m_Value - 1
    RaiseEvent Change
End Sub

Private Sub cmdUp_Click()
    If m_Value < m_Max Then m_Value = m_Value + 1
    RaiseEvent Change
End Sub

Private Sub UserControl_Resize()
cmdUp.Top = 0
cmdUp.Height = UserControl.ScaleHeight / 2
cmdUp.Left = 0
cmdUp.Width = UserControl.ScaleWidth
cmdDown.Top = cmdUp.Height
cmdDown.Height = UserControl.ScaleHeight - cmdDown.Top
cmdDown.Left = 0
cmdDown.Width = UserControl.ScaleWidth
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    'If m_Value <> New_Value Then
    '    m_Value = New_Value
    '    RaiseEvent Change
    'Else
        m_Value = New_Value
    'End If
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Min() As Long
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_Min = m_def_Min
    m_Max = m_def_Max
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
End Sub

