VERSION 5.00
Begin VB.UserControl MyIPRange 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picIPRange 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   4
      Top             =   0
      Width           =   1600
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   390
         TabIndex        =   1
         Text            =   "0"
         Top             =   15
         Width           =   375
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   780
         TabIndex        =   2
         Text            =   "0"
         Top             =   15
         Width           =   375
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   1170
         TabIndex        =   3
         Text            =   "1"
         Top             =   15
         Width           =   375
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Text            =   "127"
         Top             =   15
         Width           =   375
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   30
         Width           =   135
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   750
         TabIndex        =   6
         Top             =   30
         Width           =   135
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1140
         TabIndex        =   5
         Top             =   30
         Width           =   135
      End
   End
End
Attribute VB_Name = "MyIPRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const m_def_IPAddress = "127.0.0.1"
'Default Property Values:
Const m_def_Enabled = 0
'Property Variables:
Dim m_Enabled As Boolean



Private Sub txtIP_Change(Index As Integer)
    If Val(txtIP(Index).Text) > 255 Then txtIP(Index).Text = "255"
    If txtIP(Index).Text = "" Then txtIP(Index).Text = "0"
End Sub

Private Sub txtIP_KeyPress(Index As Integer, keyascii As Integer)
    Select Case keyascii
        Case Asc("0") To Asc("9"):
            If NewLength(txtIP(Index)) >= 3 Then
                If Index <> 3 Then txtIP(Index + 1).SetFocus
                If NewLength(txtIP(Index)) > 3 Then keyascii = 0
            End If
        Case Is = 8: 'Backspace
            If txtIP(Index).SelStart = 0 Then
                If Index <> 0 Then
                    txtIP(Index - 1).SetFocus
                    txtIP(Index - 1).SelStart = Len(txtIP(Index - 1).Text)
                Else
                    keyascii = 0
                End If
            End If
        Case Is = 9: 'Tab
            If Index <> 3 Then
                txtIP(Index + 1).SetFocus
                txtIP(Index + 1).SelStart = 0
            Else
                keyascii = 0
            End If
        Case Else: keyascii = 0
    End Select
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = picIPRange.Width
    UserControl.Height = picIPRange.Height
End Sub

'Works out the new length of the text box, presuming the key actually adds data, ie not a del or a backspace
Private Function NewLength(txt As TextBox) As Long
    NewLength = Len(txt.Text) + 1 - txt.SelLength
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get IPAddress() As String
    IPAddress = txtIP(0).Text & "." & txtIP(1).Text & "." & txtIP(2).Text & "." & txtIP(3).Text
End Property

Public Property Let IPAddress(ByVal new_ipAddress As String)
    SetIP new_ipAddress
    PropertyChanged "IPAddress"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    SetIP PropBag.ReadProperty("IPAddress", m_def_IPAddress)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("IPAddress", txtIP(0).Text & "." & txtIP(1).Text & "." & txtIP(2).Text & "." & txtIP(3).Text, m_def_IPAddress)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End Sub

Private Sub SetIP(new_ipAddress As String)
    Dim tmp() As String
    Dim i As Long
    
    tmp = Split(new_ipAddress, ".")
    
    For i = 0 To 3
        txtIP(i).Text = tmp(i)
    Next i
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    Dim i As Long
    For i = 0 To 3
        txtIP(i).Enabled = m_Enabled
    Next i
    PropertyChanged "Enabled"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
End Sub

