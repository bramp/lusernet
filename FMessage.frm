VERSION 5.00
Begin VB.Form FMessage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton 
      Caption         =   "Command1"
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Command1"
      Height          =   435
      Index           =   1
      Left            =   1380
      TabIndex        =   4
      Top             =   720
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Command1"
      Height          =   435
      Index           =   2
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Command1"
      Height          =   435
      Index           =   3
      Left            =   3900
      TabIndex        =   2
      Top             =   720
      Width           =   1155
   End
   Begin LUSerNet.MySkinner Skinner 
      Left            =   0
      Top             =   0
      _extentx        =   1693
      _extenty        =   635
      imgne           =   "FMessage.frx":0000
      imgn            =   "FMessage.frx":0166
      imgnw           =   "FMessage.frx":02CE
      imge            =   "FMessage.frx":0884
      imgw            =   "FMessage.frx":0908
      imgse           =   "FMessage.frx":098C
      imgs            =   "FMessage.frx":0A0E
      imgsw           =   "FMessage.frx":0A92
      backcolor       =   16777215
      forecolor1      =   13545141
      forecolor2      =   9402231
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   180
      ScaleHeight     =   495
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblCaption 
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message Goes Here"
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1755
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       FMessage
' FILENAME:     C:\My Code\vb\MsgBox\FMessage.frm
' AUTHOR:       Phil Fresle
' CREATED:      04-May-2001
' COPYRIGHT:    Copyright 2001 Frez Systems Limited. All Rights Reserved.
'
' DESCRIPTION:
' The custom message box form. Called from CMessage. Needs msg.RES.
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
' to use the source code in your own code, but you may not claim that you created
' the sample code. It is expressly forbidden to sell or profit from this source code
' other than by the knowledge gained or the enhanced value added by your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
' Should you wish to commission some derivative work based on the add-in provided
' here, or any consultancy work, please do not hesitate to contact us.
'
' Web Site:  http://www.frez.co.uk
' E-mail:    sales@frez.co.uk
'
' MODIFICATION HISTORY:
' 1.0       04-May-2001
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

Public Event ButtonClicked(ByVal lButton As Long)

Private m_sButtonText()         As String
Private m_bButtonsUnload()      As Boolean
Private m_lSelectedButton       As Long
Private m_lIcon                 As Long

Private Const MSG_GAP           As Long = 8
Private Const MSG_GAP_ICON      As Long = 12
Private Const LEFT_MARGIN       As Long = 60
Private Const BUTTON_GAP        As Long = 105
Private Const LABEL_GAP         As Long = 240
Private Const LABEL_NORMAL_GAP  As Long = 135
Private Const MIN_WIDTH         As Long = 1950 ' 5145
Private Const MAX_WIDTH         As Long = 8715
Private Const RES_CRITICAL      As Long = 101
Private Const RES_QUESTION      As Long = 102
Private Const RES_EXCLAMATION   As Long = 103
Private Const RES_INFORMATION   As Long = 104

Private Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long

Private Const SM_CYCAPTION      As Long = 4
Private Const SM_CYBORDER       As Long = 6
Private Const SM_CYDLGFRAME     As Long = 8
Private Const SM_CXBORDER       As Long = 5
Private Const SM_CXDLGFRAME     As Long = 7

Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal hWndInsertAfter As Long, _
     ByVal X As Long, _
     ByVal Y As Long, _
     ByVal cx As Long, _
     ByVal cy As Long, _
     ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE        As Long = 2
Private Const SWP_NOSIZE        As Long = 1
Private Const flags             As Long = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST      As Long = -1
Private Const HWND_NOTOPMOST    As Long = -2

'*******************************************************************************
' MessageIcon (PROPERTY LET)
'*******************************************************************************
Public Property Let MessageIcon(ByVal NewValue As Long)
    Dim lWidth As Long
    
    m_lIcon = NewValue
    
    ' Load the appropriate icon and size the message label
    If m_lIcon = vbCritical Then
        picIcon.Picture = LoadResPicture(RES_CRITICAL, vbResBitmap)
        picIcon.Visible = True
        
    ElseIf NewValue = vbQuestion Then
        picIcon.Picture = LoadResPicture(RES_QUESTION, vbResBitmap)
        picIcon.Visible = True
        
    ElseIf NewValue = vbExclamation Then
        picIcon.Picture = LoadResPicture(RES_EXCLAMATION, vbResBitmap)
        picIcon.Visible = True
        
    ElseIf NewValue = vbInformation Then
        picIcon.Picture = LoadResPicture(RES_INFORMATION, vbResBitmap)
        picIcon.Visible = True
        
    Else
        picIcon.Visible = False
    End If
End Property

'*******************************************************************************
' MessageTitle (PROPERTY LET)
'*******************************************************************************
Public Property Let MessageTitle(ByVal NewValue As String)
    If NewValue = "" Then
        lblCaption.Caption = App.Title
    Else
        lblCaption.Caption = NewValue
    End If
End Property

'*******************************************************************************
' ButtonText (PROPERTY LET)
'*******************************************************************************
Public Property Let ButtonText(NewValue() As String)
    Dim lCount  As Long
    Dim lBorder As Long
    
    m_sButtonText = NewValue
    
    ' Make the appropriate buttons visible
    For lCount = cmdButton.LBound To cmdButton.UBound
        cmdButton(lCount).Visible = False
    Next
    For lCount = 0 To UBound(m_sButtonText)
        cmdButton(lCount).Visible = True
    Next
End Property

'*******************************************************************************
' ButtonsUnload (PROPERTY LET)
'*******************************************************************************
Public Property Let ButtonsUnload(NewValue() As Boolean)
    m_bButtonsUnload = NewValue
End Property

'*******************************************************************************
' TopMost (PROPERTY LET)
'*******************************************************************************
Public Property Let TopMost(NewValue As Boolean)
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Property

'*******************************************************************************
' ButtonDefault (PROPERTY LET)
'*******************************************************************************
Public Property Let ButtonDefault(NewValue As Long)
    cmdButton(NewValue).Default = True
End Property

'*******************************************************************************
' SelectedButton (PROPERTY GET)
'*******************************************************************************
Public Property Get SelectedButton() As Long
    SelectedButton = m_lSelectedButton
End Property

'*******************************************************************************
' ShowMessage (SUB)
'
' PARAMETERS:
' (In/Out) - sMessage - String - String to display
'
' DESCRIPTION:
' Display the string on the form sizing the form and positioning the controls.
'*******************************************************************************
Public Sub ShowMessage(sMessage As String)
    Dim lGap        As Long
    Dim lWidth      As Long
    Dim lBorder     As Long
    Dim lCount      As Long
    Dim lTextWidth  As Long
    Dim i As Long
    
    ' Get the forms borders width
    lBorder = Me.ScaleX(Me.Skinner.imgE.Width + Me.Skinner.imgW.Width, vbHimetric, vbTwips)
    
    ' We need to resize the form and message label based on the
    ' amount of text received taking into account the limits.
    If m_lIcon = vbCritical Or m_lIcon = vbQuestion Or m_lIcon = vbExclamation Or m_lIcon = vbInformation Then
        lWidth = Me.TextWidth(sMessage) + (2 * MSG_GAP_ICON) + (picIcon.Left + picIcon.Width)
    Else
        lWidth = Me.TextWidth(sMessage) + Me.ScaleX((2 * MSG_GAP), vbPixels, vbTwips) + lBorder
    End If
    
    If lWidth < MIN_WIDTH Then
        Me.Width = MIN_WIDTH
    ElseIf lWidth > MAX_WIDTH Then
        Me.Width = MAX_WIDTH
    Else
        Me.Width = lWidth
    End If
    
    'Position caption
    lblCaption.Top = (Me.ScaleY(Me.Skinner.imgN.Height, vbHimetric, vbPixels) - lblCaption.Height) / 2
    lblCaption.Left = Me.ScaleX(Me.Skinner.imgNW.Width, vbHimetric, vbPixels) + MSG_GAP / 2
    
    ' Size the label
    If m_lIcon = vbCritical Or m_lIcon = vbQuestion Or m_lIcon = vbExclamation Or m_lIcon = vbInformation Then
        lblMessage.Left = picIcon.Left + picIcon.Width + MSG_GAP_ICON
        lblMessage.Width = Me.Width - (2 * MSG_GAP_ICON) - (picIcon.Left + picIcon.Width)
    Else
        lblMessage.Left = MSG_GAP
        lblMessage.Width = Me.Width - (2 * MSG_GAP)
    End If
    
    lblMessage.Top = Me.ScaleY(Me.Skinner.imgN.Height, vbHimetric, vbPixels) + MSG_GAP
    lblMessage.Caption = sMessage
    
    For i = 0 To 3
        cmdButton(i).Top = lblMessage.Top + lblMessage.Height + MSG_GAP
    Next i
    ' Make sure the message box is tall enough to display all of the
    ' message. The message box label grows vertically automatically
    ' to fit the text
    If m_lIcon = vbCritical Or m_lIcon = vbQuestion Or m_lIcon = vbExclamation Or m_lIcon = vbInformation Then
        lGap = picIcon.Height + 2 * MSG_GAP_ICON + 4 * MSG_GAP + cmdButton(0).Height
    Else
        lGap = (2 * lblMessage.Top + lblMessage.Height) + 2 * MSG_GAP + cmdButton(0).Height
    End If
    
    Me.Height = ScaleY(cmdButton(0).Top + cmdButton(0).Height + MSG_GAP, vbPixels, vbTwips)
    
    ' Position the buttons, taking into account the width of the form
    ' and the number of buttons (note this property must be set after
    ' the caption and icon style)
    lCount = UBound(m_sButtonText) + 1
    Select Case lCount
        Case 1
            cmdButton(0).Left = (ScaleX(Me.Width, vbTwips, vbPixels) - cmdButton(0).Width - ScaleX(lBorder, vbTwips, vbPixels) + MSG_GAP) \ 2
            cmdButton(0).Caption = m_sButtonText(0)
        Case 2
            cmdButton(0).Left = ((ScaleX(Me.Width, vbTwips, vbPixels) - BUTTON_GAP - lBorder) \ 2) - cmdButton(0).Width
            cmdButton(1).Left = BUTTON_GAP + cmdButton(0).Width + cmdButton(0).Left
            cmdButton(0).Caption = m_sButtonText(0)
            cmdButton(1).Caption = m_sButtonText(1)
        Case 3
            cmdButton(1).Left = (ScaleX(Me.Width, vbTwips, vbPixels) - cmdButton(0).Width - lBorder) \ 2
            cmdButton(0).Left = cmdButton(1).Left - cmdButton(0).Width - BUTTON_GAP
            cmdButton(2).Left = cmdButton(1).Left + cmdButton(1).Width + BUTTON_GAP
            cmdButton(0).Caption = m_sButtonText(0)
            cmdButton(1).Caption = m_sButtonText(1)
            cmdButton(2).Caption = m_sButtonText(2)
        Case 4
            cmdButton(0).Left = ((ScaleX(Me.Width, vbTwips, vbPixels) - BUTTON_GAP - lBorder) \ 2) - cmdButton(0).Width - cmdButton(1).Width - BUTTON_GAP
            cmdButton(1).Left = BUTTON_GAP + cmdButton(0).Width + cmdButton(0).Left
            cmdButton(2).Left = BUTTON_GAP + cmdButton(1).Width + cmdButton(1).Left
            cmdButton(3).Left = BUTTON_GAP + cmdButton(2).Width + cmdButton(2).Left
            cmdButton(0).Caption = m_sButtonText(0)
            cmdButton(1).Caption = m_sButtonText(1)
            cmdButton(2).Caption = m_sButtonText(2)
            cmdButton(3).Caption = m_sButtonText(3)
    End Select
    
    Me.Show vbModal
End Sub

'*******************************************************************************
' cmdButton_Click (SUB)
'
' PARAMETERS:
' (In/Out) - Index - Integer - Number of button pressed
'
' DESCRIPTION:
' Button will either unload the form or raise an event depending on
' the m_bButtonsUnload array
'*******************************************************************************
Private Sub cmdButton_Click(index As Integer)
    ' Remember the button that was pressed
    m_lSelectedButton = index
    
    ' Unload or raise an event as appropriate
    If m_bButtonsUnload(m_lSelectedButton) Then
        Unload Me
    Else
        RaiseEvent ButtonClicked(index)
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Form_Resize()
    Skinner.Repaint Me
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub lblMessage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub picIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
