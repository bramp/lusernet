VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       CMessage
' FILENAME:     C:\My Code\vb\Msg\CMessage.cls
' AUTHOR:       Phil Fresle
' CREATED:      04-May-2001
' COPYRIGHT:    Copyright 2001 Frez Systems Limited. All Rights Reserved.
'
' DESCRIPTION:
' Wrapper class for message box and error message functionality. Requires
' FMessage for the message form. Allows you to have your own style of message
' box (i.e. whatever font, button size, colour scheme you desire). If you want
' to have this message box popup instead of VB's standard one, just create a
' public Msg function in a standard module like the Msg function in
' the MMessage module, you can then still call the VB version if required by
' prefixing it as follows: VBA.Msg. Requires msg.RES.
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
' 1.1       29-Aug-2001
'           Phil Fresle
'           Updated to fix binary logic when choosing which button was pressed.
'*******************************************************************************
Option Explicit

Private WithEvents m_frmMessage As FMessage
Attribute m_frmMessage.VB_VarHelpID = -1
Private WithEvents m_frmError   As FMessage
Attribute m_frmError.VB_VarHelpID = -1

Private m_sMessage As String

'*******************************************************************************
' Msg (FUNCTION)
'
' PARAMETERS:
' (In/Out) - Prompt   - String        - The message
' (In/Out) - Buttons  - VbMsgStyle - The buttons, icon, style
' (In/Out) - Title    - String        - The titlebar text
' (In/Out) - Helpfile - String        - The helpfile to use
' (In/Out) - Context  - Long          - The context id
'
' RETURN VALUE:
' VbMsgResult - The button that was pressed
'
' DESCRIPTION:
' Provides your own custom message box functionality with a similar interface
' to VB's own Msg function.
'*******************************************************************************
Public Function msg(Prompt As String, _
                       Optional Buttons As VbMsgBoxStyle, _
                       Optional Title As String, _
                       Optional Helpfile As String, _
                       Optional Context As Long) As VbMsgBoxResult
                       
    Dim sButtons()          As String
    Dim bButtonsUnload()    As Boolean
    Dim sHelpfile           As String
    Dim lButtonType         As Long
    Dim lIconType           As Long
    Dim lDefaultType        As Long
    
    Const MASK_BUTTONS  As Long = &H7   ' 0000000111
    Const MASK_ICONS    As Long = &H70  ' 0001110000
    Const MASK_DEFAULTS As Long = &H300 ' 1100000000
    
    Set m_frmMessage = New FMessage
    
    m_sMessage = Prompt
    
    With m_frmMessage
        .MessageTitle = Title
        .HelpContextID = Context
    End With
    
    ' Decide on the text to be displayed for each button and which
    ' buttons cause the messagebox to unload
    lButtonType = (Buttons And MASK_BUTTONS)
    Select Case lButtonType
        Case vbRetryCancel
            If (Buttons And vbMsgBoxHelpButton) Then
                ReDim sButtons(2)
                ReDim bButtonsUnload(2)
                sButtons(2) = "Help"
                bButtonsUnload(2) = False
            Else
                ReDim sButtons(1)
                ReDim bButtonsUnload(1)
            End If
            sButtons(0) = "Retry"
            sButtons(1) = "Cancel"
            bButtonsUnload(0) = True
            bButtonsUnload(1) = True
        
        Case vbYesNo
            If (Buttons And vbMsgBoxHelpButton) Then
                ReDim sButtons(2)
                ReDim bButtonsUnload(2)
                sButtons(2) = "Help"
                bButtonsUnload(2) = False
            Else
                ReDim sButtons(1)
                ReDim bButtonsUnload(1)
            End If
            sButtons(0) = "Yes"
            sButtons(1) = "No"
            bButtonsUnload(0) = True
            bButtonsUnload(1) = True
        
        Case vbYesNoCancel
            If (Buttons And vbMsgBoxHelpButton) Then
                ReDim sButtons(3)
                ReDim bButtonsUnload(3)
                sButtons(3) = "Help"
                bButtonsUnload(3) = False
            Else
                ReDim sButtons(2)
                ReDim bButtonsUnload(2)
            End If
            sButtons(0) = "Yes"
            sButtons(1) = "No"
            sButtons(2) = "Cancel"
            bButtonsUnload(0) = True
            bButtonsUnload(1) = True
            bButtonsUnload(2) = True
        
        Case vbAbortRetryIgnore
            If (Buttons And vbMsgBoxHelpButton) Then
                ReDim sButtons(3)
                ReDim bButtonsUnload(3)
                sButtons(3) = "Help"
                bButtonsUnload(3) = False
            Else
                ReDim sButtons(2)
                ReDim bButtonsUnload(2)
            End If
            sButtons(0) = "Abort"
            sButtons(1) = "Retry"
            sButtons(2) = "Ignore"
            bButtonsUnload(0) = True
            bButtonsUnload(1) = True
            bButtonsUnload(2) = True
        
        Case vbOKCancel
            If (Buttons And vbMsgBoxHelpButton) Then
                ReDim sButtons(2)
                ReDim bButtonsUnload(2)
                sButtons(2) = "Help"
                bButtonsUnload(2) = False
            Else
                ReDim sButtons(1)
                ReDim bButtonsUnload(1)
            End If
            sButtons(0) = "OK"
            sButtons(1) = "Cancel"
            bButtonsUnload(0) = True
            bButtonsUnload(1) = True
        
        Case Else    ' vbOKOnly
            If (Buttons And vbMsgBoxHelpButton) Then
                ReDim sButtons(1)
                ReDim bButtonsUnload(1)
                sButtons(1) = "Help"
                bButtonsUnload(1) = False
            Else
                ReDim sButtons(0)
                ReDim bButtonsUnload(0)
            End If
            sButtons(0) = "OK"
            bButtonsUnload(0) = True
    End Select
    
    With m_frmMessage
        .ButtonText = sButtons
        .ButtonsUnload = bButtonsUnload
                
        ' What type of icon is required
        lIconType = (Buttons And MASK_ICONS)
        Select Case lIconType
            Case vbCritical
                .MessageIcon = vbCritical
            Case vbExclamation
                .MessageIcon = vbExclamation
            Case vbInformation
               .MessageIcon = vbInformation
            Case vbQuestion
                .MessageIcon = vbQuestion
            Case Else
                .MessageIcon = 0
        End Select
        
        ' Which button is the default
        lDefaultType = (Buttons And MASK_DEFAULTS)
        Select Case lDefaultType
            Case vbDefaultButton2
                .ButtonDefault = 1
            Case vbDefaultButton3
                .ButtonDefault = 2
            Case vbDefaultButton4
                .ButtonDefault = 3
            Case Else ' Buttons And vbDefaultButton1
                .ButtonDefault = 0
        End Select
        
        ' Should the text be alligned right or left?
        If (Buttons And vbMsgBoxRight) Then
            .lblMessage.Alignment = vbRightJustify
        Else
            .lblMessage.Alignment = vbLeftJustify
        End If
        
        ' Should the messagebox appear topmost?
        If (Buttons And vbMsgBoxSetForeground) Then
            .TopMost = True
        Else
            .TopMost = False
        End If
        
        ' Set the helpfile (remember the current application helpfile)
        sHelpfile = App.Helpfile
        App.Helpfile = Helpfile
        
        ' Get the messagebox form to show itself with the message
        .ShowMessage Prompt
        
        ' Reset the application helpfile
        App.Helpfile = sHelpfile
        
        ' Which key was pressed to terminate the messagebox?
        Select Case .SelectedButton
            Case 0
                Select Case lButtonType
                    Case vbAbortRetryIgnore
                        msg = vbAbort
                    Case vbOKCancel
                        msg = vbOK
                    Case vbRetryCancel
                        msg = vbRetry
                    Case vbYesNo
                        msg = vbYes
                    Case vbYesNoCancel
                        msg = vbYes
                    Case Else ' vbOKOnly
                        msg = vbOK
                End Select
                
            Case 1
                Select Case lButtonType
                    Case vbAbortRetryIgnore
                        msg = vbRetry
                    Case vbOKCancel
                        msg = vbCancel
                    Case vbRetryCancel
                        msg = vbCancel
                    Case vbYesNo
                        msg = vbNo
                    Case vbYesNoCancel
                        msg = vbNo
                End Select
            
            Case 2
                Select Case lButtonType
                    Case vbAbortRetryIgnore
                        msg = vbIgnore
                    Case vbYesNoCancel
                        msg = vbCancel
                End Select
        End Select
    End With
    
    Set m_frmMessage = Nothing
End Function

'*******************************************************************************
' DisplayError (SUB)
'
' PARAMETERS:
' (In/Out) - Prompt     - String  - The error text to display
' (In)     - HelpButton - Boolean - Whether a help button is required
' (In)     - Context    - Long    - The help context id
'
' DESCRIPTION:
' Display an error message in my own version of the message box, but provide
' a Copy button that the user can use to copy the text to the clipboard so
' they can, for instance, paste the text in an email.
'*******************************************************************************
Public Sub DisplayError(Prompt As String, _
                        Optional ByVal HelpButton As Boolean, _
                        Optional ByVal Context As Long)
    Dim sButtons()          As String
    Dim bButtonsUnload()    As Boolean
    'Dim lButtonType         As Long
    
    Set m_frmError = New FMessage
    
    m_sMessage = Prompt
    
    With m_frmError
        If App.ProductName <> "" Then
            .MessageTitle = App.ProductName
        Else
            .MessageTitle = App.Title
        End If
        .HelpContextID = Context
    End With
    
    If HelpButton Then
        ReDim sButtons(2)
        ReDim bButtonsUnload(2)
        sButtons(2) = "Help"
        bButtonsUnload(2) = False
    Else
        ReDim sButtons(1)
        ReDim bButtonsUnload(1)
    End If
    
    sButtons(0) = "Close"
    bButtonsUnload(0) = True
    sButtons(1) = "Copy"
    bButtonsUnload(1) = False
    
    With m_frmError
        .ButtonText = sButtons
        .ButtonsUnload = bButtonsUnload
        .MessageIcon = vbCritical
        .ButtonDefault = 0
        .lblMessage.Alignment = vbLeftJustify
        .TopMost = True
        
        ' Get the messagebox form to show itself with the message
        .ShowMessage Prompt
    End With
    
    Set m_frmError = Nothing
End Sub

'*******************************************************************************
' m_frmError_ButtonClicked (SUB)
'
' PARAMETERS:
' (In) - lButton - Long - Button that was pressed on the form
'
' DESCRIPTION:
' An event sink for button presses on the message form. Only those buttons that
' do not cause the form to close raise an event out.
'*******************************************************************************
Private Sub m_frmError_ButtonClicked(ByVal lButton As Long)
    Select Case lButton
        Case 1
            Clipboard.Clear
            Clipboard.SetText m_sMessage
            
        Case 2
            SendKeys "{F1}", False
    End Select
End Sub

'*******************************************************************************
' m_frmMessage_ButtonClicked (SUB)
'
' PARAMETERS:
' (In) - lButton - Long - Button that was pressed on the form
'
' DESCRIPTION:
' An event sink for button presses on the message form. Only those buttons that
' do not cause the form to close raise an event out. The only button click
' message we will get in messagebox mode is from the help button.
'*******************************************************************************
Private Sub m_frmMessage_ButtonClicked(ByVal lButton As Long)
    SendKeys "{F1}", False
End Sub
