Attribute VB_Name = "MMessage"
'*******************************************************************************
' MODULE:       MMessage
' FILENAME:     C:\My Code\vb\Msg\MMessage.bas
' AUTHOR:       Phil Fresle
' CREATED:      04-May-2001
' COPYRIGHT:    Copyright 2001 Frez Systems Limited. All Rights Reserved.
'
' DESCRIPTION:
' Optional standard module for an easy way to call my message box. If the
' function was renamed Msg instead of Msg it would be called instead of the
' standard VB/Windows message box everywhere in code unless prefixed VBA.
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

'*******************************************************************************
' Msg (FUNCTION)
'
' PARAMETERS:
' (In/Out) - Prompt   - String        - The text to display
' (In/Out) - Buttons  - VbMsgStyle - The buttons, style, icons
' (In/Out) - Title    - String        - The titlebar text
' (In/Out) - Helpfile - String        - The helpfile
' (In/Out) - Context  - Long          - The help context id
'
' RETURN VALUE:
' VbMsgResult - The button pressed to close the message box
'
' DESCRIPTION:
' Public entry point for the message box
'*******************************************************************************
Public Function msg(Prompt As String, _
                    Optional Buttons As VbMsgBoxStyle, _
                    Optional Title As String, _
                    Optional Helpfile As String, _
                    Optional Context As Long) As VbMsgBoxResult
                    
    Dim clsMessage As CMessage
    
    Set clsMessage = New CMessage
    frmMain.gSysTray.Show (True)
    msg = clsMessage.msg(Prompt, Buttons, Title, Helpfile, Context)
    
    Set clsMessage = Nothing
End Function
