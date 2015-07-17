Attribute VB_Name = "modGUI"
Option Explicit

Public Sub AddResizers(frm As Form)
    Dim handleSizeX As Long
    Dim handleSizeY As Long
    
    handleSizeX = 6 '* Screen.TwipsPerPixelX
    handleSizeY = 6 '* Screen.TwipsPerPixelY
    
    '0=N, 1=E, 2=S, 3=W
    '4=NE, 5=SE, 6=SW, 7=NW
    
    'North
    frm.lblResizer(0).Move handleSizeX, 0, frm.ScaleWidth - 2 * handleSizeX, handleSizeY
    frm.lblResizer(0).Visible = True
    frm.lblResizer(0).MousePointer = 7
    
    'East
    frm.lblResizer(1).Move frm.ScaleWidth - handleSizeX, handleSizeY, handleSizeX, frm.ScaleHeight - 2 * handleSizeY
    frm.lblResizer(1).Visible = True
    frm.lblResizer(1).MousePointer = 9
    
    'South
    frm.lblResizer(2).Move handleSizeX, frm.ScaleHeight - handleSizeY, frm.ScaleWidth - 2 * handleSizeX, handleSizeY
    frm.lblResizer(2).Visible = True
    frm.lblResizer(2).MousePointer = 7
    
    'West
    frm.lblResizer(3).Move 0, handleSizeY, handleSizeX, frm.ScaleHeight - 2 * handleSizeY ' .Left = 0
    frm.lblResizer(3).Visible = True
    frm.lblResizer(3).MousePointer = 9
    
    'North East
    frm.lblResizer(4).Move frm.ScaleWidth - handleSizeX, 0, handleSizeX, handleSizeY
    frm.lblResizer(4).Visible = True
    frm.lblResizer(4).MousePointer = 6
    
    'South East
    frm.lblResizer(5).Move frm.ScaleWidth - handleSizeX, frm.ScaleHeight - handleSizeY, handleSizeX, handleSizeY
    frm.lblResizer(5).Visible = True
    frm.lblResizer(5).MousePointer = 8
    
    'South West
    frm.lblResizer(6).Move 0, frm.ScaleHeight - handleSizeY, handleSizeX, handleSizeY
    frm.lblResizer(6).Visible = True
    frm.lblResizer(6).MousePointer = 6
    
    'North West
    frm.lblResizer(7).Move 0, 0, handleSizeX, handleSizeY
    frm.lblResizer(7).Visible = True
    frm.lblResizer(7).MousePointer = 8
    
End Sub

Sub Resizer(frm As Form, index As Integer, x As Single, y As Single, minX As Single, minY As Single)
        
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
    
    Select Case index
        Case Is = 0: V1 = True
        Case Is = 1: H2 = True
        Case Is = 2: V2 = True
        Case Is = 3: H1 = True
        Case Is = 4: V1 = True: H2 = True
        Case Is = 5: V2 = True: H2 = True
        Case Is = 6: V2 = True: H1 = True
        Case Is = 7: V1 = True: H1 = True
    End Select
    
    If H2 Then
        newX = frm.Width + x
        If (newX < minX * Screen.TwipsPerPixelX) Then newX = minX * Screen.TwipsPerPixelX
        frm.Width = newX
    ElseIf H1 Then
        newX = frm.Width - x
        If (newX < minX * Screen.TwipsPerPixelX) Then
            newX = minX * Screen.TwipsPerPixelX
            x = frm.Width - newX
        End If
        frm.Width = newX
        frm.Left = frm.Left + x
    End If
    
    If V2 Then
        newY = frm.Height + y
        If (newY < minY * Screen.TwipsPerPixelY) Then newY = minY * Screen.TwipsPerPixelY
        frm.Height = newY
    ElseIf V1 Then
        newY = frm.Height - y
        If (newY < minY * Screen.TwipsPerPixelY) Then
            newY = minY * Screen.TwipsPerPixelY
            y = frm.Height - newY
        End If
        frm.Height = newY
        frm.Top = frm.Top + y
    End If
End Sub

Public Sub MakeRounded(frm As Form)
    Dim hRgn As Long
'    'hRgn = CreateRectRgn(0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight)
    hRgn = CreateRoundRectRgn(0, 0, frm.ScaleWidth + 1, frm.ScaleHeight + 1, 2, 2)
    SetWindowRgn frm.hwnd, hRgn, True
End Sub
