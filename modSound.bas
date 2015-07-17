Attribute VB_Name = "modSound"
Option Explicit

#If Win32 Then
    Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
#Else
    Private Declare Function sndPlaySound Lib "MMSYSTEM" (lpszSoundName As Any, ByVal uFlags%) As Integer
#End If

'Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_MEMORY = &H4
Const SND_LOOP = &H8
Const SND_NOSTOP = &H103
   
Dim bytSound() As Byte ' Always store binary data in byte arrays!


Public Sub PlayWaveRes(vntResourceID As Variant, Optional vntFlags)
    '-----------------------------------------------------------------
    ' WARNING:  If you want to play sound files asynchronously in
    '           Win32, then you MUST change bytSound() from a local
    '           variable to a module-level or static variable. Doing
    '           this prevents your array from being destroyed before
    '           sndPlaySound is complete. If you fail to do this, you
    '           will pass an invalid memory pointer, which will cause
    '           a GPF in the Multimedia Control Interface (MCI).
    '-----------------------------------------------------------------
    
    bytSound = LoadResData(vntResourceID, "WAVE")
    
    If IsMissing(vntFlags) Then
       vntFlags = SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY
    End If
    
    If (vntFlags And SND_MEMORY) = 0 Then
       vntFlags = vntFlags Or SND_MEMORY
    End If
    
    sndPlaySound bytSound(0), vntFlags
End Sub

