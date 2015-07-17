Attribute VB_Name = "modTest"

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_END = 2
Public Const FILE_BEGIN = 0

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long

Public Const bufferSize = 32767 'Nice Big TCP Send Buffer, crappy on crap connections, great on LANS

Function ChangeByte(bytes As Single, Optional short As Boolean = False, Optional decPlaces As Long = 2) As String
    Dim count As Integer
    count = 0
    Dim SizeArray()
    If short Then
        SizeArray = Array(" b", " KB", " MB", " GB")
    Else
        SizeArray = Array(" Bytes", " Kilobytes", " Megabytes", " Gigabytes")
    End If
    Do While bytes >= 1000
        bytes = bytes / 1024
        count = count + 1
    Loop
    
ChangeByte = CStr(Round(bytes, decPlaces)) & SizeArray(count)
    
End Function

Public Function Seconds(difference As Single) As Long
    Seconds = CLng(difference * 24 * 60 * 60)
End Function
