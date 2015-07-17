Attribute VB_Name = "modAll"
'HELLO|ip
'HI|Version|Files|Folders|Size|
'GET|Filename|size|position|
'OK|Filename|size|position|port|
'ERROR|Filename|Error
'FIND|SearchText|
'FOUND|Match Name|Size|Match Name|Size ... |

Option Explicit

Global conn As New ADODB.Connection
Global rs As ADODB.Recordset

'Const SubNetMask = "10.38.255.255"
Public Const SubNetMask = "255.255.255.255"
'Public Const BlockedSubNetMask = "10.38.240.255"
'Public Const BlockedSubNetMask = "192.168.0.255"

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 0
Private Const MAX_PATH = 260

Private Const SearchLimit = 100

Public Const DownloadEXT = ".LUSerNet"

Public Const bufferSize = 32767 'Nice Big TCP Send Buffer, crappy on crap connections, great on LANS

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long

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

Public Const GW_HWNDPREV = 3
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

'**************************************
'Windows API/Global Declarations for :__
'     ____A Edit Registry
'**************************************

Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const KEY_ALL_ACCESS = &H3F
Public Const REG_OPTION_NON_VOLATILE = 0&
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = (1)

Private Const MAX_ADAPTER_NAME_LENGTH         As Long = 256
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH  As Long = 128
Private Const MAX_ADAPTER_ADDRESS_LENGTH      As Long = 8

Private Type IP_ADDRESS_STRING
    IpAddr(0 To 15)  As Byte
End Type

Private Type IP_MASK_STRING
    IpMask(0 To 15)  As Byte
End Type

Private Type IP_ADDR_STRING
    dwNext     As Long
    IpAddress  As IP_ADDRESS_STRING
    IpMask     As IP_MASK_STRING
    dwContext  As Long
End Type

Private Type IP_ADAPTER_INFO
    dwNext                As Long
    ComboIndex            As Long  'reserved
    sAdapterName(0 To (MAX_ADAPTER_NAME_LENGTH + 3))        As Byte
    sDescription(0 To (MAX_ADAPTER_DESCRIPTION_LENGTH + 3)) As Byte
    dwAddressLength       As Long
    sIPAddress(0 To (MAX_ADAPTER_ADDRESS_LENGTH - 1))       As Byte
    dwIndex               As Long
    uType                 As Long
    uDhcpEnabled          As Long
    CurrentIpAddress      As Long
    IpAddressList         As IP_ADDR_STRING
    GatewayList           As IP_ADDR_STRING
    DhcpServer            As IP_ADDR_STRING
    bHaveWins             As Long
    PrimaryWinsServer     As IP_ADDR_STRING
    SecondaryWinsServer   As IP_ADDR_STRING
    LeaseObtained         As Long
    LeaseExpires          As Long
End Type

Private Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (pTcpTable As Any, pdwSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal bcount As Long)

Function ShowBrowseFolders(message As String) As String
    'Opens a Browse Folders Dialog Box that displays the directories in your computer
    Dim lpIDList As Long ' Declare Varibles
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = message

    With tBrowseInfo
        .hWndOwner = frmMain.hwnd ' Owner Form
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        ShowBrowseFolders = sBuffer
    End If

    'Apparent Member Leak.... Not sure how to fix
End Function

Function FileExist(filename As String) As Long
    'Returns status of a file, ie Not Found, Access Denied etc
    Dim rValue As Long
    rValue = 0
    On Error GoTo FileExistError
    Open filename For Random Shared As #1
    Close #1

    FileExist = rValue

    Exit Function
FileExistError:
    Select Case Err.number
        Case Is = 53: rValue = 1 'File Not Found
        Case Is = 76: rValue = 2 'Folder Not Found
        Case Is = 75: rValue = 3 'Access Denied
        Case Is = 52: rValue = 4 'Computer Not Found
        Case Else: msg Err.number & " " & Err.Description
    End Select

    FileExist = rValue

End Function

Public Sub CompactTheDatabase(strDBPath As String)
    Dim strTempLoc          As String
    Dim jeEngine            As New JRO.JetEngine

    On Error GoTo CompactTheDatabaseError

    '0 = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    ' get temp db name and path - use windows temp folder
    strTempLoc = App.path & "\" & "temp.mdb"

    ' make sure no file is in the way of the temp DB
    If Dir(strTempLoc) <> "" Then Kill strTempLoc

    ' compact the database - do not encrypt
    jeEngine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTempLoc & ";Jet OLEDB:Encrypt Database=False"
    ' kill the original
    Kill strDBPath

    ' copy the new to the old
    FileCopy strTempLoc, strDBPath

    ' check for the new (old) name, then kill the temp
    If Dir(strDBPath) <> "" Then
        Kill strTempLoc
    Else
        GoTo CompactTheDatabaseError
    End If

    ' indicate success
    'Msg "The database has been successfully compacted.", vbOKOnly Or vbInformation, "Operation successful"

    ' clean up
    Set jeEngine = Nothing
    Screen.MousePointer = 0

    Exit Sub

CompactTheDatabaseError:
    Set jeEngine = Nothing
    Screen.MousePointer = 0
    If Dir(strDBPath) = "" Then
        If Len(strTempLoc) > 0 Then FileCopy strTempLoc, strDBPath
    End If
    'Msg "An error occurred while attempting to compact the database.", vbOKOnly Or vbExclamation, "Operation Error"
End Sub

Function ChangeByte(bytes As Single, Optional short As Boolean = False, Optional decPlaces As Long = 2) As String
    Dim count As Integer
    count = 0
    Dim SizeArray()
    If short Then
        SizeArray = Array(" b", " KB", " MB", " GB", " TB")
    Else
        SizeArray = Array(" Bytes", " Kilobytes", " Megabytes", " Gigabytes", " Terrabytes")
    End If
    Do While bytes >= 1000
        bytes = bytes / 1024
        count = count + 1
    Loop

    ChangeByte = CStr(Round(bytes, decPlaces)) & SizeArray(count)

End Function

Function ChangeSecond(seconds As Long) As String

    Dim output As String

    If seconds >= 86400 Then
        output = output & seconds \ 86400 & "d"
        seconds = seconds Mod 86400
    End If
    If seconds >= 3600 Then
        output = output & seconds \ 3600 & "h"
        seconds = seconds Mod 3600
    End If
    If seconds >= 60 Then
        output = output & seconds \ 60 & "m"
        seconds = seconds Mod 60
    End If
    If seconds >= 0 Then
        output = output & seconds & "s"
    End If
    ChangeSecond = output
End Function

Function SearchFile(filename As String) As String
    'This function searchs the database and returns a list of matched files in the format:
    'filename|filesize|filename|filesize|

    Dim searchTxt As String
    Dim rValue As String
    Dim count As Long

    searchTxt = Replace(filename, "'", "''")
    searchTxt = Replace(searchTxt, " ", "%")
    searchTxt = Replace(searchTxt, "*", "%")
    searchTxt = Replace(searchTxt, "[", "")
    searchTxt = Replace(searchTxt, "]", "")

    searchTxt = "%" & searchTxt & "%"

    Set rs = conn.Execute("SELECT TOP " & SearchLimit & " fileName, fileSize  FROM tblFiles WHERE fileName LIKE '" & searchTxt & "'")

    count = 0
    Do While Not rs.EOF
        rValue = rValue & rs("filename") & "|" & rs("fileSize") & "|"
        rs.MoveNext
        count = count + 1
    Loop

    If count <> 0 Then rs.Close

    Set rs = Nothing
    
    If count < SearchLimit Then
        rValue = rValue & SuperSearchFile(searchTxt, SearchLimit - count)
    End If
    
    SearchFile = rValue

End Function

Function SuperSearchFile(filename As String, limit As Long) As String
    'This function searchs the database and returns a list of matched files in directoys of the format:
    'filename|filesize|filename|filesize|
    Dim rValue As String
    Dim count As Long
    Dim innerCount As Long
    Dim folderName As String
    Dim rs2 As ADODB.Recordset
    
    'finds all folders with that search term in
    Set rs = conn.Execute("SELECT folderID, folderName FROM tblFolders WHERE folderName LIKE '" & filename & "'")

    rValue = ""
    count = 0
    
    'Loops through all folders
    Do While Not rs.EOF And count < limit
        
        'Finds all files in the folder, which doesn't have the search term in, otherwise they would of been found allready
        Set rs2 = conn.Execute("SELECT TOP " & limit - count & " fileName, fileSize FROM tblFiles WHERE folderID = " & rs("folderID") & " AND NOT (fileName LIKE '" & filename & "')")
        folderName = Right(rs("folderName"), Len(rs("folderName")) - InStrRev(rs("folderName"), "\"))
        innerCount = 0
        Do While Not rs2.EOF
            'have to clean foldername incase it contains full path
            rValue = rValue & folderName & "\" & rs2("filename") & "|" & rs2("fileSize") & "|"
            rs2.MoveNext
            count = count + 1
            innerCount = innerCount + 1
        Loop
        
        rs.MoveNext
        
        If innerCount <> 0 Then rs2.Close
        Set rs2 = Nothing
    Loop

    If count <> 0 Then rs.Close

    Set rs = Nothing
    
    SuperSearchFile = rValue

End Function

Function GetIP() As String
    GetIP = frmMain.UDPListen(0).LocalIP
End Function

Function isNothing(object As Object) As Boolean

    On Error GoTo isNothingError
    If object Then isNothing = False

    Exit Function
isNothingError:
    If Err.number = 91 Or Err.number = 340 Then isNothing = True

End Function

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Public Function seconds(difference As Single) As Long
    seconds = CLng(difference * 24 * 60 * 60)
End Function

Public Sub Hash(ByRef Data As String)
    Dim outData As String
    Dim i As Long
    outData = ""
    For i = 1 To Len(Data)
        outData = outData & Chr(Asc(Mid(Data, i, 1)) Xor 128)
    Next i
    Data = outData
End Sub

Public Sub Clean(ByRef Data As String)
    Dim outData As String
    Dim char As String * 1
    Dim i As Long
    For i = 1 To Len(Data)
        char = Mid(Data, i, 1)
        Select Case Asc(char)
            Case 32 To 126, 160 To 55: outData = outData & char
            Case Else
        End Select
    Next i
    Data = outData
End Sub

Public Sub ActivatePrevInstance()
    Dim OldTitle As String
    Dim PrevHndl As Long
    Dim result As Long
    'Save the title of the application.
    OldTitle = App.Title
    'Rename the title of this application so FindWindow
    'will not find this application instance

    App.Title = "unwanted instance"

    PrevHndl = FindWindow("ThunderRT6Main", OldTitle)

    If PrevHndl = 0 Then
        'No previous instance found.
        Exit Sub
    End If

    'Get handle to previous window.
    PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
    'Restore the program.
    result = OpenIcon(PrevHndl)
    'Activate the application.
    result = SetForegroundWindow(PrevHndl)
    'End the application.
    End
End Sub


Public Function FileType(filename As String) As String
    Dim lastDot As Long
    lastDot = InStrRev(filename, ".")
    If lastDot <> 0 Then
        FileType = LCase(Trim(Right(filename, Len(filename) - lastDot)))
    Else
        FileType = ""
    End If
End Function

Public Function SetRegValue(ByVal hKey As Long, ByVal lpszSubKey As String, ByVal sSetValue As String, ByVal sValue As String) As Boolean

    'On Error Resume Next
    Dim phkResult As Long
    Dim lResult As Long
    Dim SA As SECURITY_ATTRIBUTES
    Dim lCreate As Long

    RegCreateKeyEx hKey, lpszSubKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, phkResult, lCreate
    lResult = RegSetValueEx(phkResult, sSetValue, 0, REG_SZ, sValue, CLng(Len(sValue) + 1))
    RegCloseKey phkResult
    SetRegValue = (lResult = ERROR_SUCCESS)

End Function


Public Function LocalIPAddresses() As String()

    'api vars
    Dim cbRequired  As Long
    Dim buff()      As Byte
    Dim Adapter     As IP_ADAPTER_INFO

    'working vars
    Dim ptr1        As Long
    Dim sIPAddr     As String
    Dim sAllAddr()    As String
    ReDim Preserve sAllAddr(0)
    Call GetAdaptersInfo(ByVal 0&, cbRequired)

    If cbRequired > 0 Then

        ReDim buff(0 To cbRequired - 1) As Byte

        If GetAdaptersInfo(buff(0), cbRequired) = ERROR_SUCCESS Then

            'get a pointer to the data stored in buff()
            ptr1 = VarPtr(buff(0))

            'ptr1 is 0 when no more adapters
            Do While (ptr1 <> 0)

                'copy the data from the pointer to the
                'first adapter into the IP_ADAPTER_INFO type
                CopyMemory Adapter, ByVal ptr1, LenB(Adapter)

                'the DHCP IP address is in the
                'IpAddress.IpAddr member
                sIPAddr = TrimNull(StrConv(Adapter.IpAddressList.IpAddress.IpAddr, vbUnicode))
                sAllAddr(UBound(sAllAddr)) = Trim(sIPAddr)
                ReDim Preserve sAllAddr(UBound(sAllAddr) + 1)

                'more?
                ptr1 = Adapter.dwNext


            Loop  'Do While (ptr1 <> 0)

            ReDim Preserve sAllAddr(UBound(sAllAddr))
        End If  'If GetAdaptersInfo
    End If  'If cbRequired > 0

    'return any strings found
    LocalIPAddresses = sAllAddr


End Function

Private Function TrimNull(item As String)

    Dim pos As Integer

    'double check that there is a chr$(0) in the string
    pos = InStr(item, Chr$(0))
    If pos Then
        TrimNull = Left$(item, pos - 1)
        Else: TrimNull = item
    End If

End Function

Public Function FileTypeToImage(ByVal filename As String) As Long
    filename = LCase(filename)
    Select Case FileType(filename)
        Case Is = "txt", "nfo": FileTypeToImage = 2
        Case Is = "mp3", "wav", "au", "wma": FileTypeToImage = 3
        Case Is = "mpg", "mpeg", "avi", "asf", "rm", "ram", "qt": FileTypeToImage = 4
        Case Is = "html", "htm": FileTypeToImage = 5
        Case Is = "jpg", "gif", "bmp", "png": FileTypeToImage = 6
        Case Is = "zip", "rar", "ace": FileTypeToImage = 7
        Case Is = "iso", "bin", "cue": FileTypeToImage = 8
        Case Is = "exe", "dll", "com": FileTypeToImage = 9
        Case Is = "lusernet", "eml", "nws", "vbs", "js": FileTypeToImage = 10
        Case Else: FileTypeToImage = 1
    End Select
End Function

Public Function RandomPlace() As String
    Dim Places()
    Places = Array("at the bar", "at the pub", "in the toilet", "going to hell", "in some random bird's room", "hiding", "somewhere")
    RandomPlace = Places(Int(Rnd * (UBound(Places) + 1)))
End Function

'Public Function isLeecher() As Boolean
'
'    Static beenWarned As Boolean
'
'    isLeecher = False
'
'    If beenWarned Then Exit Function
'
'    Dim downloads As Single
'    Dim uploads As Single
'
'    downloads = (GetSetting("LUSerNet", "Main", "TotalDownload", 0) / 1024) / 1024
'    uploads = (GetSetting("LUSerNet", "Main", "TotalUpload", 0) / 1024) / 1024
'
'    If (downloads > (uploads + 400) * 8) Then
'        isLeecher = True
'        beenWarned = True
'    Else
'        isLeecher = False
'    End If
'
'End Function

Public Sub UpdateTotalDownloads(downloads As Long)
    Dim TotalDownload As Single
    TotalDownload = GetSetting("LUSerNet", "Main", "TotalDownload", 0)
    TotalDownload = TotalDownload + downloads
    SaveSetting "LUSerNet", "Main", "TotalDownload", TotalDownload
    frmMain.lblTotalDownloads.Caption = "Download " & ChangeByte(TotalDownload, True, 0)
End Sub

Public Sub UpdateTotalUploads(uploads As Long)
    Dim totalUpload As Single
    totalUpload = GetSetting("LUSerNet", "Main", "TotalUpload", 0)
    totalUpload = totalUpload + uploads
    SaveSetting "LUSerNet", "Main", "TotalUpload", totalUpload
    frmMain.lblTotalUploads.Caption = "Upload " & ChangeByte(totalUpload, True, 0)
End Sub

'Main sub :)
Public Sub Main()
    frmMain.Show
End Sub

Function isValidIP(IP As String) As Boolean
'Checks if its a valid IP and NOT a broadcast one

    Dim i As Long
    Dim ips() As String
    
    ips = Split(IP, ".")
    If UBound(ips) <> 3 Then GoTo isNotValidIP
    
    For i = 0 To 3
        If Not IsNumeric(ips(i)) Then GoTo isNotValidIP
        'Checks in range 0-254 so it doesn't hit the broadcast range
        If ips(i) < 0 Or ips(i) > 254 Then GoTo isNotValidIP
    Next i
    
    If ips(0) = 0 Then GoTo isNotValidIP

isValidIP = True
Exit Function

isNotValidIP:
isValidIP = False
End Function

Function isNumPositiveLg(number) As Boolean
    Dim tmp As Long
    On Error GoTo isNumPositiveLgError
    tmp = CLng(number)
    isNumPositiveLg = (IsNumeric(number)) And (number >= 0)
    Exit Function
isNumPositiveLgError:
    isNumPositiveLg = False
End Function

Function isNumPositiveDb(number) As Boolean
    Dim tmp As Double
    On Error GoTo isNumPositiveDbError
    tmp = CDbl(number)
    isNumPositiveDb = (IsNumeric(number)) And (number >= 0)
    Exit Function
isNumPositiveDbError:
    isNumPositiveDb = False
End Function
