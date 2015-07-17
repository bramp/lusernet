Attribute VB_Name = "mIP"
Option Explicit

Private Const WS_VERSION_REQD As Long = &H101
Private Const SOCKET_ERROR As Long = -1
Private Const AF_INET = 2

Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&

Private Const PING_TIMEOUT = 200        ' number of milliseconds to wait for the reply

Private Type WSADATA
    wVersion As Integer
    wHighVersion  As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type

Private Type IP_OPTION_INFORMATION
    Ttl             As Byte     'Time To Live
    Tos             As Byte     'Type Of Service
    Flags           As Byte     'IP header flags
    OptionsSize     As Byte     'Size in bytes of options data
    OptionsData     As Long     'Pointer to options data
End Type

Private Type ICMP_ECHO_REPLY
    Address         As Long             'Replying address
    Status          As Long             'Reply IP_STATUS, values as defined above
    RoundTripTime   As Long             'RTT in milliseconds
    DataSize        As Integer          'Reply data size in bytes
    Reserved        As Integer          'Reserved for system use
    DataPointer     As Long             'Pointer to the reply data
    Options         As IP_OPTION_INFORMATION    'Reply options
    Data            As String * 250     'Reply data which should be a copy of the string sent, NULL terminated
                                        ' this field length should be large enough to contain the string sent
End Type

Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal VersionReq As Long, WSADataReturn As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal s As String) As Long
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname As String) As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptions As IP_OPTION_INFORMATION, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function GetHostNameFromIP(ByVal sAddress As String) As String
Dim ptrHosent As Long
Dim hAddress As Long
Dim nbytes As Long
   
    If SocketsInitialize() Then
        hAddress = inet_addr(sAddress)
        If hAddress <> SOCKET_ERROR Then
            ptrHosent = gethostbyaddr(hAddress, 4, AF_INET)
            If ptrHosent <> 0 Then
                CopyMemory ptrHosent, ByVal ptrHosent, 4
                nbytes = lstrlen(ByVal ptrHosent)
                If nbytes > 0 Then
                    sAddress = Space$(nbytes)
                    CopyMemory ByVal sAddress, ByVal ptrHosent, nbytes
                    GetHostNameFromIP = sAddress
                End If
            Else
                MsgBox "Call to gethostbyaddr failed."
            End If
            SocketsCleanup
        Else
            MsgBox "String passed is an invalid IP."
        End If
    Else
        MsgBox "Sockets failed to initialize."
    End If
    
End Function

Public Function GetIPFromHostName(ByVal sHostName As String) As String
'converts a host name to an IP address.
Dim nbytes As Long
Dim ptrHosent As Long
Dim ptrName As Long
Dim ptrAddress As Long
Dim ptrIPAddress As Long
Dim sAddress As String

    If SocketsInitialize() Then
        sAddress = Space$(4)
        ptrHosent = gethostbyname(sHostName & vbNullChar)
        If ptrHosent <> 0 Then
            ptrAddress = ptrHosent + 12
            CopyMemory ptrAddress, ByVal ptrAddress, 4
            CopyMemory ptrIPAddress, ByVal ptrAddress, 4
            CopyMemory ByVal sAddress, ByVal ptrIPAddress, 4
            GetIPFromHostName = IPToText(sAddress)
        End If
    Else
        MsgBox "Sockets failed to initialize."
    End If

End Function

Private Function IPToText(ByVal IPAddress As String) As String
   
    IPToText = CStr(Asc(IPAddress)) & "." & CStr(Asc(Mid$(IPAddress, 2, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 3, 1))) & "." & CStr(Asc(Mid$(IPAddress, 4, 1)))
              
End Function

Public Function GetIPAddress() As String
Dim sHostName As String * 256
Dim lpHost As Long
Dim HOST As HOSTENT
Dim dwIPAddr As Long
Dim tmpIPAddr() As Byte
Dim i As Integer
Dim sIPAddr As String
      
    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPAddress = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
            " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    sHostName = Trim$(sHostName)
    lpHost = gethostbyname(sHostName)
    If lpHost = 0 Then
        GetIPAddress = ""
        MsgBox "Windows Sockets are not responding. " & _
            "Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    CopyMemory HOST, lpHost, Len(HOST)
    CopyMemory dwIPAddr, HOST.hAddrList, 4
    
    ReDim tmpIPAddr(0 To HOST.hLen)
    CopyMemory tmpIPAddr(0), dwIPAddr, HOST.hLen
    
    For i = 0 To HOST.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
   
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup

End Function

Public Function GetIPHostName() As String
Dim sHostName As String * 256

    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
    GetIPHostName = ""
    MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
        " has occurred.  Unable to successfully get Host Name."
    SocketsCleanup
    Exit Function
    End If
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup

End Function

Private Function HiByte(ByVal wParam As Integer) As Byte
  
  'note: VB4-32 users should declare this function As Integer
   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function

Private Function LoByte(ByVal wParam As Integer) As Byte

  'note: VB4-32 users should declare this function As Integer
   LoByte = wParam And &HFF&

End Function

Private Sub SocketsCleanup()

    If WSACleanup() <> 0 Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub

Private Function SocketsInitialize() As Boolean
Dim WSAD As WSADATA
Dim sLoByte As String
Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> 0 Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function

Public Sub Ping(sIPAdress As String, LBL_DEST As ListBox)
Dim hFile       As Long
Dim lRet        As Long
Dim lIPAddress  As Long
Dim strMessage  As String
Dim pOptions    As IP_OPTION_INFORMATION
Dim pReturn     As ICMP_ECHO_REPLY
Dim pWsaData    As WSADATA
    
    strMessage = "Echo this string of data"
    
    WSAStartup &H101, pWsaData
    
    lIPAddress = inet_addr(sIPAdress)
    
    '   open up a file handle for doing the ping
    hFile = IcmpCreateFile()
    
    pOptions.Ttl = 1
    
    lRet = IcmpSendEcho(hFile, lIPAddress, strMessage, Len(strMessage), pOptions, _
                        pReturn, Len(pReturn), PING_TIMEOUT)
    If lRet = 0 Then
        LBL_DEST.AddItem "Ping failed with error " & pReturn.Status
        LBL_DEST.ListIndex = LBL_DEST.ListCount - 1
    Else
        If pReturn.Status <> 0 Then
            LBL_DEST.AddItem "Error -> Ping failed to complete, code = " & pReturn.Status
            LBL_DEST.ListIndex = LBL_DEST.ListCount - 1
        Else
            LBL_DEST.AddItem "Success -> completion time is " & pReturn.RoundTripTime & "ms."
            LBL_DEST.ListIndex = LBL_DEST.ListCount - 1
        End If
    End If
                        
    IcmpCloseHandle hFile
    
    WSACleanup
    
End Sub
