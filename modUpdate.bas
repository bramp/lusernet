Attribute VB_Name = "modUpdate"
'Public Const UpdateTXTURL = "http://www.lusernet.34sp.com/files/"
Public Const UpdateTXTURL = "http://localhost:82/files/"

Private Const MAX_ADAPTER_NAME_LENGTH         As Long = 256
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH  As Long = 128
Private Const MAX_ADAPTER_ADDRESS_LENGTH      As Long = 8

Const ERROR_SUCCESS = 0&

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

Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

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
                sAllAddr(UBound(sAllAddr)) = sIPAddr
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
