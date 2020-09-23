VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A ping implementation (Class module)"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

 ' MultiUse = -1  'True
 ' Persistable = 0  'NotPersistable
 ' DataBindingBehavior = 0  'vbNone
 ' DataSourceBehavior = 0   'vbNone
 ' MTSTransactionMode = 0   'NotAnMTSObject
'End
Option Explicit
'-----CONSTANT DECLARATION-----
Private Const IP_STATUS_BASE = 11000
Private Const IP_SUCCESS = 0
Private Const IP_BUF_TOO_SMALL = (IP_STATUS_BASE + 1)
Private Const IP_DEST_NET_UNREACHABLE = (IP_STATUS_BASE + 2)
Private Const IP_DEST_HOST_UNREACHABLE = (IP_STATUS_BASE + 3)
Private Const IP_DEST_PROT_UNREACHABLE = (IP_STATUS_BASE + 4)
Private Const IP_DEST_PORT_UNREACHABLE = (IP_STATUS_BASE + 5)
Private Const IP_NO_RESOURCES = (IP_STATUS_BASE + 6)
Private Const IP_BAD_OPTION = (IP_STATUS_BASE + 7)
Private Const IP_HW_ERROR = (IP_STATUS_BASE + 8)
Private Const IP_PACKET_TOO_BIG = (IP_STATUS_BASE + 9)
Private Const IP_REQ_TIMED_OUT = (IP_STATUS_BASE + 10)
Private Const IP_BAD_REQ = (IP_STATUS_BASE + 11)
Private Const IP_BAD_ROUTE = (IP_STATUS_BASE + 12)
Private Const IP_TTL_EXPIRED_TRANSIT = (IP_STATUS_BASE + 13)
Private Const IP_TTL_EXPIRED_REASSEM = (IP_STATUS_BASE + 14)
Private Const IP_PARAM_PROBLEM = (IP_STATUS_BASE + 15)
Private Const IP_SOURCE_QUENCH = (IP_STATUS_BASE + 16)
Private Const IP_OPTION_TOO_BIG = (IP_STATUS_BASE + 17)
Private Const IP_BAD_DESTINATION = (IP_STATUS_BASE + 18)
Private Const IP_ADDR_DELETED = (IP_STATUS_BASE + 19)
Private Const IP_SPEC_MTU_CHANGE = (IP_STATUS_BASE + 20)
Private Const IP_MTU_CHANGE = (IP_STATUS_BASE + 21)
Private Const IP_UNLOAD = (IP_STATUS_BASE + 22)
Private Const IP_ADDR_ADDED = (IP_STATUS_BASE + 23)
Private Const IP_GENERAL_FAILURE = (IP_STATUS_BASE + 50)
Private Const MAX_IP_STATUS = IP_STATUS_BASE + 50
Private Const IP_PENDING = (IP_STATUS_BASE + 255)
Private Const PING_TIMEOUT = 10000
Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYSSTATUS_LEN = 256
Private Const WSADESCRIPTION_LEN_1 = WSADESCRIPTION_LEN + 1
Private Const WSASYSSTATUS_LEN_1 = WSASYSSTATUS_LEN + 1
Private Const SOCKET_ERROR = -1

'-----TYPE DECLARATION-----
Private Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Byte
    OptionsData As Long
End Type

Private Type ICMP_ECHO_REPLY
    Address As Long
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    DataPointer As Long
    Options As IP_OPTION_INFORMATION
    Data As String * 128
End Type

Private Type Inet_Address
    Byte4 As String * 1
    Byte3 As String * 1
    Byte2 As String * 1
    Byte1 As String * 1
End Type

Private Type hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type tagWSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN_1
    szSystemStatus As String * WSASYSSTATUS_LEN_1
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As String * 200
End Type

'-----PRIVATE FUNCTION DECLARATION-----
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptions As IP_OPTION_INFORMATION, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequested As Integer, lpWSAData As tagWSAData) As Integer
Private Declare Function WSACleanup Lib "wsock32" () As Integer
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal addr As String) As Long
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal addr As Long) As Long
Private Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal HostName As String) As Long
Private Declare Function GetHostName Lib "wsock32.dll" Alias "gethostname" (ByVal HostName As String, HostLen As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'-----MISC DECLARATIONS-----
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Private IPLong As Inet_Address
Public Event PingError(peErrDescription As String)
Private m_lastIP As String
Public Event PingOK(poSpeed As Long)

Public Function DoPing(ByVal HostAddress As String) As Boolean


    Dim hFile As Long
    Dim lRet As Long
    Dim lIPAddress As Long
    Dim strMessage As String
    Dim pOptions As IP_OPTION_INFORMATION
    Dim pReturn As ICMP_ECHO_REPLY
    Dim iVal As Integer
    Dim lPingRet As Long
    Dim pWsaData As tagWSAData
    Dim buffsize As Long
    Dim TTL As Byte
    
    
    buffsize = 128
    TTL = 255
    strMessage = "ICMP ECHO DATA"
    iVal = WSAStartup(&H101, pWsaData)
    lIPAddress = inet_addr(ICMPGetHostByName(HostAddress))
    hFile = IcmpCreateFile()
    pOptions.TTL = TTL
    lRet = IcmpSendEcho(hFile, lIPAddress, strMessage, Len(strMessage), pOptions, pReturn, Len(pReturn), PING_TIMEOUT)


    If lRet = 0 Then


        DoPing = False
            RaiseEvent PingError("PING failure! Error = " & pReturn.Status)
        Else


            If pReturn.Status <> 0 Then


                DoPing = False
                    RaiseEvent PingError("PING failure! Error = " & pReturn.Status)
                Else


                    DoPing = True
                        RaiseEvent PingOK(pReturn.RoundTripTime)
                    End If

                End If

                lRet = IcmpCloseHandle(hFile)
                iVal = WSACleanup()
End Function

Public Function ICMPGetHostByName(Host As String) As String

    Dim szString As String
    Dim sMsg As String
    Dim PointerToPointer As Long
    Dim hostent As hostent
    Dim ListAddress As Long
    Dim ListAddr As Long
    Dim Address As Long
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))
    PointerToPointer = GetHostByName(Host)


    If PointerToPointer = -1 Then
        ICMPGetHostByName = "0"
    Else


        If PointerToPointer <> 0 Then
            CopyMemory hostent.h_name, ByVal PointerToPointer, Len(hostent) ' Copy Winsock structure to the VisualBasic structure
            ListAddress = hostent.h_addr_list ' Get the ListAddress of the Address List
            CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure to the VisualBasic structure
            Dim barray(128) As Byte
            CopyMemory barray(1), ByVal ListAddr, 32
            Dim i As Integer
            CopyMemory Address, ByVal ListAddr, 4 ' Get the first list entry from the Address List
            m_lastIP = vbInet_Ntoa(Address)
        Else
            m_lastIP = "Not There"
        End If

    End If

    ICMPGetHostByName = m_lastIP
End Function



Private Function vbInet_Ntoa(Address As Long) As String

    CopyMemory IPLong, Address, 4
    vbInet_Ntoa = CStr(Asc(IPLong.Byte4)) + "." + CStr(Asc(IPLong.Byte3)) + "." + CStr(Asc(IPLong.Byte2)) + "." + CStr(Asc(IPLong.Byte1))
End Function



Public Property Get LastIP() As String
    LastIP = m_lastIP
End Property

