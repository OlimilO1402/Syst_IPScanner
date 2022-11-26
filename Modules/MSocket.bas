Attribute VB_Name = "MSocket"
'Dieser Quellcode stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!

'------------- Anfang Projektdatei Project1.vbp -------------
'--------- Anfang Formular "Form1" alias Form1.frm  ---------
' Steuerelement: Listen-Steuerelement "List1"
' Steuerelement: Beschriftungsfeld "Label1"

Option Explicit
'wsock32.dll
'Ws2_32.dll
Private Declare Function WSAGetLastError Lib "Ws2_32" () As Long
Private Declare Function WSAStartup Lib "Ws2_32" (ByVal wVersionRequired As Long, lpWSAData As WinSocketDataType) As Long
Private Declare Function WSACleanup Lib "Ws2_32" () As Long
Private Declare Function gethostname Lib "Ws2_32" (ByVal HostName As String, ByVal HostLen As Integer) As Long
Private Declare Function gethostbyname Lib "Ws2_32" (ByVal HostName As String) As Long
Private Declare Function gethostbyaddr Lib "Ws2_32" (ByVal addr As String, ByVal laenge As Integer, ByVal typ As Integer) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Type HostDeType
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type

Const WS_VERSION_REQD    As Long = &H101&
Const MIN_SOCKETS_REQD   As Long = 1&
Const SOCKET_ERROR       As Long = -1&
Const WSADescription_Len As Long = 256&
Const WSASYS_Status_Len  As Long = 128&

Private Type WinSocketDataType
    wversion       As Integer
    wHighVersion   As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets    As Integer
    iMaxUdpDg      As Integer
    lpszVendorInfo As Long
End Type

Public Function GetMyIP() As IpAddress
    Dim sIP As String
    Dim L As List: Set L = GetIPs
    sIP = L.First
    Set GetMyIP = MNew.IPAddressV4(sIP)
End Function

Public Function GetIPs() As List
    Set GetIPs = MNew.List(vbString)
    Dim IP As String, Host As String
    Dim X As Integer
    
    InitSocketAPI
    Host = MyHostName
    
    Do
        IP = HostByName(Host, X)
        If Len(IP) <> 0 Then GetIPs.Add IP
        X = X + 1
    Loop While Len(IP) > 0
    
    CleanSockets
End Function

Private Sub InitSocketAPI()
    Dim Result As Integer
    Dim SocketData As WinSocketDataType
  
    Result = WSAStartup(WS_VERSION_REQD, SocketData)
    If Result <> 0 Then
        Call MsgBox("'winsock.dll' antwortet nicht !")
        End
    End If
End Sub

Public Function MyHostName() As String
    Dim HostName As String * 256
    
    If gethostname(HostName, 256) = SOCKET_ERROR Then
        MsgBox "Windows Sockets error " & str(WSAGetLastError())
        Exit Function
    Else
        MyHostName = NextChar(Trim$(HostName), Chr$(0))
    End If
End Function

Private Function HostByName(Name As String, Optional X As Integer = 0) As String
    Dim MemIp() As Byte
    Dim y As Integer
    Dim HostDeAddress As Long, HostIp As Long
    Dim IpAddress As String
    Dim Host As HostDeType
  
    HostDeAddress = gethostbyname(Name)
    If HostDeAddress = 0 Then
        HostByName = ""
        Exit Function
    End If
    
    Call RtlMoveMemory(Host, HostDeAddress, LenB(Host))
    
    For y = 0 To X
        Call RtlMoveMemory(HostIp, Host.hAddrList + 4 * y, 4)
        If HostIp = 0 Then
            HostByName = ""
            Exit Function
        End If
    Next y
    
    ReDim MemIp(1 To Host.hLength)
    Call RtlMoveMemory(MemIp(1), HostIp, Host.hLength)
    
    IpAddress = ""
    
    For y = 1 To Host.hLength
        IpAddress = IpAddress & MemIp(y) & "."
    Next y
    
    IpAddress = Left$(IpAddress, Len(IpAddress) - 1)
    HostByName = IpAddress
End Function

Private Sub CleanSockets()
    Dim Result As Long
    
    Result = WSACleanup()
    If Result <> 0 Then
        Call MsgBox("Socket Error " & Trim$(str$(Result)) & " in Prozedur 'CleanSockets' aufgetreten !")
        'End
    End If
End Sub

Private Function NextChar(Text As String, Char As String) As String
    Dim pos As Integer
    
    pos = InStr(1, Text, Char)
    If pos = 0 Then
        NextChar = Text
        Text = ""
    Else
        NextChar = Left$(Text, pos - 1)
        Text = Mid$(Text, pos + Len(Char))
    End If
End Function
