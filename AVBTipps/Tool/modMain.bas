Attribute VB_Name = "modMain"
Option Explicit

'Konstanten
Private Const MAX_PREFERRED_LENGTH As Long = -1
Private Const ERROR_SUCCESS        As Long = 0&
Private Const ERROR_ACCESS_DENIED  As Long = 5&
Private Const ERROR_MORE_DATA      As Long = 234&
Private Const ERROR_DHCP_JET_ERROR As Long = 20013&
Private Const NERR_SUCCESS         As Long = 0&

Private Const SUCCESS                   As Long = 1
Private Const DnsFreeRecordList         As Long = 1
Private Const DNS_TYPE_A                As Long = &H1
Private Const DNS_QUERY_BYPASS_CACHE    As Long = &H8

Private Const sckConnected As Long = 7

Private Enum DHCP_SEARCH_INFO_TYPE
    DhcpClientIpAddress = 0&
    DhcpClientHardwareAddress = 1&
    DhcpClientName = 2&
End Enum

'Benutzerdefinierte Typen
Private Type WKSTA_USER_INFO_1
    wkui1_username     As Long
    wkui1_logon_domain As Long
    wkui1_oth_domains  As Long
    wkui1_logon_server As Long
End Type

Public Type WKSTA_USER_INFO_1_STR
    wkui1_username     As String
    wkui1_logon_domain As String
    wkui1_oth_domains  As String
    wkui1_logon_server As String
End Type

Private Type VBDnsRecord
    pNext           As Long
    pName           As Long
    wType           As Integer
    wDataLength     As Integer
    Flags           As Long
    dwTel           As Long
    dwReserved      As Long
    prt             As Long
    others(35)      As Byte
End Type

Private Type DHCPDS_SERVER
    Version       As Long
    ServerName    As Long
    ServerAddress As Long
    Flags         As Long
    State         As Long
    DsLocation    As Long
    DsLocType     As Long
End Type

Private Type DHCPDS_SERVERS
    Flags       As Long
    NumElements As Long
    Servers     As Long 'DHCPDS_SERVER
End Type

Private Type BYTE_IPADDRESS
    Byte0 As Byte
    Byte1 As Byte
    Byte2 As Byte
    Byte3 As Byte
End Type

Private Type DHCP_CLIENT_SEARCH_INFO
    SearchType As DHCP_SEARCH_INFO_TYPE
    ClientData As Long
End Type

Private Type DHCP_BINARY_DATA
    DataLength As Long
    Data       As Long
End Type

Private Type DATE_TIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Type DHCP_HOST_INFO
    IpAddress   As Long
    NetBiosName As Long
    HostName    As Long
End Type

Private Type DHCP_CLIENT_INFO
    ClientIpAddress    As Long
    SubnetMask         As Long
    ClientHardwareAddress As DHCP_BINARY_DATA
    ClientName         As Long
    ClientComment      As Long
    ClientLeaseExpires As DATE_TIME
    OwnerHost          As DHCP_HOST_INFO
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" (ByRef pTo As Any, ByRef uFrom As Any, ByVal lSize As Long)
Private Declare Function StrLenW Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlen Lib "kernel32" (ByVal straddress As Long) As Long

Private Declare Function NetWkstaUserEnum Lib "netapi32" (ByVal ServerName As Long, ByVal Level As Long, ByRef bufptr As Long, ByVal prefmaxlen As Long, ByRef entriesread As Long, ByRef totalentries As Long, ByRef resume_handle As Long) As Long
Declare Function NetApiBufferFree Lib "netapi32" (ByVal lBuffer As Long) As Long
Declare Function NetGetDCName Lib "netapi32" (ByRef lpServer As Any, ByRef lpDomain As Any, ByRef vBuffer As Any) As Long

Private Declare Function DnsQuery Lib "dnsapi" Alias "DnsQuery_A" (ByVal strname As String, ByVal wType As Integer, ByVal fOptions As Long, ByVal pServers As Long, ByRef ppQueryResultsSet As Long, ByVal pReserved As Long) As Long
Private Declare Function DnsRecordListFree Lib "dnsapi" (ByVal pDnsRecord As Long, ByVal FreeType As Long) As Long

       
Private Declare Function inet_ntoa Lib "ws2_32" (ByVal pIP As Long) As Long
Private Declare Function inet_addr Lib "ws2_32" (ByVal sAddr As String) As Long

Private Declare Function GetRTTAndHopCount Lib "iphlpapi" (ByVal lDestIPAddr As Long, ByRef lHopCount As Long, ByVal lMaxHops As Long, ByRef lRTT As Long) As Long
         
Private Declare Function DhcpEnumServers Lib "dhcpsapi" (ByVal Flags As Long, ByVal IdInfo As Long, ByRef Servers As Long, ByVal CallbackFn As Long, ByVal CallbackData As Long) As Long
Private Declare Function DhcpGetClientInfo Lib "dhcpsapi" (ByVal ServerIpAddress As Long, ByVal SearchInfo As Long, ByRef ClientInfo As Long) As Long
Private Declare Sub DhcpRpcFreeMemory Lib "dhcpsapi" (ByVal BufferPointer As Long)
         
'Variablen
Public Users() As WKSTA_USER_INFO_1_STR
Private DhcpServer() As String
         
Public Function ListDhcpServer() As Boolean

    Dim bRet As Boolean
    Dim lItem As Long
    Dim pServers As Long
    Dim tDSS  As DHCPDS_SERVERS
    Dim tDS() As DHCPDS_SERVER
    
    If DhcpEnumServers(0&, 0&, pServers, 0&, 0&) = ERROR_SUCCESS Then
    
        Call CopyMemory(tDSS, ByVal pServers, LenB(tDSS))
        
        ReDim tDS(tDSS.NumElements - 1)
        ReDim DhcpServer(tDSS.NumElements - 1)
        
        Call CopyMemory(tDS(0), ByVal tDSS.Servers, LenB(tDS(0)) * tDSS.NumElements)
        
        For lItem = 0 To tDSS.NumElements - 1
            
            DhcpServer(lItem) = PtrStr(tDS(lItem).ServerName)
            
        Next
    
        bRet = True
        
    Else
    
        bRet = False
    
    End If

    ListDhcpServer = bRet
End Function

Public Function GetMacFromClient(ByVal sClientIP As String) As String

    Dim lItem As Long
    Dim lCount As Long
    Dim lRet As Long
    Dim sRet As String
    Dim pDhcpServer As Long
    Dim pClientData As Long
    Dim pSearchInfo As Long
    Dim pBuffer As Long
    Dim tDCI As DHCP_CLIENT_INFO
    Dim tDCSI As DHCP_CLIENT_SEARCH_INFO

    If Len(sClientIP) <> 0 Then

    lCount = UBound(DhcpServer)

    For lItem = 0 To lCount

        pDhcpServer = StrPtr(DhcpServer(lItem))
        pClientData = PtrFromIP(sClientIP)

        With tDCSI

            .SearchType = DhcpClientIpAddress
            .ClientData = pClientData

        End With

        pSearchInfo = VarPtr(tDCSI)

        lRet = DhcpGetClientInfo(pDhcpServer, pSearchInfo, pBuffer)

        If lRet = ERROR_SUCCESS Then

            Call CopyMemory(tDCI, ByVal pBuffer, LenB(tDCI))

            sRet = MacFromPtr(tDCI.ClientHardwareAddress)

            Call DhcpRpcFreeMemory(pBuffer)

            Exit For

        Else
        
            sRet = vbNullString
            
        End If

    Next

    End If
    
    GetMacFromClient = sRet

End Function
         
Public Sub FormPosition_Get(ByRef F As Form)

    Dim buf As String
    Dim l As Integer, t As Integer
    Dim h As Integer, w As Integer
    Dim pos As Integer

    buf = GetSetting("AvBremenLV", "FormPosition", "Position", "")

    If buf = "" Then
        
        F.Move (Screen.Width - F.Width) \ 2, (Screen.Height - F.Height) \ 2
    Else
        pos = InStr(buf, ",")
        l = CInt(Left(buf, pos - 1))
        buf = Mid(buf, pos + 1)
        pos = InStr(buf, ",")
        t = CInt(Left(buf, pos - 1))
        buf = Mid(buf, pos + 1)
        pos = InStr(buf, ",")
        w = CInt(Left(buf, pos - 1))
        h = CInt(Mid(buf, pos + 1))
        F.Move l, t, w, h
    End If
End Sub

Public Sub FormPosition_Put(ByRef F As Form)
    Dim buf As String
    buf = F.Left & "," & F.Top & "," & F.Width & "," & F.Height
    SaveSetting "AvBremenLV", "FormPosition", "Position", buf
End Sub
         
Public Function SimplePing(ByVal sIPadr As String, Optional ByVal lMaxHops As Long = 1) As Boolean

    Dim lIPadr      As Long
    Dim lHopsCount  As Long
    Dim lRTT        As Long
    'Dim lMaxHops    As Long
    
    'lMaxHops = 1
    
    SimplePing = CheckPort(sIPadr, 445)
    
'    lIPadr = inet_addr(sIPadr)
'
'    If GetRTTAndHopCount(lIPadr, lHopsCount, lMaxHops, lRTT) = SUCCESS Then
'
'        SimplePing = True
'
'    Else
'
'        SimplePing = False
'
'    End If
    
End Function
         
Public Function GetPDCName() As String

    Dim lpBuffer As Long
    Dim nRet As Long
    Dim yServer() As Byte
    Dim sLocal As String
    
    yServer = MakeServerName(vbNullString)
    
    nRet = NetGetDCName(yServer(0), yServer(0), lpBuffer)
    
    If nRet = 0 Then
        sLocal = PointerToStringW(lpBuffer)
    End If
    
    If lpBuffer Then Call NetApiBufferFree(lpBuffer)
    
    GetPDCName = sLocal
    
End Function

Public Function MakeServerName(ByVal ServerName As String)

    Dim yServer() As Byte
    
    If ServerName <> "" Then
        If InStr(1, ServerName, "\\") = 0 Then
            ServerName = "\\" & ServerName
        End If
    End If
    
    yServer = ServerName & vbNullChar
    MakeServerName = yServer
    
End Function

Public Function PointerToStringW(lpStringW As Long) As String

    Dim buffer() As Byte
    Dim nLen As Long
    
    If lpStringW Then
        nLen = lstrlenW(lpStringW) * 2
        
        If nLen Then
            ReDim buffer(0 To (nLen - 1)) As Byte
            CopyMemory2 buffer(0), ByVal lpStringW, nLen
            PointerToStringW = buffer
        End If
    End If
    
End Function

Public Function HostName2IP(ByVal sAddr As String, Optional sDnsServers As String) As String
    
    Dim pRecord     As Long
    Dim pNext       As Long
    Dim uRecord     As VBDnsRecord
    Dim lPtr        As Long
    Dim vSplit      As Variant
    Dim laServers() As Long
    Dim pServers    As Long
    Dim sName       As String

    If LenB(sDnsServers) <> 0 Then
    
        vSplit = Split(sDnsServers)
        
        ReDim laServers(0 To UBound(vSplit) + 1)
        
        laServers(0) = UBound(laServers)
        
        For lPtr = 0 To UBound(vSplit)
            laServers(lPtr + 1) = inet_addr(vSplit(lPtr))
        Next
        
        pServers = VarPtr(laServers(0))
    End If
    
    If DnsQuery(sAddr, DNS_TYPE_A, DNS_QUERY_BYPASS_CACHE, pServers, pRecord, 0) = 0 Then
        
        pNext = pRecord
        
        Do While pNext <> 0
            
            Call CopyMemory(uRecord, pNext, Len(uRecord))
            
            If uRecord.wType = DNS_TYPE_A Then
                
                lPtr = inet_ntoa(uRecord.prt)
                sName = String(lstrlen(lPtr), 0)
                
                Call CopyMemory(ByVal sName, lPtr, Len(sName))
                
                If LenB(HostName2IP) <> 0 Then
                    HostName2IP = HostName2IP & " "
                End If
                
                HostName2IP = HostName2IP & sName
            
            End If
            
            pNext = uRecord.pNext
        
            DoEvents
            
        Loop
        
        Call DnsRecordListFree(pRecord, DnsFreeRecordList)
    
    End If
End Function

Public Function LoggedOnUser(strServer As String) As Long
    
  Dim bufptr          As Long
  Dim dwServer        As Long
  Dim dwEntriesread   As Long
  Dim dwTotalentries  As Long
  Dim dwResumehandle  As Long
  Dim nStatus         As Long
  Dim nStructSize     As Long
  Dim cnt             As Long
  Dim wui1            As WKSTA_USER_INFO_1
  
  'strServer muß mit "\\" beginnen
  'bServer = strServer & vbNullString
  dwServer = StrPtr(strServer)
  
  Do
    'PC Connecten und Liste der angemeldeten User abfragen
    'MAX_PREFERRED_LENGTH bewirkt das die NetApi32 den BufferSize
    'selber bestimmt und den Buffer Allociert
    'Dieser Aufruf erzwingt die Struktur Level 1, alternativ kann
    'auch Level 0 genutzt werden der nur den Benutzernamen ermittelt
    nStatus = NetWkstaUserEnum(dwServer, 1, bufptr, MAX_PREFERRED_LENGTH, _
      dwEntriesread, dwTotalentries, dwResumehandle)
    
    ReDim Users(dwTotalentries)
    
    'wieviel insgesamt
    If nStatus = NERR_SUCCESS Or nStatus = ERROR_MORE_DATA Then
      
      If dwEntriesread > 0 Then
        
        ' Länge ermitteln damit die richtige Anzahl Bytes aus dem Speicher kopiert wird
        nStructSize = LenB(wui1)
        
        For cnt = 0 To dwEntriesread - 1
          
          'Alle gelesenen User in die Struktur kopieren
           CopyMemory wui1, ByVal bufptr + (nStructSize * cnt), nStructSize
           
           'Alle Stringpointer als Strings in die neue Struktur kpoieren
           Users(cnt).wkui1_username = PtrStr(wui1.wkui1_username)
           Users(cnt).wkui1_logon_domain = PtrStr(wui1.wkui1_logon_domain)
           Users(cnt).wkui1_logon_server = PtrStr(wui1.wkui1_logon_server)
           Users(cnt).wkui1_oth_domains = PtrStr(wui1.wkui1_oth_domains)
           
           DoEvents
        
        Next cnt
      
      End If
    
    Else
    
      LoggedOnUser = nStatus
    
    End If
  Loop While nStatus = ERROR_MORE_DATA
  
  NetApiBufferFree bufptr
End Function

Private Function PtrStr(ByVal lpString As Long) As String
  Dim buff() As Byte
  Dim nSize As Long
  
  'Pointer benutzen um Strings aus Speicher zu kopieren
  If lpString Then
    
    'its Unicode, so mult. by 2
    nSize = StrLenW(lpString) * 2
    If nSize Then
      ReDim buff(0 To (nSize - 1)) As Byte
      CopyMemory buff(0), lpString, nSize
      PtrStr = buff
    End If
  End If
End Function

Private Function CheckPort(ByVal Server As String, ByVal Port As Long) As Boolean

    Dim bolRet As Boolean
    Dim SockObject As Object

    Set SockObject = CreateObject("MSWinsock.Winsock.1")

    SockObject.Protocol = 0
    SockObject.Close
    SockObject.Connect Server, Port

    Call Pause(0.02)
    
    Select Case SockObject.State

    Case sckConnected
        bolRet = True

    Case Else
        bolRet = False

    End Select

    Call SockObject.Close

    Set SockObject = Nothing

    CheckPort = bolRet

End Function

Private Sub SleepLong(ByVal lngSeconds As Long)

    Dim t As Single, b As Boolean

    t = Timer

    Do

        Sleep 1

        DoEvents

        b = Timer - t > lngSeconds

    Loop Until b

End Sub

Public Sub Pause(ByVal SecsDelay As Single)

    Dim TimeOut   As Single
    Dim PrevTimer As Single

    PrevTimer = Timer
    TimeOut = PrevTimer + SecsDelay

    Do While PrevTimer < TimeOut

        Sleep 1
        DoEvents

        If Timer < PrevTimer Then TimeOut = TimeOut - 86400
        PrevTimer = Timer

    Loop

End Sub

Private Function PtrFromIP(ByVal sIP As String) As Long

    Dim tIP As BYTE_IPADDRESS
    Dim sSplit() As String
    Dim lPtr As Long
    
    sSplit = Split(sIP, ".")
    
    tIP.Byte3 = CByte(sSplit(0))
    tIP.Byte2 = CByte(sSplit(1))
    tIP.Byte1 = CByte(sSplit(2))
    tIP.Byte0 = CByte(sSplit(3))

    Call CopyMemory2(lPtr, tIP, LenB(tIP))

    PtrFromIP = lPtr

End Function

Private Function MacFromPtr(ByRef DBD As DHCP_BINARY_DATA, Optional ByVal sSep As _
    String = ":") As String

    Dim bMac() As Byte
    Dim lItem As Long
    Dim sTmp As String
    Dim sTmp2 As String

    ReDim bMac(DBD.DataLength - 1)

    Call CopyMemory(bMac(0), ByVal DBD.Data, DBD.DataLength)

    For lItem = 0 To DBD.DataLength - 1
        
        sTmp2 = CStr(Hex(bMac(lItem)))
        
        If Len(sTmp2) = 1 Then
        
            sTmp2 = "0" & sTmp2
            
        End If
        
        sTmp = sTmp & sTmp2

        If lItem <> DBD.DataLength - 1 Then

            sTmp = sTmp & sSep

        End If

    Next lItem

    MacFromPtr = sTmp

End Function

