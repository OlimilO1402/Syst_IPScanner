Attribute VB_Name = "NetConnect"
Option Explicit

Private Type WSAdata
   wVersion                 As Integer
   wHighVersion             As Integer
   szDescription(0 To 255)  As Byte
   szSystemStatus(0 To 128) As Byte
   iMaxSockets              As Integer
   iMaxUdpDg                As Integer
   lpVendorInfo             As Long
End Type

Private Type Hostent
   h_name      As Long
   h_aliases   As Long
   h_addrtype  As Integer
   h_length    As Integer
   h_addr_list As Long
End Type

Private Type IP_OPTION_INFORMATION
   TTL         As Byte
   Tos         As Byte
   Flags       As Byte
   OptionsSize As Long
   OptionsData As String * 128
End Type

Private Type IP_ECHO_REPLY
   Address(0 To 3) As Byte
   Status          As Long
   RoundTripTime   As Long
   DataSize        As Integer
   Reserved        As Integer
   data            As Long
   Options         As IP_OPTION_INFORMATION
End Type

Private Declare Function GetHostByName Lib "WSOCK32.DLL" Alias "gethostbyname" (ByVal Hostname As String) As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAdata As WSAdata) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean

Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" _
        (ByVal addr As String, ByVal laenge As Integer, ByVal typ As Integer) As Long
        
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Type HostDeType
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type

Private Const SOCKET_ERROR = 0

Public Function PingA(ByVal sServer As String) As Long

    Dim i As Integer
    Dim Server As String
    
    Server = sServer
    PingA = 0
    For i = 0 To 1
        PingA = PingB(sServer)
        If PingA Then
            Exit For
        
        End If
    
    Next i

End Function

Private Function PingB(ByVal Server As String) As Long
    
    Dim hFile As Long, lpWSAdata As WSAdata
    Dim hHostent As Hostent, AddrList As Long
    Dim Address As Long, rIP As String
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    Dim Hostname As String

    PingB = 0 'Rückgabe anfangs auf null setzen
    If Left(Server, 7) = "http://" Then Server = Mid(Server, 8) 'http:// entfernen

    Call WSAStartup(&H101, lpWSAdata)

    If GetHostByName(Server + String(64 - Len(Server), 0)) <> SOCKET_ERROR Then
        CopyMemory hHostent.h_name, ByVal GetHostByName(Server + String(64 - Len(Server), 0)), Len(hHostent)
        CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
        CopyMemory Address, ByVal AddrList, 4
   
    End If

    hFile = IcmpCreateFile()
    If hFile = 0 Then Exit Function 'Bei Fehler abbrechen

    OptInfo.TTL = 255

    'Ping senden
    If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
        rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
    Else
        'Fehler aufgetreten
        Exit Function
    End If
    
    If EchoReply.Status = 0 Then
        PingB = IIf(EchoReply.RoundTripTime, EchoReply.RoundTripTime, 1)
    Else
        'Keine Antwort bekommen
    End If
    
End Function

Public Function HostByAddress(ByVal Addresse$) As String
    
    Dim x As Integer
    Dim HostDeAddress As Long
    Dim aa As String, BB As String * 5
    Dim HOST As HostDeType
  
    aa = Chr$(Val(NextChar(Addresse, ".")))
    aa = aa + Chr$(Val(NextChar(Addresse, ".")))
    aa = aa + Chr$(Val(NextChar(Addresse, ".")))
    aa = aa + Chr$(Val(Addresse))
    
    HostDeAddress = gethostbyaddr(aa, Len(aa), 2)
    If HostDeAddress = 0 Then
        HostByAddress = ""
        Exit Function
    End If
    
    Call RtlMoveMemory(HOST, HostDeAddress, LenB(HOST))
 
    aa = ""
    x = 0
    Do
       Call RtlMoveMemory(ByVal BB, HOST.hName + x, 1)
       If Left$(BB, 1) = Chr$(0) Then Exit Do
       aa = aa + Left$(BB, 1)
       x = x + 1
    Loop
    
    HostByAddress = aa
End Function

Private Function NextChar(Text$, Char$) As String
    
    Dim POS As Integer
    
    POS = InStr(1, Text, Char)
    If POS = 0 Then
        NextChar = Text
        Text = ""
    Else
        NextChar = Left$(Text, POS - 1)
        Text = Mid$(Text, POS + Len(Char))
    End If
End Function

Public Function GetMacAdresse(IpAdresse As String) As String

    Dim bLanAdapter() As Byte
    Dim i As Long
    Dim numAdapter As Long
    Dim macaddr As String
    
    'Alle Adapter auslesen
    numAdapter = modNetBios.NB_EnumLanAdapter(bLanAdapter)
        
    'wurde mindestens ein aktiver Adapter gefunden
    If numAdapter > 0 Then
        'Für jeden Adapter die MAC-Adresse auslesen
        For i = 1 To numAdapter
            'diesen Adapter initalisieren
            Call modNetBios.NB_ResetAdapter(bLanAdapter(i), 20, 30)
                ' Probieren die MAC-Adresse über diesen Adapter zu ermitteln
                GetMacAdresse = modNetBios.NB_GetMACAddress(bLanAdapter(i), IpAdresse)
                ' Wenn eine MAC-Adresse über diesen Adapter ermittelt wurde,
                ' dann die MAC-Adresse anzeigen
                If Len(GetMacAdresse) > 0 Then
                    Exit For
                    
                End If
        
        Next i
    
    Else
        GetMacAdresse = ""
    
    End If

End Function
