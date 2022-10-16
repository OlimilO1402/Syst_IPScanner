VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "www.activevb.de"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   1830
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Local IPs"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

Option Explicit
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSAData As WinSocketDataType) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal HostName As String, ByVal HostLen As Integer) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal HostName As String) As Long
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" (ByVal addr As String, ByVal laenge As Integer, ByVal typ As Integer) As Long
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

Private Function GetIPs() As Collection
    Dim List As Collection: Set List = New Collection
    Dim IP As String, Host As String
    Dim x As Integer
    
    Call InitSocketAPI
    Host = MyHostName
    'List1.Clear
    
    Do
        IP = HostByName(Host, x)
        If Len(IP) <> 0 Then List.Add IP
        x = x + 1
    Loop While Len(IP) > 0
    
    Call CleanSockets
    Set GetIPs = List
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

Private Function MyHostName() As String
    Dim HostName As String * 256
    
    If gethostname(HostName, 256) = SOCKET_ERROR Then
        MsgBox "Windows Sockets error " & Str(WSAGetLastError())
        Exit Function
    Else
        MyHostName = NextChar(Trim$(HostName), Chr$(0))
    End If
End Function

Private Function HostByName(Name As String, Optional x As Integer = 0) As String
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
    
    For y = 0 To x
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
        Call MsgBox("Socket Error " & Trim$(Str$(Result)) & " in Prozedur 'CleanSockets' aufgetreten !")
        
        End
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

Private Sub Form_Load()
    Dim List As Collection: Set List = GetIPs
    Dim i As Long
    For i = 1 To List.Count
        List1.AddItem List.Item(i)
    Next
End Sub
