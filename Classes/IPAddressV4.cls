VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPAddressV4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type TByte4
    b1 As Byte 'kleinster Wert
    b2 As Byte 'nächsthöherer Wert
    b3 As Byte 'nächsthöherer Wert
    b4 As Byte 'höchster Wert
End Type
Private Type BArr4
    bytes(0 To 3) As Byte
End Type
Private Type BLng4
    bytes As Long
End Type
Private Type BCur4
    bytes As Currency
End Type
Private Type TInt2
    i1 As Integer
    i2 As Integer
End Type
Private Type TNetAdress
    b1to4 As TByte4
    i3 As Integer
    i4 As Integer
    i5 As Integer
    i6 As Integer
    i7 As Integer
    i8 As Integer
End Type
Private m_ip As TNetAdress
Private m_Name   As String
Private m_Server As String
Private m_NsluResult As String
Private m_PingResult As String
Private m_StatResult As String
Private m_Port       'As Long 'by default it's missing

'IPAddressV4: 192.168.178.1 'unsinged-Byte1.unsinged-Byte2.unsinged-Byte3.unsinged-Byte4

'             ui21:ui22:ui23:ui24:ui25:ui26:ui27:ui28
'IPAddressV6: 1234:5678:90AB:CDEF:0123:4567:890A:BCDE: UInt16_1:UInt16_2:UInt16_3:UInt16_4:UInt16_5:UInt16_6:UInt16_7:UInt16_8
'MACAddress:   12:34:56:78:AB:CD

Friend Sub New_(StrLngBytesNewAddress, Optional aName As String)
    'StrLngBytesNewAddress: you can either give a string, or a bytearray or a long, or a currency
    m_Name = aName
    Dim va
    Select Case VarType(StrLngBytesNewAddress)
    Case vbCurrency
        Dim cb As BCur4: cb.bytes = StrLngBytesNewAddress / CCur(10000): LSet m_ip = cb
        Exit Sub
    Case vbLong
        Dim lb As BLng4: lb.bytes = StrLngBytesNewAddress: LSet m_ip = lb
        Exit Sub
    Case (vbArray Or vbByte)
        va = StrLngBytesNewAddress
    Case (vbArray Or vbVariant)
        va = StrLngBytesNewAddress
    Case vbString
        va = Split(StrLngBytesNewAddress, ".")
    End Select
    Dim b4 As BArr4, i As Long
    For i = 0 To UBound(va)
        b4.bytes(3 - i) = CByte(va(i))
        If i > 2 Then Exit For
    Next
    LSet m_ip = b4
End Sub
Friend Sub NewC(other As IPAddressV4)
    Debug.Print other.IPToStr
    Me.LAddress = other.LAddress
End Sub
Friend Function Clone() As IPAddressV4
    Set Clone = New IPAddressV4: Clone.NewC Me
End Function

Friend Function IPToStr() As String
    With m_ip
        IPToStr = CStr(.b4) & "." & CStr(.b3) & "." & CStr(.b2) & "." & CStr(.b1)
        If Not IsEmpty(m_Port) Then IPToStr = IPToStr & ":" & CStr(CLng(m_Port))
    End With
End Function

Friend Function ToStr() As String
    ToStr = IPToStr & IIf(Len(m_Name), " " & m_Name, "")
End Function

Friend Function ToInfoStr() As String
    ToInfoStr = Me.Name & vbCrLf & String(Len(m_Name), "=") & vbCrLf & _
                IIf(Len(m_Server), "Server: " & Me.Server & vbCrLf, "") & _
                Me.NsluResult & vbCrLf & _
                Me.PingResult & vbCrLf & _
                Me.StatResult
End Function
Public Property Let LAddress(ByVal Value As Long)
    Dim lb As BLng4: lb.bytes = Value: LSet m_ip = lb
End Property
Public Property Get LAddress() As Long
    Dim lb As BLng4: LSet lb = m_ip: LAddress = lb.bytes
End Property

Friend Function Address() As Currency
    Dim cb As BCur4: LSet cb = m_ip: Address = cb.bytes * CCur(10000)
End Function

Public Property Get Name() As String
    Name = m_Name
End Property
Friend Property Let Name(Value As String)
    m_Name = Value
End Property
Public Property Get Port()
    Port = m_Port
End Property
Public Property Let Port(Value)
    m_Port = Value
End Property
Public Property Get Server() As String
    Server = m_Server
End Property

Public Property Get IsValid() As Boolean
    IsValid = 0 < InStr(1, m_PingResult, "Antwort")
    If Not IsValid Then IsValid = 0 < Len(m_Name)
End Property
Public Property Get Compare(other As IPAddressV4) As Integer
    If other.b1 < Me.b1 Then Compare = 1:  Exit Property
    If Me.b1 < other.b1 Then Compare = -1: Exit Property
    If other.b2 < Me.b2 Then Compare = 1:  Exit Property
    If Me.b2 < other.b2 Then Compare = -1: Exit Property
    If other.b3 < Me.b3 Then Compare = 1:  Exit Property
    If Me.b3 < other.b3 Then Compare = -1: Exit Property
    If other.b4 < Me.b4 Then Compare = 1:  Exit Property
    If Me.b4 < other.b4 Then Compare = -1: Exit Property
End Property

Public Sub OneUp()
    If m_ip.b1 < 255 Then
        m_ip.b1 = m_ip.b1 + 1
    ElseIf m_ip.b2 < 255 Then
        m_ip.b1 = 0
        m_ip.b2 = m_ip.b2 + 1
    ElseIf m_ip.b3 < 255 Then
        m_ip.b1 = 0: m_ip.b2 = 0
        m_ip.b3 = m_ip.b3 + 1
    ElseIf m_ip.b4 < 255 Then
        m_ip.b1 = 0: m_ip.b2 = 0: m_ip.b3 = 0
        m_ip.b4 = m_ip.b4 + 1
    Else
        m_ip.b1 = 0: m_ip.b2 = 0: m_ip.b3 = 0: m_ip.b4 = 0
    End If
End Sub
Public Sub Add(ByVal Sum As Currency)
    Dim cdSrc As BCur4: cdSrc.bytes = Sum / 10000
    Dim curIP As BCur4: LSet curIP = m_ip
    curIP.bytes = curIP.bytes + cdSrc.bytes
    LSet m_ip = curIP
End Sub

Public Property Get PingResult() As String
    PingResult = m_PingResult
End Property
Public Property Get StatResult() As String
    StatResult = m_StatResult
End Property
Public Property Get NsluResult() As String
    NsluResult = m_NsluResult
End Property
Public Property Let b1(ByVal Value As Byte)
    m_ip.b1 = Value
End Property
Public Property Get b1() As Byte
    b1 = m_ip.b1
End Property
Public Property Let b2(ByVal Value As Byte)
    m_ip.b2 = Value
End Property
Public Property Get b2() As Byte
    b2 = m_ip.b2
End Property
Public Property Let b3(ByVal Value As Byte)
    m_ip.b3 = Value
End Property
Public Property Get b3() As Byte
    b3 = m_ip.b3
End Property
Public Property Let b4(ByVal Value As Byte)
    m_ip.b4 = Value
End Property
Public Property Get b4() As Byte
    b4 = m_ip.b4
End Property

Public Sub CallPing(Optional WSShell As Object)
    'Dim WScript As Object: Set WScript = CreateObject("WScript")
    'Dim WSShell As Object: Set WSShell = CreateObject("WScript.Shell")
    If WSShell Is Nothing Then Set WSShell = CreateObject("WScript.Shell")
    Dim WSSExec As Object: Set WSSExec = WSShell.Exec("ping -n 1 -w 100 " & Me.IPToStr)
    m_PingResult = WSSExec.StdOut.ReadAll
    If InStr(m_PingResult, "Empfangen = 1") Then
        'm_Name = ParseNameFromPingResult(m_PingResult)
        Set WSSExec = WSShell.Exec("ping -a " & Me.IPToStr)
        m_StatResult = WSSExec.StdOut.ReadAll
        m_Name = ParseNameFromPingResult(m_StatResult)
    End If
End Sub
Function ParseNameFromPingResult(s As String) As String
    Dim pos1 As Long: pos1 = InStr(1, s, "Ping wird ausgef")
    If pos1 <= 0 Then pos1 = 3
    pos1 = pos1 + 16 + 5 + 4
    Dim pos2 As Long: pos2 = InStr(pos1, s, "[")
    If pos1 < pos2 Then
        ParseNameFromPingResult = Trim(Mid$(s, pos1, pos2 - pos1))
    End If
End Function

'what about nslookup *IP*
Public Sub CallNslookup(Optional WSShell As Object)
Try: On Error GoTo Catch
    If WSShell Is Nothing Then Set WSShell = CreateObject("WScript.Shell")
    Dim WSSExec As Object: Set WSSExec = WSShell.Exec("nslookup " & Me.IPToStr)
    m_NsluResult = WSSExec.StdOut.ReadAll
    ParseNamesFromNslookupResult m_NsluResult
    Exit Sub
Catch:
    If Err Then
        MsgBox Err.Number & " " & Err.Description
    End If
End Sub

Public Sub CallDNS()
Try: On Error GoTo Catch
    Dim sName As String
    Dim rv As Long
    rv = MDns.IP2HostName(Me.IPToStr, sName)
    m_Name = sName
Catch:
    If Err Then
        MsgBox Err.Number & " " & Err.Description
    End If
End Sub

Sub ParseNamesFromNslookupResult(s As String)
    If Len(s) = 0 Then Exit Sub
Try: On Error GoTo Catch
    Dim sLines() As String: sLines = Split(s, vbCrLf)
    Dim sLine    As String
    Dim sElems() As String
    Dim i As Long
    For i = 0 To UBound(sLines)
        sLine = Trim$(sLines(i))
        If Len(sLine) Then
            sElems = Split(sLine, ":")
            If sElems(0) = "Name" Then m_Name = Trim$(sElems(1))
            If sElems(0) = "Server" Then m_Server = Trim$(sElems(1))
            Debug.Print m_Name & " " & m_Server
        End If
    Next
Catch:
End Sub
Public Function ReadFromStr(s As String) As Boolean
    'm_ip As Byte4
    'm_Name   As String
    'm_Server As String
    'm_NsluResult As String
    'm_PingResult As String
    'm_StatResult As String
Try: On Error GoTo Catch
    Dim sl() As String: sl = Split(s, vbTab)
    Dim ub As Long: ub = UBound(sl)
    Dim i As Long
    If ub > 0 Then New_ sl(i)
    i = i + 1: If ub >= i Then m_Name = sl(i)
    i = i + 1: If ub >= i Then m_Server = sl(i)
    i = i + 1: If ub >= i Then m_NsluResult = sl(i)
    i = i + 1: If ub >= i Then m_PingResult = sl(i)
    i = i + 1: If ub >= i Then m_StatResult = sl(i)
        
    Exit Function
Catch:
    
End Function

Public Function WriteToStr() As String
    WriteToStr = Me.IPToStr & vbTab & Me.Name & vbTab & Me.Server & vbTab & Me.NsluResult & vbTab & Me.PingResult & vbTab & Me.StatResult
End Function

