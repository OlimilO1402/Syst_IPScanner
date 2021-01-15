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
Private Type BArr4
    bytes(0 To 3) As Byte
End Type
Private Type BLng4
    bytes As Long
End Type
Private Type BCur4
    bytes As Currency
End Type
Private Type Byte4
    b1 As Byte
    b2 As Byte
    b3 As Byte
    b4 As Byte
End Type
Private m_ip As Byte4
Private m_Name As String
Private m_PingResult As String
Private m_StatResult As String

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
        b4.bytes(i) = CByte(va(i))
        If i > 2 Then Exit For
    Next
    LSet m_ip = b4
End Sub
Friend Sub NewC(other As IPAddressV4)
    Me.LAddress = other.LAddress
End Sub
Friend Function Clone() As IPAddressV4
    Set Clone = New IPAddressV4: Clone.NewC Me
End Function

Friend Function IPToStr() As String
    With m_ip
        IPToStr = CStr(.b1) & "." & CStr(.b2) & "." & CStr(.b3) & "." & CStr(.b4)
    End With
End Function

Friend Function ToStr() As String
    ToStr = IPToStr & IIf(Len(m_Name), " " & m_Name, "")
End Function

Friend Function ToInfoStr() As String
    ToInfoStr = Me.Name & vbCrLf & String(Len(m_Name), "=") & vbCrLf & _
                Me.PingResult & vbCrLf & Me.StatResult
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

Public Property Get IsActive() As Boolean
    IsActive = 0 < InStr(1, m_PingResult, "Antwort")
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

Public Property Get PingResult() As String
    PingResult = m_PingResult
End Property
Public Property Get StatResult() As String
    StatResult = m_StatResult
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
    If WSShell Is Nothing Then Set WSShell = CreateObject("WScript.Shell")
    'Dim WScript As Object: Set WScript = CreateObject("WScript")
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

