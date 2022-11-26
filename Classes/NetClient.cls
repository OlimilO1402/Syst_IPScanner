VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'stores IP-Addresses, MAC-Addresses
'return results from ping etc
'some other stuff

Private m_Name   As String
Private m_Server As String
Private m_NsluResult As String
Private m_PingResult As String
Private m_StatResult As String

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Server() As String
    Server = m_Server
End Property

Public Property Get Result_Nslookup() As String
    Result_Nslookup = m_NsluResult
End Property

Public Property Get Result_Ping() As String
    Result_Ping = m_PingResult
End Property

Public Property Get Result_Netstat() As String
    Result_Netstat = m_StatResult
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


