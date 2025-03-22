VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPPingScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event FoundIP(aIP As IPAddress, ByRef out_Cancel As Boolean)

Public Sub Scan(IPBase As IPAddress, ByVal IPFrom As Long, ByVal IPTo As Long)
    Dim WSShell As Object: Set WSShell = CreateObject("WScript.Shell")
    Dim newIP  As IPAddress: Set newIP = IPBase.Clone
    Dim ib As Long
    Dim Cancel As Boolean
    Dim netCli As NetClient: Set netCli = New NetClient
    
    'IPTo = Min(IPTo, 255)
    For ib = IPFrom To IPTo
        netCli.CallNslookup WSShell
        'newIP.CallNslookup WSShell
        RaiseEvent FoundIP(newIP, Cancel)
        If Cancel Then Exit For
        Set newIP = newIP.Clone
        newIP.OneUp
    Next
End Sub

Public Sub Scan2(IPBase As IPAddress, ByVal IPFrom As Long, ByVal IPTo As Long)
    Dim newIP As IPAddress: Set newIP = IPBase.Clone
    Dim ib As Long
    Dim Cancel As Boolean
    'IPTo = Min(IPTo, 255)
    For ib = IPFrom To IPTo
        'newIPV4.CallNslookup WSShell
        newIP.CallDNS
        RaiseEvent FoundIP(newIP, Cancel)
        If Cancel Then Exit For
        Set newIPV4 = newIP.Clone
        newIP.OneUp
    Next
End Sub


'    On Error Resume Next
'    Dim WScript    As Object: Set WScript = CreateObject("WScript")
'    Dim objShell   As Object: Set objShell = CreateObject("WScript.Shell")
'    Dim IPAddrRoot As String: IPAddrRoot = Text1.Text
'    Dim IPAddress  As String
'    Dim PingResult As String
'    Dim StatResult As String
'    Dim objExec    As Object
'    Dim j As Long
'    For j = 0 To 50 '254
'        Label1.Caption = j & " / 254"
'        IPAddress = IPAddrRoot & "." & j
'        Set objExec = objShell.Exec("ping -n 1 -w 100 " & IPAddress)
'        PingResult = objExec.StdOut.ReadAll
'        List1.AddItem IPAddress & ": " & PingResult
'        If InStr(PingResult, "Empfangen = 1") Then
'            List2.AddItem IPAddress
'            'Set objExec = objShell.Exec("nbtstat -A " & IPAddress)
'            Set objExec = objShell.Exec("ping -a " & IPAddress)
'            'Sleep 8000
'            StatResult = objExec.StdOut.ReadAll
'            IPAddresses.Add StatResult, IPAddress
'        End If
'        If j Mod 10 = 0 Then DoEvents
'    Next

