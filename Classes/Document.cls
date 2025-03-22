VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_IPBase      As IPAddress
Private m_LastIP      As IPAddress
Private m_SearchNIPs  As Long
Private m_StartIPb4   As Long 'was is das gleich wieder?

'Private m_IPAddresses As IPAddresses
Private m_NetClients  As List 'Of NetClient

Private m_FileName    As String
Private m_Changed     As Boolean
Public WithEvents IPPingScanner As IPPingScanner
Attribute IPPingScanner.VB_VarHelpID = -1

Private Sub Class_Initialize()
    'Set m_IPAddresses = New IPAddresses
    Set m_NetClients = MNew.List(vbObject)
End Sub

Friend Sub New_(aIPBase As IPAddress, aLastIP As IPAddress, nSearchIPs As Long, aScanner As IPPingScanner) ', StartIPb4 As Long)
    Set m_IPBase = aIPBase: Set m_LastIP = aLastIP: m_SearchNIPs = nSearchIPs: Set IPPingScanner = aScanner
End Sub

Public Sub Scan()
    IPPingScanner.Scan m_IPBase, m_StartIPb4, m_StartIPb4 + m_SearchNIPs - 1
End Sub

Public Property Let FileName(ByVal Value As String)
    m_FileName = Value
End Property
Public Property Get FileName() As String
    FileName = m_FileName
End Property

'Public Property Get IPAddresses() As IPAddresses
'    Set IPAddresses = m_IPAddresses
'End Property

Public Property Get IPBase() As IPAddress
    Set IPBase = m_IPBase
End Property
Public Property Set IPBase(Value As IPAddress)
    Set m_IPBase = Value
    m_Changed = True
End Property

Public Property Get LastIP() As IPAddress
    Set LastIP = m_LastIP
End Property
Public Property Set LastIP(Value As IPAddress)
    Set m_LastIP = Value
    m_Changed = True
End Property

Public Property Get SearchNIPs() As Long
    SearchNIPs = m_SearchNIPs
End Property
Public Property Let SearchNIPs(ByVal Value As Long)
    m_SearchNIPs = Value
    m_Changed = True
End Property

Public Property Get StartIPb4() As Long
    StartIPb4 = m_StartIPb4
End Property
Public Property Let StartIPb4(ByVal Value As Long)
    m_StartIPb4 = Value
    m_Changed = True
End Property

'Public Sub IPAddresses_Add(aIP As IpAddress) 'As IPAddress
'    m_IPAddresses.Add aIP
'    m_Changed = True
'End Sub

Public Sub NetClients_Add(aNC As NetClient)
    m_NetClients.Add aNC
    m_Changed = True
End Sub

Public Property Let IsDataChanged(ByVal Value As Boolean)
    m_Changed = Value
End Property
Public Property Get IsDataChanged() As Boolean
    IsDataChanged = m_Changed
End Property

Private Sub IPPingScanner_FoundIP(aIPV4 As IPAddress, out_Cancel As Boolean)
    'If Not m_IPAddresses.Contains(aIPV4.IPToStr) Then
        '
    End If
End Sub
