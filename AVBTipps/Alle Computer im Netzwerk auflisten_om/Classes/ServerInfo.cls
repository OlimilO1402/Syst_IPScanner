VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type SERVER_INFO_API
    PlatformId   As Long
    ServerName   As Long
    Type         As Long
    VerMajor     As Long
    VerMinor     As Long
    Comment      As Long
End Type

Private m_Name  As String

Private m_Users As List '(Of UserInfo)

Declare Function NetUserEnum Lib "netapi32" (lpServer As Any, ByVal Level As Long, ByVal Filter As Long, lpBuffer As Long, ByVal PrefMaxLen As Long, lpEntriesRead As Long, lpTotalEntries As Long, lpResumeHandle As Long) As Long

Friend Sub New_(ServerName As String)
    m_Name = ServerName
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property
