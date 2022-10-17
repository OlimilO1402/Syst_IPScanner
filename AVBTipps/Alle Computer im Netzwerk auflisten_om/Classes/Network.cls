VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Network"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Servers As List

'https://learn.microsoft.com/en-us/windows/win32/api/lmserver/nf-lmserver-netserverenum
Declare Function NetServerEnum Lib "netapi32" (lpServer As Any, ByVal lLevel As Long, vBuffer As Any, lPreferedMaxLen As Long, lEntriesRead As Long, lTotalEntries As Long, ByVal lServerType As Long, ByVal sDomain As String, vResume As Any) As Long
'NET_API_STATUS NET_API_FUNCTION NetServerEnum(
'  [in, optional]      LMCSTR  servername,
'  [in]                DWORD   level,
'  [out]               LPBYTE  *bufptr,
'  [in]                DWORD   prefmaxlen,
'  [out]               LPDWORD entriesread,
'  [out]               LPDWORD totalentries,
'  [in]                DWORD   servertype,
'  [in, optional]      LMCSTR  domain,
'  [in, out, optional] LPDWORD resume_handle
');
