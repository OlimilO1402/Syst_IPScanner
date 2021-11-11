Attribute VB_Name = "MDns"
Option Explicit
#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Private Const DNS_TYPE_PTR          As Long = &HC
Private Const DNS_QUERY_STANDARD    As Long = &H0
Private Const DnsFreeRecordListDeep As Long = &H1&

Private Enum DNS_STATUS
    ERROR_BAD_IP_FORMAT = -3&
    ERROR_NO_PTR_RETURNED = -2&
    ERROR_NO_RR_RETURNED = -1&
    DNS_STATUS_SUCCESS = 0&
End Enum
'typedef struct _DnsRecordW {
'  struct _DnsRecordW *pNext;
'  PWSTR              pName;
'  WORD               wType;
'  WORD               wDataLength;
'  union {
'    DWORD            DW;
'    DNS_RECORD_FLAGS S;
'  } Flags;
'  DWORD              dwTtl;
'  DWORD              dwReserved;
'  union {
'    DNS_A_DATA          A;
'    DNS_SOA_DATAW       SOA;
'    DNS_SOA_DATAW       Soa;
'    DNS_PTR_DATAW       PTR;
'    DNS_PTR_DATAW       Ptr;
'    DNS_PTR_DATAW       NS;
'    DNS_PTR_DATAW       Ns;
'    DNS_PTR_DATAW       CNAME;
'    DNS_PTR_DATAW       Cname;
'    DNS_PTR_DATAW       DNAME;
'    DNS_PTR_DATAW       Dname;
'    DNS_PTR_DATAW       MB;
'    DNS_PTR_DATAW       Mb;
'    DNS_PTR_DATAW       MD;
'    DNS_PTR_DATAW       Md;
'    DNS_PTR_DATAW       MF;
'    DNS_PTR_DATAW       Mf;
'    DNS_PTR_DATAW       MG;
'    DNS_PTR_DATAW       Mg;
'    DNS_PTR_DATAW       MR;
'    DNS_PTR_DATAW       Mr;
'    DNS_MINFO_DATAW     MINFO;
'    DNS_MINFO_DATAW     Minfo;
'    DNS_MINFO_DATAW     RP;
'    DNS_MINFO_DATAW     Rp;
'    DNS_MX_DATAW        MX;
'    DNS_MX_DATAW        Mx;
'    DNS_MX_DATAW        AFSDB;
'    DNS_MX_DATAW        Afsdb;
'    DNS_MX_DATAW        RT;
'    DNS_MX_DATAW        Rt;
'    DNS_TXT_DATAW       HINFO;
'    DNS_TXT_DATAW       Hinfo;
'    DNS_TXT_DATAW       ISDN;
'    DNS_TXT_DATAW       Isdn;
'    DNS_TXT_DATAW       TXT;
'    DNS_TXT_DATAW       Txt;
'    DNS_TXT_DATAW       X25;
'    DNS_NULL_DATA       Null;
'    DNS_WKS_DATA        WKS;
'    DNS_WKS_DATA        Wks;
'    DNS_AAAA_DATA       AAAA;
'    DNS_KEY_DATA        KEY;
'    DNS_KEY_DATA        Key;
'    DNS_SIG_DATAW       SIG;
'    DNS_SIG_DATAW       Sig;
'    DNS_ATMA_DATA       ATMA;
'    DNS_ATMA_DATA       Atma;
'    DNS_NXT_DATAW       NXT;
'    DNS_NXT_DATAW       Nxt;
'    DNS_SRV_DATAW       SRV;
'    DNS_SRV_DATAW       Srv;
'    DNS_NAPTR_DATAW     NAPTR;
'    DNS_NAPTR_DATAW     Naptr;
'    DNS_OPT_DATA        OPT;
'    DNS_OPT_DATA        Opt;
'    DNS_DS_DATA         DS;
'    DNS_DS_DATA         Ds;
'    DNS_RRSIG_DATAW     RRSIG;
'    DNS_RRSIG_DATAW     Rrsig;
'    DNS_NSEC_DATAW      NSEC;
'    DNS_NSEC_DATAW      Nsec;
'    DNS_DNSKEY_DATA     DNSKEY;
'    DNS_DNSKEY_DATA     Dnskey;
'    DNS_TKEY_DATAW      TKEY;
'    DNS_TKEY_DATAW      Tkey;
'    DNS_TSIG_DATAW      TSIG;
'    DNS_TSIG_DATAW      Tsig;
'    DNS_WINS_DATA       WINS;
'    DNS_WINS_DATA       Wins;
'    DNS_WINSR_DATAW     WINSR;
'    DNS_WINSR_DATAW     WinsR;
'    DNS_WINSR_DATAW     NBSTAT;
'    DNS_WINSR_DATAW     Nbstat;
'    DNS_DHCID_DATA      DHCID;
'    DNS_NSEC3_DATA      NSEC3;
'    DNS_NSEC3_DATA      Nsec3;
'    DNS_NSEC3PARAM_DATA NSEC3PARAM;
'    DNS_NSEC3PARAM_DATA Nsec3Param;
'    DNS_TLSA_DATA       TLSA;
'    DNS_TLSA_DATA       Tlsa;
'    DNS_UNKNOWN_DATA    UNKNOWN;
'    DNS_UNKNOWN_DATA    Unknown;
'    PBYTE               pDataPtr;
'  } Data;
'} DNS_RECORDW, *PDNS_RECORDW;
Private Type VBDnsRecord
    pNext       As LongPtr
    pName       As LongPtr
    wType       As Integer
    wDataLength As Integer
    flags       As Long
    dwTTL       As Long
    dwReserved  As Long
    DataPtr     As LongPtr
End Type
'Private Type VBDnsRecord
'    pNext       As Long
'    pName       As Long
'    wType       As Integer
'    wDataLength As Integer
'    flags       As Long
'    dwTTL       As Long
'    dwReserved  As Long
'    DataPtr     As Long
'    'others(9)   As Long
'End Type

Private Declare Function DnsQuery Lib "dnsapi" Alias "DnsQuery_A" ( _
    ByVal Name As String, _
    ByVal wType As Integer, _
    ByVal Options As Long, _
    ByRef aipServers As Any, _
    ByRef ppQueryResultsSet As Long, _
    ByVal pReserved As Long) As Long

Private Declare Function DnsRecordListFree Lib "dnsapi" ( _
    ByVal pDnsRecord As LongPtr, _
    ByVal DnsFreeRecordListDeep As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
    ByRef pDst As Any, _
    ByRef pSrc As Any, _
    ByVal BytLen As Long)

Private Declare Function StrCopyA Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal retval As String, _
    ByVal PTR As Long) As Long

Private Declare Function StrLenA Lib "kernel32" Alias "lstrlenA" ( _
    ByVal PTR As LongPtr) As Long
    
''https://docs.microsoft.com/en-us/windows/win32/api/windns/nf-windns-dnsquery_w
''DNS_STATUS DnsQuery_W(
''  [in]                PCWSTR      pszName,
''  [in]                WORD        wType,
''  [in]                DWORD       Options,
''  [in, out, optional] PVOID       pExtra,
''  [out, optional]     PDNS_RECORD *ppQueryResults,
''  [out, optional]     PVOID * pReserved
'');
Private Declare Function DnsQuery_W Lib "dnsapi" (ByVal pName As LongPtr, ByVal wType As Integer, ByVal Options As Long, ByRef aipServers As Any, ByRef ppQueryResultsSet As Any, ByVal pReserved As Long) As DNS_STATUS
'
'Private Declare Function DnsRecordListFree Lib "dnsapi" (ByVal pDnsRecord As LongPtr, ByVal DnsFreeRecordListDeep As Long) As Long
'
'Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal bytlen As Long)
'
'Private Declare Function lstrcpyW Lib "kernel32" (ByVal retval As LongPtr, ByVal PTR As LongPtr) As Long
'
'Private Declare Function lstrlenW Lib "kernel32" (ByVal PTR As LongPtr) As Long
'
'https://stackoverflow.com/questions/5139511/vb6-how-to-get-the-remote-computer-name-based-on-the-given-ip-address
'
'Public Function IP2HostName(ByVal ip As String, ByRef HostName As String) As Long
'    Dim Octets() As String
'    Dim OctX As Long
'    Dim NumPart As Long
'    Dim BadIP As Boolean
'    Dim lngDNSRec As Long
'    Dim Record As VBDnsRecord
'    Dim Length As Long
'    'Returns DNS_STATUS Enum values, otherwise a DNS system error code.
'
'    ip = Trim$(ip)
'    If Len(ip) = 0 Then IP2HostName = ERROR_BAD_IP_FORMAT: Exit Function
'    Octets = Split(ip, ".")
'    If UBound(Octets) <> 3 Then IP2HostName = ERROR_BAD_IP_FORMAT: Exit Function
'    For OctX = 0 To 3
'        If IsNumeric(Octets(OctX)) Then
'            NumPart = CInt(Octets(OctX))
'            If 0 <= NumPart And NumPart <= 255 Then
'                Octets(OctX) = CStr(NumPart)
'            Else
'                BadIP = True
'                Exit For
'            End If
'        Else
'            BadIP = True
'            Exit For
'        End If
'    Next
'    If BadIP Then IP2HostName = ERROR_BAD_IP_FORMAT: Exit Function
'
'    ip = Octets(3) & "." & Octets(2) & "." & Octets(1) & "." & Octets(0) & ".IN-ADDR.ARPA"
'
'    IP2HostName = DnsQuery_W(ip, DNS_TYPE_PTR, DNS_QUERY_STANDARD, ByVal 0, lngDNSRec, 0)
'    If IP2HostName = DNS_STATUS_SUCCESS Then
'        If lngDNSRec <> 0 Then
'            RtlMoveMemory Record, ByVal lngDNSRec, LenB(Record)
'
'            With Record
'                If .wType = DNS_TYPE_PTR Then
'                    Length = lstrlenW(.prt)
'                    HostName = String$(Length, 0)
'                    lstrcpyW HostName, .prt
'                Else
'                    IP2HostName = ERROR_NO_PTR_RETURNED
'                End If
'            End With
'            DnsRecordListFree lngDNSRec, DnsFreeRecordListDeep
'        Else
'            IP2HostName = ERROR_NO_RR_RETURNED
'        End If
'    'Else
'        'Return with DNS error code.
'    End If
'End Function
'

Public Function IP2HostName(ByVal IP As String, ByRef HostName As String) As Long
    Dim Octets() As String
    Dim OctX As Long
    Dim NumPart As Long
    Dim BadIP As Boolean
    Dim lngDNSRec As Long
    Dim Record As VBDnsRecord
    Dim Length As Long
    'Returns DNS_STATUS Enum values, otherwise a DNS system error code.

    IP = Trim$(IP)
    If Len(IP) = 0 Then IP2HostName = ERROR_BAD_IP_FORMAT: Exit Function
    Octets = Split(IP, ".")
    If UBound(Octets) <> 3 Then IP2HostName = ERROR_BAD_IP_FORMAT: Exit Function
    For OctX = 0 To 3
        If IsNumeric(Octets(OctX)) Then
            NumPart = CInt(Octets(OctX))
            If 0 <= NumPart And NumPart <= 255 Then
                Octets(OctX) = CStr(NumPart)
            Else
                BadIP = True
                Exit For
            End If
        Else
            BadIP = True
            Exit For
        End If
    Next
    If BadIP Then IP2HostName = ERROR_BAD_IP_FORMAT: Exit Function
    
    'what the heck is IN-ADDR.ARPA
    IP = Octets(3) & "." & Octets(2) & "." & Octets(1) & "." & Octets(0) & ".IN-ADDR.ARPA"

    IP2HostName = DnsQuery(IP, DNS_TYPE_PTR, DNS_QUERY_STANDARD, ByVal 0, lngDNSRec, 0)
    If IP2HostName = DNS_STATUS_SUCCESS Then
        If lngDNSRec <> 0 Then
            RtlMoveMemory Record, ByVal lngDNSRec, LenB(Record)

            With Record
                If .wType = DNS_TYPE_PTR Then
                    Length = StrLenA(.DataPtr)
                    HostName = String$(Length, 0)
                    StrCopyA HostName, .DataPtr
                Else
                    IP2HostName = ERROR_NO_PTR_RETURNED
                End If
            End With
            DnsRecordListFree lngDNSRec, DnsFreeRecordListDeep
        Else
            IP2HostName = ERROR_NO_RR_RETURNED
        End If
    'Else
        'Return with DNS error code.
    End If
End Function
