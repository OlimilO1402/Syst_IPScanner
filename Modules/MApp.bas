Attribute VB_Name = "MApp"
Option Explicit
Private m_Doc As Document
Private Const mExt As String = "ipscan"

Sub Main()
    FrmIPPingScanner.Show
    NewDoc
End Sub

Public Function GetMyIP() As String
    Dim s As String: s = MNew.InternetURL("http://checkip.dyndns.org").Read
    Dim i As Long:   i = InStr(1, s, "IP Address: "): If i = 0 Then Exit Function
    Dim L As Long:   L = InStr(1, s, "</body>"):      If L = 0 Then Exit Function
    i = i + 12:    L = L - i
    GetMyIP = Mid(s, i, L)
End Function

Public Sub NewDoc()
    Set m_Doc = CreateNewDoc
    'Set m_Doc.IPPingScanner = FrmIPPingScanner.IPPingScanner
End Sub
Public Property Get DefaultFileName() As String
    Dim d As Date: d = Now:
    DefaultFileName = App.Path & "\IpScan-" & Year(d) & "-" & Month(d) & "-" & Day(d) & "_" & Hour(d) & "-" & Minute(d) & "-" & Second(d) & "." & mExt
    'Debug.Print DefaultFileName
End Property
Public Property Get Doc() As Document
    Set Doc = m_Doc
End Property

Private Function CreateNewDoc() As Document
    Dim ib As IPAddressV4: Set ib = MNew.IPAddressV4("192.168.178")
    Set CreateNewDoc = MNew.Document(ib, ib.Clone, 50, FrmIPPingScanner.IPPingScanner)
End Function
Public Property Get FileName() As String
    FileName = m_Doc.FileName
    If Len(FileName) = 0 Then
        FileName = DefaultFileName
        'jetzt erst den SaveAs-Dialog öffnen
    End If
End Property
Public Property Let FileName(ByVal Value As String)
    m_Doc.FileName = Value
End Property
'Public Property Get DefaultFileName() As String
'    DefaultFileName = "IPScan" & Now & "."
'End Property

Function DlgFileOpen_Show(aDlg As CommonDialog) As VbMsgBoxResult
Try: On Error GoTo Catch
    With aDlg
        .Filter = "ipscan-files [*." & mExt & "]|*." & mExt
        .FilterIndex = 0
        .DefaultExt = "*.ipscan"
        .FileName = MApp.FileName
        .InitDir = App.Path
        .ShowOpen
    End With
Catch:
    DlgFileOpen_Show = IIf(Err.Number = MSComDlg.ErrorConstants.cdlCancel, VbMsgBoxResult.vbCancel, VbMsgBoxResult.vbOK)
End Function

Function DlgFileSave_Show(aDlg As CommonDialog) As VbMsgBoxResult
Try: On Error GoTo Catch
    With aDlg
        .Filter = "ipscan-files [*.ipscan]|*.ipscan"
        .FilterIndex = 0
        .DefaultExt = "*.ipscan"
        .FileName = MApp.FileName
        '.InitDir = App.Path
        .ShowSave
    End With
Catch:
    DlgFileSave_Show = IIf(Err.Number = MSComDlg.ErrorConstants.cdlCancel, VbMsgBoxResult.vbCancel, VbMsgBoxResult.vbOK)
End Function


Public Sub FileOpen(Optional ByVal aFNm As String = "")
Try: On Error GoTo Finally
    Dim FNm As String:  FNm = IIf(Len(aFNm), aFNm, MApp.FileName)
    Dim FNr As Integer: FNr = FreeFile
    Open FNm For Binary Access Read As FNr
    
    Dim s As String
    
    'read the Base IP
    s = BinaryReadString(FNr): Set m_Doc.IPBase = MNew.IPAddressV4(s)
    
    'read the last IP
    s = BinaryReadString(FNr): Set m_Doc.LastIP = MNew.IPAddressV4(s)
    
    'read how much IPs to search in one step
    s = BinaryReadString(FNr): m_Doc.SearchNIPs = CLng(s)
    
    s = BinaryReadString(FNr): m_Doc.StartIPb4 = CLng(s)
    
    Dim c As Long
    s = BinaryReadString(FNr): c = CLng(s)
    Dim i As Long
    Dim ip As IPAddressV4
    'If c > 0 Then Set IPAddresses = New IPAddresses
    For i = 0 To c - 1
        s = BinaryReadString(FNr)
        Set ip = New IPAddressV4
        ip.ReadFromStr s
        m_Doc.IPAddresses_Add ip
    Next
    
Finally:
    Close FNr
    If Err Then
        MsgBox Err.Description
    End If
End Sub

Public Sub FileSave(Optional ByVal aFNm As String = "")
Try: On Error GoTo Finally
    Dim FNm As String:  FNm = IIf(Len(aFNm), aFNm, MApp.FileName)
    Dim FNr As Integer: FNr = FreeFile
    DelFile FNm
    Open FNm For Binary Access Write As FNr
    BinaryWriteString FNr, m_Doc.IPBase.IPToStr
    BinaryWriteString FNr, m_Doc.LastIP.IPToStr
    BinaryWriteString FNr, CStr(m_Doc.SearchNIPs)
    BinaryWriteString FNr, CStr(m_Doc.StartIPb4)
    BinaryWriteString FNr, CStr(m_Doc.IPAddresses.Count)
    Dim i As Long
    Dim ip As IPAddressV4
    Dim s As String
    For i = 1 To m_Doc.IPAddresses.Count '- 1
        Set ip = m_Doc.IPAddresses.ItemI(i)
        s = ip.WriteToStr
        BinaryWriteString FNr, s
    Next
Finally:
    Close FNr
    If Err Then
        MsgBox Err.Description
    End If
End Sub
Private Sub BinaryWriteString(aFNr As Integer, s As String)
    Dim L As Long: L = Len(s)
    Put aFNr, , L
    Put aFNr, , s
End Sub
Private Function BinaryReadString(aFNr As Integer) As String
    Dim L As Long, s As String
    Get aFNr, , L
    s = Space(L)
    Get aFNr, , s
    BinaryReadString = s
End Function

Private Sub DelFile(aPFN As String)
    On Error Resume Next
    Kill aPFN
    On Error GoTo 0
End Sub
