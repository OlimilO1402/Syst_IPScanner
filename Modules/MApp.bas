Attribute VB_Name = "MApp"
Option Explicit
Private m_Doc As Document
Private Const mExt As String = "ipscan"

Sub Main()
    FrmIPPingScanner.Show
    NewDoc
End Sub
Public Sub NewDoc()
    Set m_Doc = CreateNewDoc
    'Set m_Doc.IPPingScanner = FrmIPPingScanner.IPPingScanner
End Sub

Public Property Get Doc() As Document
    Set Doc = m_Doc
End Property

Private Function CreateNewDoc() As Document
    Dim ib As IPAddressV4: Set ib = MNew.IPAddressV4("192.168.178")
    Set CreateNewDoc = MNew.Document(ib, ib.Clone, 50, FrmIPPingScanner.IPPingScanner)
End Function
Public Property Get FileName() As String
    FileName = m_Doc.FileName
End Property
Public Property Let FileName(ByVal Value As String)
    m_Doc.FileName = Value
End Property
Public Property Get DefaultFileName() As String
    DefaultFileName = "IPScan" & Now & "."
End Property

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
    Dim FNm As String:  FNm = IIf(Len(aFNm), MApp.FileName, aFNm)
    Dim FNr As Integer: FNr = FreeFile
    Open FNm For Binary Access Read As FNr
    Dim s As String: s = Space(LOF(FNr))
    Get FNr, , s
    Dim sLines() As String: sLines = Split(s, vbCrLf)
    Dim l As Long, i As Long, ul As Long: ul = UBound(sLines)
    
    
    
    'read the Base IP
               If l <= ul Then Set m_Doc.IPBase = MNew.IPAddressV4(sLines(l))
    
    'read the last IP
    l = l + 1: If l <= ul Then Set m_Doc.LastIP = MNew.IPAddressV4(sLines(l))
    
    'read how much IPs to search in one step
    l = l + 1: If l <= ul Then m_Doc.SearchNIPs = CLng(sLines(l))
    
    l = l + 1: If l <= ul Then m_Doc.StartIPb4 = CLng(sLines(l))
    Dim c As Long
    l = l + 1: If l <= ul Then c = CLng(sLines(l))
    
    'If c > 0 Then Set IPAddresses = New IPAddresses
    For i = 0 To c - 1
        m_Doc.IPAddresses_Add MNew.IPAddressV4(sLines(l))
        'IPAddresses.Add MNew.IPAddressV4(sLines(l))
    Next
    i = i + 1: If i <= ul Then Set m_Doc.LastIP = MNew.IPAddressV4(sLines(i))
    i = i + 1: If i <= ul Then Set m_Doc.LastIP = MNew.IPAddressV4(sLines(i))
    i = i + 1: If i <= ul Then Set m_Doc.LastIP = MNew.IPAddressV4(sLines(i))
    i = i + 1: If i <= ul Then Set m_Doc.LastIP = MNew.IPAddressV4(sLines(i))
    i = i + 1: If i <= ul Then Set m_Doc.LastIP = MNew.IPAddressV4(sLines(i))
    i = i + 1: If i <= ul Then Set m_Doc.LastIP = MNew.IPAddressV4(sLines(i))
    i = i + 1: If i <= ul Then Set m_Doc.LastIP = MNew.IPAddressV4(sLines(i))
    '????????????
    
Finally:
    Close FNr
    
End Sub

Public Sub FileSave(Optional ByVal aFNm As String = "")
Try: On Error GoTo Finally
    Dim FNm As String:  FNm = IIf(Len(aFNm), MApp.FileName, aFNm)
    Dim FNr As Integer: FNr = FreeFile
    DelFile FNm
    Open FNm For Binary Access Write As FNr
    Put FNr, , m_Doc.IPBase.IPToStr & vbCrLf
    Put FNr, , m_Doc.LastIP.IPToStr & vbCrLf
    Put FNr, , CStr(m_Doc.SearchNIPs)
    Put FNr, , CStr(m_Doc.StartIPb4)
    Put FNr, , CStr(m_Doc.IPAddresses.Count)
    Dim i As Long
    For i = 0 To m_Doc.IPAddresses.Count - 1
        Put FNr, , m_Doc.IPAddresses.Item(CStr(i)).WriteToStr
    Next
Finally:
    Close FNr
End Sub

Private Sub DelFile(aPFN As String)
    On Error Resume Next
    Kill aPFN
    On Error GoTo 0
End Sub
