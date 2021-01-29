Attribute VB_Name = "MApp"
Option Explicit
Public IPBase      As IPAddressV4
Public LastIP      As IPAddressV4
Public SearchNIPs  As Long
Public StartIPb4   As Long
Public IPAddresses As IPAddresses

Sub Main()
    Set IPBase = MNew.IPAddressV4("192.168.178")
    Set LastIP = IPBase.Clone
    Set IPAddresses = New IPAddresses
    SearchNIPs = 50
    FrmIPPingScanner.Show
End Sub

Public Sub OnFileOpen()
    Dim FNm As String:  FNm = MNew.GetOpenFileName
    Dim FNr As Integer: FNr = FreeFile
Try: On Error GoTo Finally
    Open FNm For Binary Access Write As FNr
    Dim s As String: s = Space(LOF(FNr))
    Get FNr, , s
    Dim sLines() As String: sLines = Split(s, vbCrLf)
    Dim l As Long, i As Long, u As Long: ul = UBound(sLines)
    
               If l <= ul Then Set IPBase = MNew.IPAddressV4(sLines(l))
    l = l + 1: If l <= ul Then Set LastIP = MNew.IPAddressV4(sLines(l))
    l = l + 1: If l <= ul Then SearchNIPs = CLng(sLines(l))
    l = l + 1: If l <= ul Then StartIPb4 = CLng(sLines(l))
    Dim c As Long
    l = l + 1: If l <= ul Then c = CLng(sLines(l))
    If c > 0 Then Set IPAddresses = New IPAddresses
    For i = 0 To c - 1
        IPAddresses.Add MNew.IPAddressV4(sLines(l))
    Next
    i = i + 1: If i <= u Then Set LastIP = MNew.IPAddressV4(sLines(i))
    i = i + 1: If i <= u Then Set LastIP = MNew.IPAddressV4(sLines(i))
    i = i + 1: If i <= u Then Set LastIP = MNew.IPAddressV4(sLines(i))
    i = i + 1: If i <= u Then Set LastIP = MNew.IPAddressV4(sLines(i))
    i = i + 1: If i <= u Then Set LastIP = MNew.IPAddressV4(sLines(i))
    i = i + 1: If i <= u Then Set LastIP = MNew.IPAddressV4(sLines(i))
    i = i + 1: If i <= u Then Set LastIP = MNew.IPAddressV4(sLines(i))
    
Finally:
    Close FNr
    
End Sub

Public Sub OnFileSave()
    Dim FNm As String:  FNm = MNew.GetSaveFileName
    Dim FNr As Integer: FNr = FreeFile
Try: On Error GoTo Finally
    DelFile FNm
    Open FNm For Binary Access Write As FNr
    Put FNr, , IPBase.IPToStr & vbCrLf
    Put FNr, , LastIP.IPToStr & vbCrLf
    Put FNr, , CStr(SearchNIPs)
    Put FNr, , CStr(StartIPb4)
    Put FNr, , CStr(IPAddresses.Count)
    Dim i As Long
    For i = 0 To IPAddresses.Count - 1
        Put FNr, , IPAddresses.Item(i).WriteToStr
    Next
Finally:
    Close FNr
End Sub

Private Sub DelFile(aPFN As String)
    On Error Resume Next
    Kill aPFN
    On Error GoTo 0
End Sub
