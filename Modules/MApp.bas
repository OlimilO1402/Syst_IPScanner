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
    
Finally:
    Close FNr
    
End Sub

Public Sub OnFileSave()
    Dim FNm As String:  FNm = MNew.GetSaveFileName
    Dim FNr As Integer: FNr = FreeFile
Try: On Error GoTo Finally
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


