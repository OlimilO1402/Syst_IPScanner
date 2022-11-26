Attribute VB_Name = "MNew"
Option Explicit
'
'Public Function IPAddress(v1, v2, v3, v4, v5, v6, v7, v8) As IPAddress
'    Set IPAddress = New IPAddresses: IPAddress.New_ v1, v2, v3, v4, v5, v6, v7, v8
'End Function
'
'Public Function IPAddressV4(StrLngBytesNewAddress, Optional aName As String) As IPAddressV4
'    Set IPAddressV4 = New IPAddressV4: IPAddressV4.New_ StrLngBytesNewAddress, aName
'End Function
'
'Public Function IPAddressV6(ByVal d1 As Integer, ByVal d2 As Integer, ByVal d3 As Integer, ByVal d4 As Integer, ByVal d5 As Integer, ByVal d6 As Integer, ByVal d7 As Integer, ByVal d8 As Integer) As IPAddressV6
'    Set IPAddressV6 = New IPAddressV6: IPAddressV6.New_ d1, d2, d3, d4, d5, d6, d7, d8
'End Function


Public Function IPAddress(ByVal i1 As Integer, ByVal i2 As Integer, ByVal i3 As Integer, ByVal i4 As Integer, Optional i5, Optional i6, Optional i7, Optional i8) As IPAddress
    Set IPAddress = New IPAddresses: IPAddress.New_ i1, i2, i3, i4, i5, i6, i7, i8
End Function

Public Function IPAddressV(StrLngBytesNewAddress) As IPAddress
    Set IPAddressV = New IPAddress: IPAddressV.New_ StrLngBytesNewAddress
End Function

Public Function IPAddressV4(ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte, ByVal b4 As Byte, Optional Port As Integer) As IPAddress
    Set IPAddressV4 = New IPAddress: IPAddressV4.NewV4 b1, b2, b3, b4, Port
End Function

Public Function IPAddressV6(ByVal i1 As Integer, ByVal i2 As Integer, ByVal i3 As Integer, ByVal i4 As Integer, ByVal i5 As Integer, ByVal i6 As Integer, ByVal i7 As Integer, ByVal i8 As Integer) As IPAddress
    Set IPAddressV6 = New IPAddress: IPAddressV6.NewV6 i1, i2, i3, i4, i5, i6, i7, i8
End Function


Public Function Document(aIPBase As IPAddress, aLastIP As IPAddress, nSearchIPs As Long, aScanner As IPPingScanner) As Document
    Set Document = New Document: Document.New_ aIPBase, aLastIP, nSearchIPs, aScanner
End Function
    
Public Function Splitter(bMDI As Boolean, MyOwner As Object, MyContainer As Object, Name As String, LeftTop As Control, RghtBot As Control) As Splitter
    Set Splitter = New Splitter: Splitter.New_ bMDI, MyOwner, MyContainer, Name, LeftTop, RghtBot
End Function

Public Function InternetURL(sURL As String) As InternetURL
    Set InternetURL = New InternetURL: InternetURL.New_ sURL
End Function

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                     Optional ByVal aFileName As String, _
                     Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

Public Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function
    
'Public Function GetOpenFileName() As String
'Try: On Error GoTo Catch
'    With FrmIPPingScanner.SaveFileDialog
'        .CancelError = True
'        .Filter = "ipscan-files [*.ipscan]|*.ipscan"
'        .FilterIndex = 0
'        .DefaultExt = "*.ipscan"
'        .FileName = "IPScan-" & Now
'        .InitDir = App.Path
'    'FrmIPPingScanner.SaveFileDialog.ShowColor
'    'FrmIPPingScanner.SaveFileDialog.ShowFont
'    'FrmIPPingScanner.SaveFileDialog.ShowHelp
'    'FrmIPPingScanner.SaveFileDialog.ShowOpen
'    'FrmIPPingScanner.SaveFileDialog.ShowPrinter
'        .ShowOpen
'    End With
'    Exit Function
'Catch:
'    If Err.Number <> MSComDlg.ErrorConstants.cdlCancel Then
'        MsgBox MessCommonDlgError(Err.Number)
'    End If
'End Function
'
'Public Function GetSaveFileName() As String
'Try: On Error GoTo Catch
'    With FrmIPPingScanner.SaveFileDialog
'        .CancelError = True
'        .Filter = "ipscan-files [*.ipscan]|*.ipscan"
'        .FilterIndex = 0
'        .DefaultExt = "*.ipscan"
'        .FontName = "IPScan-" & Now
'        .InitDir = App.Path
'    'FrmIPPingScanner.SaveFileDialog.ShowColor
'    'FrmIPPingScanner.SaveFileDialog.ShowFont
'    'FrmIPPingScanner.SaveFileDialog.ShowHelp
'    'FrmIPPingScanner.SaveFileDialog.ShowOpen
'    'FrmIPPingScanner.SaveFileDialog.ShowPrinter
'        .ShowSave
'    End With
'    Exit Function
'Catch:
'    If Err.Number <> MSComDlg.ErrorConstants.cdlCancel Then
'        MsgBox MessCommonDlgError(Err.Number)
'    End If
'End Function
'
'Public Function MessCommonDlgError(e As MSComDlg.ErrorConstants) As String
Public Function MessCommonDlgError(e As Long) As String
    Dim s As String
    Select Case e
    Case 0: s = "What error?"
'    Case MSComDlg.ErrorConstants.cdlDialogFailure:        s = "Dialog Failure"         '= -32768 (&HFFFF8000)
'    Case MSComDlg.ErrorConstants.cdlHelp:                 s = "Help"                   '= 32751 (&H7FEF)
'    Case MSComDlg.ErrorConstants.cdlAlloc:                s = "Alloc"                  '= 32752 (&H7FF0)
'    Case MSComDlg.ErrorConstants.cdlCancel:               s = "Cancel"                 '= 32755 (&H7FF3)
'    Case MSComDlg.ErrorConstants.cdlMemLockFailure:       s = "Mem Lock Failure"       '= 32757 (&H7FF5)
'    Case MSComDlg.ErrorConstants.cdlMemAllocFailure:      s = "Mem Alloc Failure"      '= 32758 (&H7FF6)
'    Case MSComDlg.ErrorConstants.cdlLockResFailure:       s = "Lock Res Failure"       '= 32759 (&H7FF7)
'    Case MSComDlg.ErrorConstants.cdlLoadResFailure:       s = "Load Res Failure"       '= 32760 (&H7FF8)
'    Case MSComDlg.ErrorConstants.cdlFindResFailure:       s = "Find Res Failure"       '= 32761 (&H7FF9)
'    Case MSComDlg.ErrorConstants.cdlLoadStrFailure:       s = "Load Str Failure"       '= 32762 (&H7FFA)
'    Case MSComDlg.ErrorConstants.cdlNoInstance:           s = "No Instance"            '= 32763 (&H7FFB)
'    Case MSComDlg.ErrorConstants.cdlNoTemplate:           s = "No Template"            '= 32764 (&H7FFC)
'    Case MSComDlg.ErrorConstants.cdlInitialization:       s = "Initialization"         '= 32765 (&H7FFD)
'    Case MSComDlg.ErrorConstants.cdlInvalidPropertyValue: s = "Invalid Property Value" '= 380 (&H17C)
'    Case MSComDlg.ErrorConstants.cdlSetNotSupported:      s = "Set Not Supported"      '= 383 (&H17F)
'    Case MSComDlg.ErrorConstants.cdlGetNotSupported:      s = "Get Not Supported"      '= 394 (&H18A)
'    Case MSComDlg.ErrorConstants.cdlInvalidSafeModeProcCall: s = "Invalid Safe Mode Proc Call" '= 680 (&H2A8)
'    Case MSComDlg.ErrorConstants.cdlBufferTooSmall:       s = "Buffer Too Small"       '= 20476 (&H4FFC)
'    Case MSComDlg.ErrorConstants.cdlInvalidFileName:      s = "Invalid FileName"       '= 20477 (&H4FFD)
'    Case MSComDlg.ErrorConstants.cdlSubclassFailure:      s = "Subclass Failure"       '= 20478 (&H4FFE)
'    Case MSComDlg.ErrorConstants.cdlNoFonts:              s = "No Fonts"               '= 24574 (&H5FFE)
'    Case MSComDlg.ErrorConstants.cdlPrinterNotFound:      s = "Printer Not Found"      '= 28660 (&H6FF4)
'    Case MSComDlg.ErrorConstants.cdlCreateICFailure:      s = "Create IC Failure"      '= 28661 (&H6FF5)
'    Case MSComDlg.ErrorConstants.cdlDndmMismatch:         s = "Dndm Mismatch"          '= 28662 (&H6FF6)
'    Case MSComDlg.ErrorConstants.cdlNoDefaultPrn:         s = "No Default Prn"         '= 28663 (&H6FF7)
'    Case MSComDlg.ErrorConstants.cdlNoDevices:            s = "No Devices"             '= 28664 (&H6FF8)
'    Case MSComDlg.ErrorConstants.cdlInitFailure:          s = "Init Failure"           ' 28665 (&H6FF9)
'    Case MSComDlg.ErrorConstants.cdlGetDevModeFail:       s = "Get Dev Mode Fail"      '= 28666 (&H6FFA)
'    Case MSComDlg.ErrorConstants.cdlLoadDrvFailure:       s = "Load Drv Failure"       '= 28667 (&H6FFB)
'    Case MSComDlg.ErrorConstants.cdlRetDefFailure:        s = "Ret Def Failure"        '= 28668 (&H6FFC)
'    Case MSComDlg.ErrorConstants.cdlParseFailure:         s = "Parse Failure"          '= 28669 (&H6FFD)
    End Select
    MessCommonDlgError = s
End Function
