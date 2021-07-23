Attribute VB_Name = "MNew"
Option Explicit

Public Function IPAddressV4(StrLngBytesNewAddress, Optional aName As String) As IPAddressV4
    Set IPAddressV4 = New IPAddressV4: IPAddressV4.New_ StrLngBytesNewAddress, aName
End Function

Public Function Document(aIPBase As IPAddressV4, aLastIP As IPAddressV4, nSearchIPs As Long, aScanner As IPPingScanner) As Document
    Set Document = New Document: Document.New_ aIPBase, aLastIP, nSearchIPs, aScanner
End Function
    
Public Function Splitter(bMDI As Boolean, MyOwner As Object, MyContainer As Object, Name As String, LeftTop As Control, RghtBot As Control) As Splitter
    Set Splitter = New Splitter: Splitter.New_ bMDI, MyOwner, MyContainer, Name, LeftTop, RghtBot
End Function

Public Function InternetURL(sURL As String) As InternetURL
    Set InternetURL = New InternetURL: InternetURL.New_ sURL
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
Public Function MessCommonDlgError(e As MSComDlg.ErrorConstants) As String
    Dim s As String
    Select Case e
    Case cdlDialogFailure:        s = "Dialog Failure"         '= -32768 (&HFFFF8000)
    Case cdlHelp:                 s = "Help"                   '= 32751 (&H7FEF)
    Case cdlAlloc:                s = "Alloc"                  '= 32752 (&H7FF0)
    Case cdlCancel:               s = "Cancel"                 '= 32755 (&H7FF3)
    Case cdlMemLockFailure:       s = "Mem Lock Failure"       '= 32757 (&H7FF5)
    Case cdlMemAllocFailure:      s = "Mem Alloc Failure"      '= 32758 (&H7FF6)
    Case cdlLockResFailure:       s = "Lock Res Failure"       '= 32759 (&H7FF7)
    Case cdlLoadResFailure:       s = "Load Res Failure"       '= 32760 (&H7FF8)
    Case cdlFindResFailure:       s = "Find Res Failure"       '= 32761 (&H7FF9)
    Case cdlLoadStrFailure:       s = "Load Str Failure"       '= 32762 (&H7FFA)
    Case cdlNoInstance:           s = "No Instance"            '= 32763 (&H7FFB)
    Case cdlNoTemplate:           s = "No Template"            '= 32764 (&H7FFC)
    Case cdlInitialization:       s = "Initialization"         '= 32765 (&H7FFD)
    Case cdlInvalidPropertyValue: s = "Invalid Property Value" '= 380 (&H17C)
    Case cdlSetNotSupported:      s = "Set Not Supported"      '= 383 (&H17F)
    Case cdlGetNotSupported:      s = "Get Not Supported"      '= 394 (&H18A)
    Case cdlInvalidSafeModeProcCall: s = "Invalid Safe Mode Proc Call" '= 680 (&H2A8)
    Case cdlBufferTooSmall:       s = "Buffer Too Small"       '= 20476 (&H4FFC)
    Case cdlInvalidFileName:      s = "Invalid FileName"       '= 20477 (&H4FFD)
    Case cdlSubclassFailure:      s = "Subclass Failure"       '= 20478 (&H4FFE)
    Case cdlNoFonts:              s = "No Fonts"               '= 24574 (&H5FFE)
    Case cdlPrinterNotFound:      s = "Printer Not Found"      '= 28660 (&H6FF4)
    Case cdlCreateICFailure:      s = "Create IC Failure"      '= 28661 (&H6FF5)
    Case cdlDndmMismatch:         s = "Dndm Mismatch"          '= 28662 (&H6FF6)
    Case cdlNoDefaultPrn:         s = "No Default Prn"         '= 28663 (&H6FF7)
    Case cdlNoDevices:            s = "No Devices"             '= 28664 (&H6FF8)
    Case cdlInitFailure:          s = "Init Failure"           ' 28665 (&H6FF9)
    Case cdlGetDevModeFail:       s = "Get Dev Mode Fail"      '= 28666 (&H6FFA)
    Case cdlLoadDrvFailure:       s = "Load Drv Failure"       '= 28667 (&H6FFB)
    Case cdlRetDefFailure:        s = "Ret Def Failure"        '= 28668 (&H6FFC)
    Case cdlParseFailure:         s = "Parse Failure"          '= 28669 (&H6FFD)
    End Select
    MessCommonDlgError = s
End Function
