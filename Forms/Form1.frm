VERSION 5.00
Begin VB.Form FrmIPPingScanner 
   Caption         =   "IPScanner"
   ClientHeight    =   6375
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   12495
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   10800
      TabIndex        =   11
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Panel1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   12345
      TabIndex        =   6
      Top             =   480
      Width           =   12375
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5235
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2295
      End
      Begin VB.PictureBox Panel2 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   2400
         ScaleHeight     =   5505
         ScaleWidth      =   9705
         TabIndex        =   7
         Top             =   0
         Width           =   9735
         Begin VB.TextBox TxtIPInfo 
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5295
            Left            =   2400
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Beides
            TabIndex        =   10
            Top             =   0
            Width           =   7095
         End
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5235
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2295
         End
      End
   End
   Begin VB.CommandButton BtnScanNextXXIPs 
      Caption         =   "Scan next 50 IPs"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.TextBox TxtIPBase 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4800
      TabIndex        =   0
      Text            =   "192.168.178"
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label LblDTime 
      AutoSize        =   -1  'True
      Caption         =   "            "
      Height          =   315
      Left            =   9000
      TabIndex        =   5
      Top             =   0
      Width           =   1620
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Valid IP-addresses:"
      Height          =   195
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Invalid IP-addresses:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1470
   End
   Begin VB.Label Label2 
      Caption         =   "IP-Base:"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuopt 
      Caption         =   "&Option"
      Begin VB.Menu mnuOptOptionN 
         Caption         =   "Option N"
      End
      Begin VB.Menu mnuOptDllOrNslookup 
         Caption         =   "Dll or Nslookup"
         Begin VB.Menu mnuOptDllOrNslookupExe 
            Caption         =   "Nslookup.exe"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptDllOrNslookupDll 
            Caption         =   "Dns.dll"
         End
      End
      Begin VB.Menu mnuOptionsUserPCName 
         Caption         =   "Get User+Computername"
      End
      Begin VB.Menu mnuOptionNetViewDomain 
         Caption         =   "net view /domain:"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " &? "
      Begin VB.Menu mnuHelpExternalIP 
         Caption         =   "My external IP"
      End
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "Info"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopAddPort 
         Caption         =   "Add port"
      End
      Begin VB.Menu mnuPopEditPort 
         Caption         =   "Edit port"
      End
      Begin VB.Menu mnuOptPingIP 
         Caption         =   "Ping IP"
      End
      Begin VB.Menu mnuOptPingIPport 
         Caption         =   "Ping IP:port"
      End
      Begin VB.Menu mnuOptNslookupIP 
         Caption         =   "Nslookup IP"
      End
      Begin VB.Menu mnuOptNslookupIPport 
         Caption         =   "Nslookup IP:port"
      End
   End
End
Attribute VB_Name = "FrmIPPingScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents IPPingScanner As IPPingScanner
Attribute IPPingScanner.VB_VarHelpID = -1
Private m_CancelFlag As VbMsgBoxResult
Private WithEvents Splitter1 As Splitter
Attribute Splitter1.VB_VarHelpID = -1
Private Splitter2 As Splitter
Private mnuPopListBox As ListBox

Private Sub Command1_Click()
    Dim ipv6 As IPAddressV6: Set ipv6 = MNew.IPAddressV6(1, 2, 3, 4, 5, 6, 7, 8)
    MsgBox ipv6.ToStr
End Sub

'Private Sub Command1_Click()
'    Dim sIP As String: sIP = "192.168.178.20"
'    Dim sName As String
'    Dim rv As Long
'    rv = MDns.IP2HostName(sIP, sName)
'    MsgBox sName
'End Sub

Private Sub Form_Load()
    Set IPPingScanner = New IPPingScanner
    'TxtIPBase.Text = MApp.IPBase.IPToStr
    Panel1.BorderStyle = BorderStyleConstants.vbTransparent
    Panel2.BorderStyle = 0
    
    Set Splitter1 = MNew.Splitter(False, Me, Panel1, "Splitter1", List1, Panel2)
    Splitter1.BorderStyle = bsXPStyl
    Splitter1.LeftTopPos = List1.Width
    
    Set Splitter2 = MNew.Splitter(False, Me, Panel2, "Splitter2", List2, TxtIPInfo)
    Splitter2.BorderStyle = bsXPStyl
    Splitter2.LeftTopPos = List2.Width
    mnuPopup.Visible = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        PopupMenu mnuopt
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MApp.Doc.IsDataChanged Then
        Dim mbr As VbMsgBoxResult: mbr = MsgBox("Data changed, save it?", vbYesNoCancel)
        Select Case mbr
        Case vbYes:  mnuFileSave_Click
                     'Cancel = True
        Case vbNo:   'just quit
        Case vbCancel: Cancel = True
        End Select
    End If
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single

'    L = List1.Left: T = List1.Top: W = List1.Width: H = Me.ScaleHeight - T
'    If W > 0 And H > 0 Then List1.Move L, T, W, H
'    L = List2.Left: T = List2.Top: W = List2.Width ': H = Me.ScaleHeight - T
'    If W > 0 And H > 0 Then List2.Move L, T, W, H
'    L = TxtIPInfo.Left: T = TxtIPInfo.Top: W = Me.ScaleWidth - L
'    If W > 0 And H > 0 Then TxtIPInfo.Move L, T, W, H
    Dim brdr As Single: 'brdr = 8 * Screen.TwipsPerPixelX
    L = 0: T = Panel1.Top
    W = Me.ScaleWidth - L
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Panel1.Move L, T, W, H
    
End Sub

Public Sub UpdateView()
    Dim Doc As Document: Set Doc = MApp.Doc
    Me.TxtIPBase.Text = Doc.IPBase.ToStr
    Dim IP As IpAddress
    Dim i As Long
    For i = 1 To Doc.IPAddresses.Count
        Set IP = Doc.IPAddresses.ItemI(i)
        If IP.IsValid Then
            List2.AddItem IP.IPToStr
        Else
            List1.AddItem IP.IPToStr
        End If
    Next
    BtnScanNextXXIPs.Caption = "Scan next " & MApp.Doc.SearchNIPs & " IPs"
End Sub

' v ############################## v '  Menu mnuFile  ' v ############################## v '
Private Sub mnuFileNew_Click()
    'MApp.Doc.IPAddresses.Clear
    
    MApp.NewDoc
    
    TxtIPInfo.Text = ""
    List1.Clear
    List2.Clear
    
    'MApp.Doc.StartIPb4 = 0
    'StartIPb4 = 0
    'SearchNIPs = 50
End Sub

Private Sub mnuFileOpen_Click()
    'Me.FileDialog.FileName = "IPScan" & Now
    Dim OFD As OpenFileDialog: Set OFD = New OpenFileDialog
    If MApp.DlgFileOpen_Show(OFD) = vbCancel Then Exit Sub
    MApp.FileOpen OFD.FileName
    MApp.Doc.IsDataChanged = False
    UpdateView 'MApp.Doc
End Sub

Private Sub mnuFileSave_Click()
    MApp.FileSave
    'TODO: we should first ask if filesave was True
    MApp.Doc.IsDataChanged = False
End Sub

Private Sub mnuFileSaveAs_Click()
'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/cc144156(v=vs.85)
    Dim SFD As SaveFileDialog: Set SFD = New SaveFileDialog
    If MApp.DlgFileSave_Show(SFD) = vbCancel Then Exit Sub
    MApp.FileSave SFD.FileName
    'TODO: we should first ask if filesave was True
    MApp.Doc.IsDataChanged = False

End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub
' ^ ############################## ^ '  Menu mnuFile  ' ^ ############################## ^ '

' v ############################## v '  Menu mnuOpt   ' v ############################## v '
Private Sub mnuOptOptionN_Click()
    Dim s As String
    s = InputBox("n: ", "How many ips to search each time?", MApp.Doc.SearchNIPs)
    If s = vbNullString Then Exit Sub
    If Not IsNumeric(s) Then
        mnuOptOptionN_Click
        Exit Sub
    End If
    MApp.Doc.SearchNIPs = CLng(s)
    BtnScanNextXXIPs.Caption = "Scan next " & MApp.Doc.SearchNIPs & " IPs"
End Sub

Private Sub mnuOptDllOrNslookupDll_Click()
    mnuOptDllOrNslookupExe.Checked = False
    mnuOptDllOrNslookupDll.Checked = True
End Sub
Private Sub mnuOptDllOrNslookupExe_Click()
    mnuOptDllOrNslookupExe.Checked = True
    mnuOptDllOrNslookupDll.Checked = False
End Sub

Private Sub mnuPopAddPort_Click()
    Dim i As Long: i = mnuPopListBox.ListIndex
    If i < 0 Then Exit Sub
    Dim s As String: s = mnuPopListBox.List(i)
    Dim aIP As IpAddress
    If MApp.Doc.IPAddresses.Contains(s) Then
        Set aIP = MApp.Doc.IPAddresses.Item(s)
        'aIPV4.CallPing
        TxtIPInfo.Text = aIP.ToInfoStr
    End If
End Sub

Private Sub mnuOptPingIP_Click()
    Dim i As Long: i = mnuPopListBox.ListIndex
    If i < 0 Then Exit Sub
    Dim s As String: s = mnuPopListBox.List(i)
    Dim aIPV4 As IpAddress
    If MApp.Doc.IPAddresses.Contains(s) Then
        Set aIPV4 = MApp.Doc.IPAddresses.Item(s)
        aIPV4.CallPing
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub
Private Sub mnuOptPingIPport_Click()
    Dim i As Long: i = mnuPopListBox.ListIndex
    If i < 0 Then Exit Sub
    Dim s As String: s = mnuPopListBox.List(i)
    Dim aIPV4 As IpAddress
    If MApp.Doc.IPAddresses.Contains(s) Then
        Set aIPV4 = MApp.Doc.IPAddresses.Item(s)
        Set aIPV4 = aIPV4.Clone
        Dim sp As String: sp = InputBox("Use port-number for IP: " & aIPV4.ToStr, "Port")
        If StrPtr(sp) = 0 Then Exit Sub
        aIPV4.Port = CLng(sp)
        aIPV4.CallPing
        MApp.Doc.IPAddresses_Add aIPV4
        mnuPopListBox.AddItem aIPV4.ToStr
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub
Private Sub mnuOptNslookupIP_Click()
    MsgBox "n.y.i. todo: mnuOptNslookupIP_Click"
End Sub
Private Sub mnuOptNslookupIPport_Click()
    MsgBox "n.y.i. todo: mnuOptNslookupIPport_Click"
End Sub

Private Sub mnuOptionsUserPCName_Click()
Try: On Error GoTo Catch
    'den Usernamen und den PC-Namen herausfinden und ins Netzwerk in eine Datei schreiben bzw anhängen
    Dim IP As IpAddress: Set IP = MSocket.GetMyIP
    Dim un As String:          un = MApp.UserName
    Dim cn As String:          cn = MApp.ComputerName
    Dim hn As String:          hn = MSocket.MyHostName
    'MsgBox "IP-Adress:    " & ip.ToStr & vbCrLf & _
    '       "Userrname:    " & un & vbCrLf & _
    '       "Computername: " & cn & vbCrLf & _
    '       "Hostname:     " & hn
    Dim pfn As PathFileName: Set pfn = MNew.PathFileName("C:\Install\IPScanner\IPScanner.bin")
    If Not pfn.PathExists Then
        If MsgBox("Path does not exist, Create Path?" & vbCrLf & pfn.Path, vbOKCancel) = vbCancel Then Exit Sub
        If Not pfn.PathCreate Then
            MsgBox "Could not create path: " & vbCrLf & pfn.Path
            Exit Sub
        End If
    End If
    Dim fc As String
    fc = "IP-Address: " & IP.ToStr & "; Username: " & un & "; Computername: " & cn & vbCrLf & _
         "Also write to file?"
    If MsgBox(fc, vbOKCancel) = vbCancel Then Exit Sub
    
    fc = "IP-Address: " & vbTab & IP.ToStr & vbTab & "; Username: " & vbTab & un & vbTab & "; Computername: " & vbTab & cn
    
    pfn.OpenFile FileMode_Append, FileAccess_Write
    pfn.WriteStr fc
    
'z.B. so:
'* open Command prompt
'* get the machine name with: nbtstat --a 192.168.2.77
'* get the user    name with: net view /domain:xxxxx.de > c:\ip\ip.txt
    
    GoTo Finally
Catch:
    ErrHandler "mnuOptionsUserPCName"
Finally:
    If Not pfn Is Nothing Then pfn.CloseFile
End Sub

Private Sub mnuOptionNetViewDomain_Click()
    'https://itstillworks.com/user-name-ip-address-6909133.html
    'nbtstat --a ip
    'net view /domain:ad > c:\ip\ip.txt
    
    'oder Tipp von hsachse:
    'wmic /node:10.0.0.112 computersystem get username /value
    'oder Tipp0479
    'http://www.activevb.de/tipps/vb6tipps/tipp0479.html
End Sub

' ^ ############################## ^ '  Menu mnuOpt   ' ^ ############################## ^ '

' v ############################## v '  Menu mnuInfo  ' v ############################## v '

Private Sub mnuHelpExternalIP_Click()
    Dim ipv4 As IpAddress: Set ipv4 = MNew.IpAddress(MApp.GetMyIP)
    Dim s As String: s = InputBox("My external IP:", , ipv4.ToStr)
End Sub

Private Sub mnuHelpInfo_Click()
    'MsgBox "MBO-Ing.com IPPingScanner " vbcrlf & App.CompanyName & " " & App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
    MsgBox App.CompanyName & " " & App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
' ^ ############################## ^ '  Menu mnuInfo  ' ^ ############################## ^ '

Private Sub TxtIPBase_LostFocus()
    Dim NewBaseIP As IpAddress: Set NewBaseIP = MNew.IpAddress(TxtIPBase.Text)
    TxtIPBase.Text = NewBaseIP.IPToStr
    Set MApp.Doc.IPBase = NewBaseIP
    MApp.Doc.StartIPb4 = MApp.Doc.IPBase.b1
    Set MApp.Doc.LastIP = MApp.Doc.IPBase.Clone
End Sub

Private Sub BtnScanNextXXIPs_Click()
    'MApp.DataChanged = True
    LblDTime.Caption = "Scanning..."
    DoEvents
    Dim dt As Single: dt = Timer
    Dim c As Long: c = MApp.Doc.SearchNIPs - 1
    
    If mnuOptDllOrNslookupExe.Checked Then
        IPPingScanner.Scan MApp.Doc.LastIP, MApp.Doc.StartIPb4, MApp.Doc.StartIPb4 + c
    Else
        IPPingScanner.Scan2 MApp.Doc.LastIP, MApp.Doc.StartIPb4, MApp.Doc.StartIPb4 + c
    End If
    
    MApp.Doc.StartIPb4 = MApp.Doc.StartIPb4 + c + 1
    MApp.Doc.LastIP.Add MApp.Doc.SearchNIPs
    dt = Timer - dt
    LblDTime.Caption = "Ready! time: " & Format(dt, "#.00") & " sec"
End Sub

Private Sub IPPingScanner_FoundIP(aIPV4 As IpAddress, out_Cancel As Boolean)
    If Not MApp.Doc.IPAddresses.Contains(aIPV4.IPToStr) Then
        MApp.Doc.IPAddresses.Add aIPV4
    Else
        Dim mbr As VbMsgBoxResult
        mbr = MsgBox("Already exists: " & aIPV4.ToStr, vbOKCancel)
        out_Cancel = mbr = vbCancel
    End If
    If aIPV4.IsValid Then
        List1.AddItem ""
        List2.AddItem aIPV4.IPToStr
    Else
        List1.AddItem aIPV4.IPToStr
    End If
    DoEvents
End Sub

Private Sub List1_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then Exit Sub
    Dim aIPV4 As IpAddress
    If MApp.Doc.IPAddresses.Contains(List1.List(i)) Then
        Set aIPV4 = MApp.Doc.IPAddresses.Item(List1.List(i))
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub
Private Sub List1_DblClick()
    LB_DblClick List1
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim btn As MouseButtonConstants: btn = Button
    If btn = vbRightButton Then
        Set mnuPopListBox = List1
        PopupMenu mnuPopup
    End If
End Sub

Private Sub List2_Click()
    Dim i As Long: i = List2.ListIndex
    Dim s As String: s = List2.List(i)
    If i < 0 Then Exit Sub
    Dim aIPV4 As IpAddress
    If MApp.Doc.IPAddresses.Contains(s) Then
        Set aIPV4 = MApp.Doc.IPAddresses.Item(s)
        Debug.Print aIPV4.IPToStr
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub

Private Sub List2_DblClick()
    LB_DblClick List2
End Sub
Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim btn As MouseButtonConstants: btn = Button
    If btn = vbRightButton Then
        Set mnuPopListBox = List2
        PopupMenu mnuPopup
    End If
End Sub

Private Sub LB_DblClick(aLB As ListBox)
    Dim i As Long: i = aLB.ListIndex
    If i < 0 Then Exit Sub
    Dim s As String: s = aLB.List(i)
    Dim aIPV4 As IpAddress
    If MApp.Doc.IPAddresses.Contains(s) Then
        Set aIPV4 = MApp.Doc.IPAddresses.Item(s)
        If mnuOptDllOrNslookupDll.Checked Then
            aIPV4.CallNslookup
        End If
        aIPV4.CallPing
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub

'Private Sub UpdateView(aDoc As Document)

    'TxtIPBase.Text = aDoc.

'End Sub


''copy this same function to every class, form or module
''the name of the class or form will be added automatically
''in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
'' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional WinApiError, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKCancel, _
                            Optional bRetry As Boolean) As VbMsgBoxResult

    If bRetry Then

        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)

    Else

        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)

    End If

End Function


